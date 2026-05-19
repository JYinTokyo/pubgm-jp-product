#!/usr/bin/env python3
"""
pubgm-jp-product 대시보드 생성기
- spin_kpi_stat / crate_kpi_info 기반 상품별 KPI 대시보드
- 판매 중인 상품: 매일 재쿼리
- 완료된 상품: product_cache.json 에 저장, 재쿼리 안 함
- 과금 레벨 기준: 판매기간 30일 정규화 / 직전 3개월 / 직전 1개월
"""
import json, os, sys, ssl, time, urllib.request, urllib.error, http.client
from datetime import datetime, date, timedelta

DATABRICKS_HOST = "https://krafton-hq.cloud.databricks.com"
WAREHOUSE_ID    = "e87bfc435cb9ad4e"
TODAY           = date.today().isoformat()          # YYYY-MM-DD
DATA_CUTOFF     = (date.today() - timedelta(days=2)).isoformat()  # 데이터는 2일 전까지
CACHE_FILE      = os.path.join(os.path.dirname(__file__), 'product_cache.json')
CACHE_VER       = 1
SPIN_XLSX       = os.path.join(os.path.dirname(__file__), 'Spin.xlsx')
# pubgm-jp-item 프로젝트의 Spin.xlsx 참조 (없으면 그쪽 것 사용)
if not os.path.exists(SPIN_XLSX):
    SPIN_XLSX = os.path.join(os.path.dirname(__file__), '../pubgm-jp-item/Spin.xlsx')

def get_token():
    env_file = os.path.join(os.path.dirname(__file__), '.env')
    if os.path.exists(env_file):
        with open(env_file) as f:
            for line in f:
                line = line.strip()
                if line.startswith('DATABRICKS_TOKEN='):
                    return line.split('=', 1)[1].strip().strip('"\'')
    return os.environ.get('DATABRICKS_TOKEN', '')

CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE

def api(token, path, method='GET', body=None):
    req = urllib.request.Request(
        f'{DATABRICKS_HOST}/api/2.0/sql/{path}',
        data=json.dumps(body).encode() if body else None,
        headers={'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'},
        method=method
    )
    with urllib.request.urlopen(req, context=CTX) as r:
        return json.loads(r.read())

def run_query(token, sql, label=''):
    if label:
        print(f"  [{label}]...", flush=True)
    result = api(token, 'statements', 'POST', {
        'statement': sql, 'warehouse_id': WAREHOUSE_ID, 'wait_timeout': '0s'
    })
    stmt_id = result['statement_id']
    poll_errors = 0
    while result.get('status', {}).get('state') in ('PENDING', 'RUNNING'):
        time.sleep(5)
        try:
            result = api(token, f'statements/{stmt_id}')
            poll_errors = 0
        except Exception as e:
            poll_errors += 1
            if poll_errors > 10:
                raise
            time.sleep(15)
    state = result.get('status', {}).get('state')
    if state != 'SUCCEEDED':
        raise RuntimeError(f"Query failed [{label}]: {result.get('status')}")
    cols = [c['name'] for c in result['manifest']['schema']['columns']]
    return [dict(zip(cols, row)) for row in result.get('result', {}).get('data_array', [])]

# ── VTuber 상품 목록 (카테고리 필터용) ─────────────────────────────────────────
VTUBER_SPIN_IDS = {230073002, 230073003, 234073031, 235073032, 236073031,
                   237073031, 238073031, 238073032, 240073032}
VTUBER_BOX_NAMES = {'Salome 宝箱_JP', 'UC_Change_reason3029', 'Nakiri Ayame 宝箱'}

# ── 과금 레벨 분류 CASE 식 ───────────────────────────────────────────────────
LEVEL_CASE = """CASE
  WHEN avg_uc = 0               THEN 'Non-Paid'
  WHEN avg_uc <= 1200           THEN 'Lv_1'
  WHEN avg_uc <= 3000           THEN 'Lv_2'
  WHEN avg_uc <= 6000           THEN 'Lv_3'
  WHEN avg_uc <= 18000          THEN 'Lv_4'
  WHEN avg_uc <= 30000          THEN 'Lv_5'
  WHEN avg_uc <= 60000          THEN 'Lv_6'
  ELSE                               'Lv_7'
END"""

LEVELS = ['Non-Paid', 'Lv_1', 'Lv_2', 'Lv_3', 'Lv_4', 'Lv_5', 'Lv_6', 'Lv_7']

# ── SQL: 전체 스핀 기본 통계 (spin_kpi_stat) ────────────────────────────────
SQL_SPIN_ALL = f"""
SELECT
  CAST(spin_contents AS STRING) AS spin_id,
  spin_type,
  CAST(launch_spin_date AS STRING) AS launch,
  CAST(DATE_ADD(launch_spin_date, CAST(duration AS INT)) AS STRING) AS end_date,
  CAST(duration AS INT) AS duration,
  CAST(total_user   AS BIGINT) AS total_user,
  CAST(total_amount AS BIGINT) AS total_uc,
  CAST(nbu_user     AS BIGINT) AS npu_user,
  CAST(nbu_amount   AS BIGINT) AS npu_uc,
  CAST(return_user  AS BIGINT) AS ret_user,
  CAST(return_amount AS BIGINT) AS ret_uc
FROM pubgm_mart.spin_kpi_stat
WHERE country = 'JP'
  AND launch_spin_date >= ADD_MONTHS(CURRENT_DATE(), -25)
QUALIFY ROW_NUMBER() OVER (PARTITION BY spin_contents ORDER BY std_dt DESC) = 1
ORDER BY launch_spin_date DESC
"""

# ── SQL: 전체 상자 기본 통계 (crate_kpi_info) ──────────────────────────────
SQL_BOX_ALL = f"""
SELECT
  crate AS box_id,
  CAST(launch_chest_date AS STRING) AS launch,
  CAST(DATE_ADD(launch_chest_date, CAST(duration AS INT)) AS STRING) AS end_date,
  CAST(duration AS INT) AS duration,
  CAST(total_user   AS BIGINT) AS total_user,
  CAST(total_amount AS BIGINT) AS total_uc,
  CAST(nbu_user     AS BIGINT) AS npu_user,
  CAST(nbu_amount   AS BIGINT) AS npu_uc,
  CAST(return_user  AS BIGINT) AS ret_user,
  CAST(return_amount AS BIGINT) AS ret_uc
FROM pubgm_mart.crate_kpi_info
WHERE country = 'JP'
  AND crate NOT LIKE '%KR'
  AND launch_chest_date >= ADD_MONTHS(CURRENT_DATE(), -25)
QUALIFY ROW_NUMBER() OVER (PARTITION BY crate ORDER BY std_dt DESC) = 1
ORDER BY launch_chest_date DESC
"""

def make_spin_level_sql(spin_ids, cutoff):
    """지정된 spin_ids에 대해 판매기간/3m/1m 과금 레벨 한번에 계산"""
    ids_str = ','.join(str(i) for i in spin_ids)
    return f"""
WITH spin_meta AS (
  SELECT
    CAST(spin_contents AS STRING) AS spin_id,
    launch_spin_date,
    LEAST(DATE_ADD(launch_spin_date, CAST(duration AS INT)), DATE('{cutoff}')) AS eff_end,
    DATEDIFF(
      LEAST(DATE_ADD(launch_spin_date, CAST(duration AS INT)), DATE('{cutoff}')),
      launch_spin_date
    ) + 1 AS period_days,
    DATE_TRUNC('month', ADD_MONTHS(launch_spin_date, -3)) AS c3m_start,
    LAST_DAY(ADD_MONTHS(launch_spin_date, -1))            AS prev1m_end,
    DATE_TRUNC('month', ADD_MONTHS(launch_spin_date, -1)) AS prev1m_start
  FROM pubgm_mart.spin_kpi_stat
  WHERE country = 'JP'
    AND spin_contents IN ({ids_str})
  QUALIFY ROW_NUMBER() OVER (PARTITION BY spin_contents ORDER BY std_dt DESC) = 1
),
spin_buyers AS (
  SELECT sm.spin_id, CAST(s.vopenid AS STRING) AS vopenid,
    ABS(SUM(COALESCE(CAST(s.realchg AS DOUBLE), 0))) AS uc_used
  FROM pubgm_mart.user_uc_chg_subreason_info s
  JOIN spin_meta sm ON CAST(s.subreason AS BIGINT) = CAST(sm.spin_id AS BIGINT)
    AND s.std_dt >= sm.launch_spin_date AND s.std_dt <= sm.eff_end
  WHERE s.addorreduce = 0
  GROUP BY sm.spin_id, s.vopenid
),
charge_data AS (
  SELECT sb.spin_id, CAST(u.vopenid AS STRING) AS vopenid,
    SUM(CASE WHEN u.date >= sm.launch_spin_date AND u.date <= sm.eff_end
        THEN COALESCE(u.PaidChgAmount, 0) ELSE 0 END) AS period_charge,
    sm.period_days,
    SUM(CASE WHEN u.date >= sm.c3m_start AND u.date <= sm.prev1m_end
        THEN COALESCE(u.PaidChgAmount, 0) ELSE 0 END) AS charge_3m,
    SUM(CASE WHEN u.date >= sm.prev1m_start AND u.date <= sm.prev1m_end
        THEN COALESCE(u.PaidChgAmount, 0) ELSE 0 END) AS charge_1m
  FROM spin_buyers sb
  JOIN spin_meta sm ON sb.spin_id = sm.spin_id
  LEFT JOIN pubgm_mart.ucpurchase_individual u ON CAST(u.vopenid AS STRING) = sb.vopenid
    AND u.date >= sm.c3m_start AND u.date <= sm.eff_end
  GROUP BY sb.spin_id, u.vopenid, sm.period_days, sm.launch_spin_date, sm.eff_end,
           sm.c3m_start, sm.prev1m_end, sm.prev1m_start
),
classified AS (
  SELECT sb.spin_id, sb.vopenid, sb.uc_used,
    {LEVEL_CASE.replace('avg_uc', 'COALESCE(cd.period_charge,0)/NULLIF(cd.period_days,0)*30')} AS lv_period,
    {LEVEL_CASE.replace('avg_uc', 'COALESCE(cd.charge_3m,0)/3.0')} AS lv_3m,
    {LEVEL_CASE.replace('avg_uc', 'COALESCE(cd.charge_1m,0)')} AS lv_1m
  FROM spin_buyers sb
  LEFT JOIN charge_data cd ON sb.spin_id = cd.spin_id AND sb.vopenid = cd.vopenid
)
SELECT spin_id, method, pay_level,
  COUNT(DISTINCT vopenid) AS users,
  CAST(SUM(uc_used) AS BIGINT) AS total_uc
FROM (
  SELECT spin_id, vopenid, uc_used, 'period' AS method, lv_period AS pay_level FROM classified
  UNION ALL
  SELECT spin_id, vopenid, uc_used, '3m',    lv_3m   FROM classified
  UNION ALL
  SELECT spin_id, vopenid, uc_used, '1m',    lv_1m   FROM classified
)
GROUP BY spin_id, method, pay_level
ORDER BY spin_id, method, pay_level
"""

def make_box_level_sql(box_ids, cutoff):
    """지정된 box_ids(crate 이름)에 대해 판매기간/3m/1m 과금 레벨 계산"""
    ids_str = ','.join(f"'{b}'" for b in box_ids)
    return f"""
WITH box_meta AS (
  SELECT
    crate AS box_id,
    launch_chest_date,
    LEAST(DATE_ADD(launch_chest_date, CAST(duration AS INT)), DATE('{cutoff}')) AS eff_end,
    DATEDIFF(
      LEAST(DATE_ADD(launch_chest_date, CAST(duration AS INT)), DATE('{cutoff}')),
      launch_chest_date
    ) + 1 AS period_days,
    DATE_TRUNC('month', ADD_MONTHS(launch_chest_date, -3)) AS c3m_start,
    LAST_DAY(ADD_MONTHS(launch_chest_date, -1))            AS prev1m_end,
    DATE_TRUNC('month', ADD_MONTHS(launch_chest_date, -1)) AS prev1m_start
  FROM pubgm_mart.crate_kpi_info
  WHERE country = 'JP'
    AND crate IN ({ids_str})
  QUALIFY ROW_NUMBER() OVER (PARTITION BY crate ORDER BY std_dt DESC) = 1
),
box_buyers AS (
  SELECT bm.box_id, CAST(u.vopenid AS STRING) AS vopenid,
    SUM(CAST(u.amount AS BIGINT)) AS uc_used
  FROM pubgm_mart.user_use_uc_info u
  JOIN box_meta bm ON u.product = bm.box_id
    AND u.std_dt >= bm.launch_chest_date AND u.std_dt <= bm.eff_end
  WHERE u.country = 'JP'
  GROUP BY bm.box_id, u.vopenid
),
charge_data AS (
  SELECT bb.box_id, CAST(u.vopenid AS STRING) AS vopenid,
    SUM(CASE WHEN u.date >= bm.launch_chest_date AND u.date <= bm.eff_end
        THEN COALESCE(u.PaidChgAmount, 0) ELSE 0 END) AS period_charge,
    bm.period_days,
    SUM(CASE WHEN u.date >= bm.c3m_start AND u.date <= bm.prev1m_end
        THEN COALESCE(u.PaidChgAmount, 0) ELSE 0 END) AS charge_3m,
    SUM(CASE WHEN u.date >= bm.prev1m_start AND u.date <= bm.prev1m_end
        THEN COALESCE(u.PaidChgAmount, 0) ELSE 0 END) AS charge_1m
  FROM box_buyers bb
  JOIN box_meta bm ON bb.box_id = bm.box_id
  LEFT JOIN pubgm_mart.ucpurchase_individual u ON CAST(u.vopenid AS STRING) = bb.vopenid
    AND u.date >= bm.c3m_start AND u.date <= bm.eff_end
  GROUP BY bb.box_id, u.vopenid, bm.period_days, bm.launch_chest_date, bm.eff_end,
           bm.c3m_start, bm.prev1m_end, bm.prev1m_start
),
classified AS (
  SELECT bb.box_id, bb.vopenid, bb.uc_used,
    {LEVEL_CASE.replace('avg_uc', 'COALESCE(cd.period_charge,0)/NULLIF(cd.period_days,0)*30')} AS lv_period,
    {LEVEL_CASE.replace('avg_uc', 'COALESCE(cd.charge_3m,0)/3.0')} AS lv_3m,
    {LEVEL_CASE.replace('avg_uc', 'COALESCE(cd.charge_1m,0)')} AS lv_1m
  FROM box_buyers bb
  LEFT JOIN charge_data cd ON bb.box_id = cd.box_id AND bb.vopenid = cd.vopenid
)
SELECT box_id, method, pay_level,
  COUNT(DISTINCT vopenid) AS users,
  CAST(SUM(uc_used) AS BIGINT) AS total_uc
FROM (
  SELECT box_id, vopenid, uc_used, 'period' AS method, lv_period AS pay_level FROM classified
  UNION ALL
  SELECT box_id, vopenid, uc_used, '3m',    lv_3m   FROM classified
  UNION ALL
  SELECT box_id, vopenid, uc_used, '1m',    lv_1m   FROM classified
)
GROUP BY box_id, method, pay_level
ORDER BY box_id, method, pay_level
"""

def load_spin_names():
    """Spin.xlsx에서 ID→명칭 매핑 로드"""
    names = {}
    if not os.path.exists(SPIN_XLSX):
        print(f"Warning: Spin.xlsx not found at {SPIN_XLSX}")
        return names
    try:
        import openpyxl
        wb = openpyxl.load_workbook(SPIN_XLSX, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[1] and row[2]:
                names[str(int(row[1]))] = str(row[2])
        wb.close()
        print(f"Spin names loaded: {len(names)}")
    except Exception as e:
        print(f"Warning: Spin.xlsx load failed: {e}")
    return names

def load_cache():
    if not os.path.exists(CACHE_FILE):
        return {'ver': CACHE_VER, 'spins': {}, 'boxes': {}}
    with open(CACHE_FILE) as f:
        data = json.load(f)
    if data.get('ver') != CACHE_VER:
        print("Cache version mismatch — rebuilding")
        return {'ver': CACHE_VER, 'spins': {}, 'boxes': {}}
    return data

def save_cache(cache):
    with open(CACHE_FILE, 'w') as f:
        json.dump(cache, f, ensure_ascii=False)

def build_level_map(rows, id_key='spin_id'):
    """쿼리 결과를 {id: {method: {pay_level: {users, uc}}}} 로 변환"""
    result = {}
    for r in rows:
        pid = r[id_key]
        method = r['method']
        lv = r['pay_level']
        if pid not in result:
            result[pid] = {'period': {}, '3m': {}, '1m': {}}
        result[pid][method][lv] = {
            'u': int(r['users'] or 0),
            'uc': int(r['total_uc'] or 0)
        }
    return result

def is_active(end_date_str):
    """end_date가 오늘 이후이면 판매 중"""
    try:
        return end_date_str >= TODAY
    except:
        return False

def main():
    token = get_token()
    if not token:
        print("ERROR: DATABRICKS_TOKEN not found"); sys.exit(1)

    spin_names = load_spin_names()
    cache = load_cache()

    print("Querying spin basic stats...")
    raw_spins = run_query(token, SQL_SPIN_ALL, 'spin_kpi_stat')
    print("Querying box basic stats...")
    raw_boxes = run_query(token, SQL_BOX_ALL, 'crate_kpi_info')

    print(f"Spins: {len(raw_spins)}, Boxes: {len(raw_boxes)}")

    # 판매 중 vs 완료 분리
    active_spin_ids = [r['spin_id'] for r in raw_spins if is_active(r['end_date'])]
    done_spin_ids   = [r['spin_id'] for r in raw_spins
                       if not is_active(r['end_date']) and r['spin_id'] not in cache['spins']]
    active_box_ids  = [r['box_id'] for r in raw_boxes if is_active(r['end_date'])]
    done_box_ids    = [r['box_id'] for r in raw_boxes
                       if not is_active(r['end_date']) and r['box_id'] not in cache['boxes']]

    print(f"Active spins: {len(active_spin_ids)}, New completed spins: {len(done_spin_ids)}")
    print(f"Active boxes: {len(active_box_ids)}, New completed boxes: {len(done_box_ids)}")

    spin_level = {}
    box_level = {}

    # 판매 중 상품: 오늘 기준으로 재쿼리
    query_spin_ids = active_spin_ids + done_spin_ids
    if query_spin_ids:
        print(f"Querying spin pay level ({len(query_spin_ids)} products)...")
        rows = run_query(token, make_spin_level_sql(query_spin_ids, DATA_CUTOFF), 'spin_level')
        spin_level = build_level_map(rows, 'spin_id')

    query_box_ids = active_box_ids + done_box_ids
    if query_box_ids:
        print(f"Querying box pay level ({len(query_box_ids)} products)...")
        rows = run_query(token, make_box_level_sql(query_box_ids, DATA_CUTOFF), 'box_level')
        box_level = build_level_map(rows, 'box_id')

    # 완료된 상품 캐시에 저장
    for sid in done_spin_ids:
        if sid in spin_level:
            cache['spins'][sid] = spin_level[sid]
    for bid in done_box_ids:
        if bid in box_level:
            cache['boxes'][bid] = box_level[bid]
    save_cache(cache)

    # 캐시에서 완료 상품 레벨 데이터 로드
    for sid, lv in cache['spins'].items():
        if sid not in spin_level:
            spin_level[sid] = lv
    for bid, lv in cache['boxes'].items():
        if bid not in box_level:
            box_level[bid] = lv

    # ── JSON 데이터 구축 ──────────────────────────────────────────────────────
    def make_basic(r, id_key):
        return {
            'id':       r[id_key],
            'launch':   r['launch'],
            'end':      r['end_date'],
            'days':     int(r['duration'] or 0),
            'active':   is_active(r['end_date']),
            'total_u':  int(r['total_user'] or 0),
            'total_uc': int(r['total_uc'] or 0),
            'npu_u':    int(r['npu_user'] or 0),
            'npu_uc':   int(r['npu_uc'] or 0),
            'ret_u':    int(r['ret_user'] or 0),
            'ret_uc':   int(r['ret_uc'] or 0),
        }

    spin_data = []
    for r in raw_spins:
        sid = r['spin_id']
        d = make_basic(r, 'spin_id')
        d['name'] = spin_names.get(sid, sid)  # 명칭 없으면 ID 그대로
        d['spin_type'] = r.get('spin_type', '')
        d['vtuber'] = int(sid) in VTUBER_SPIN_IDS if sid.isdigit() else False
        d['level'] = spin_level.get(sid, {})
        spin_data.append(d)

    box_data = []
    for r in raw_boxes:
        bid = r['box_id']
        d = make_basic(r, 'box_id')
        d['vtuber'] = bid in VTUBER_BOX_NAMES
        d['level'] = box_level.get(bid, {})
        box_data.append(d)

    gen_date = datetime.now().strftime('%Y-%m-%d %H:%M')
    tpl = os.path.join(os.path.dirname(__file__), 'template.html')
    dist_dir = os.path.join(os.path.dirname(__file__), 'dist')
    os.makedirs(dist_dir, exist_ok=True)
    out = os.path.join(dist_dir, 'index.html')

    with open(tpl) as f:
        html = f.read()

    html = html.replace('%%SPIN_DATA%%', json.dumps(spin_data, ensure_ascii=False))
    html = html.replace('%%BOX_DATA%%',  json.dumps(box_data,  ensure_ascii=False))
    html = html.replace('%%GEN_DATE%%',  gen_date)
    html = html.replace('%%CUTOFF%%',    DATA_CUTOFF)

    with open(out, 'w') as f:
        f.write(html)

    print(f"Generated {out} (spins:{len(spin_data)}, boxes:{len(box_data)})")

if __name__ == '__main__':
    main()
