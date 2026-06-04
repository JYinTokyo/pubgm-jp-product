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
WAREHOUSE_ID    = "e87bfc435cb9ad4e"  # external warehouse (shared 불안정 시 대체)
TODAY           = date.today().isoformat()          # YYYY-MM-DD
DATA_CUTOFF     = (date.today() - timedelta(days=2)).isoformat()  # 데이터는 2일 전까지
CACHE_FILE      = os.path.join(os.path.dirname(__file__), 'product_cache.json')
CACHE_VER       = 8
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
    attempts = 1 if method == 'POST' else 5
    for attempt in range(attempts):
        try:
            with urllib.request.urlopen(req, context=CTX, timeout=60) as r:
                return json.loads(r.read())
        except (http.client.IncompleteRead, urllib.error.URLError, TimeoutError, OSError) as e:
            if attempt == attempts - 1:
                raise
            time.sleep(15 * (attempt + 1))

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
    rows = list(result.get('result', {}).get('data_array', []))
    # 페이지네이션: next_chunk_index 있으면 추가 청크 fetch
    next_chunk = result.get('result', {}).get('next_chunk_index')
    while next_chunk is not None:
        chunk = api(token, f'statements/{stmt_id}/result/chunks/{next_chunk}')
        rows.extend(chunk.get('data_array', []))
        next_chunk = chunk.get('next_chunk_index')
    return [dict(zip(cols, row)) for row in rows]

# ── VTuber 상품 목록 (카테고리 필터용) ─────────────────────────────────────────
VTUBER_SPIN_IDS = {230073002, 230073003, 234073031, 235073032, 236073031,
                   237073031, 238073031, 238073032, 240073032}
VTUBER_BOX_NAMES = {'Salome 宝箱_JP', 'UC_Change_reason3029', 'Nakiri Ayame 宝箱',
                    '440 Usada Pekora CRATE_JP'}

# ── アニメ 상품 목록 ─────────────────────────────────────────────────────────
ANIME_SPIN_IDS = {238073024, 238073025, 240073031, 241073031, 241073033}
ANIME_BOX_NAMES = {
    'SAKAMOTO DAYS 宝箱', 'Kaiju No. 8 CRATE', 'Levi Crate', 'Tensura Crate',
    'Frieren:Beyond Journey\'s End宝箱_JP',
    'Code Geass 1 宝箱_JP', 'Code Geass 2 宝箱_JP',
    '龙珠宝箱1', '龙珠宝箱2', '龙珠宝箱320JP',
    'SPY XFAMILY Event Crate_JP',
}

# ── MUMMY 상품 목록 ──────────────────────────────────────────────────────────
MUMMY_BOX_NAMES = {'MUMMY宝箱_JP', '3.9冰木乃伊宝箱', '烈焰古神宝箱(3.2木乃伊宝箱)'}

# ── 표시 명칭 오버라이드 ────────────────────────────────────────────────────
SPIN_NAME_OVERRIDES = {
    '230073002': 'Usada Pekora 1-1',
    '230073003': 'Usada Pekora 1-2',
    '234073031': 'Pmarusama 1',
    '235073032': 'Nijisanji',
    '238073031': 'Pmarusama 2',
    '238073032': 'Sakura Miko',
    '240073032': 'Houshou Marine',
    '238073024': 'Attack on Titan-Eren',
    '238073025': 'Attack on Titan-Armin',
    '240073031': 'TokyoRevengers',
    '241073031': 'DAN DA DAN',
    '241073033': 'Kaiju No.8 Spin',
}
BOX_NAME_OVERRIDES = {
    '440 Usada Pekora CRATE_JP':              'Usada Pekora 2',
    'Nakiri Ayame 宝箱':                       'Nakiri Ayame',
    'Salome 宝箱_JP':                          'Salome',
    'UC_Change_reason3029':                   'Kanae & Kuzuha',
    'Frieren:Beyond Journey\'s End宝箱_JP':    'Frieren',
    'SAKAMOTO DAYS 宝箱':                      'SAKAMOTO DAYS',
    'Kaiju No. 8 CRATE':                      'Kaiju No.8 Crate',
    'Levi Crate':                             'Attack on Titan-Levi',
    'Tensura Crate':                          'Tensura',
    'Code Geass 1 宝箱_JP':                   'Code Geass 1',
    'Code Geass 2 宝箱_JP':                   'Code Geass 2',
    'SPY XFAMILY Event Crate_JP':             'SPY XFAMILY',
    'MUMMY宝箱_JP':                            'MUMMY',
}

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
# launch_spin_date: D+61 이후 파이프라인이 날짜를 밀어서 기록하는 아티팩트 존재
# → n_per_launch DESC, launch_spin_date ASC 로 가장 많이 기록된 (= shift 전) 날짜 선택
# end_date: MAX(std_dt) = 실제 마지막 데이터 날짜
SQL_SPIN_ALL = f"""
SELECT
  CAST(spin_contents AS STRING) AS spin_id,
  spin_type,
  CAST(actual_launch AS STRING) AS launch,
  CAST(max_std_dt AS STRING) AS end_date,
  CAST(max_std_dt AS STRING) AS last_std_dt,
  DATEDIFF(DATE(max_std_dt), DATE(actual_launch)) AS duration,
  CAST(total_user   AS BIGINT) AS total_user,
  CAST(total_amount AS BIGINT) AS total_uc,
  CAST(nbu_user     AS BIGINT) AS npu_user,
  CAST(nbu_amount   AS BIGINT) AS npu_uc,
  CAST(return_user  AS BIGINT) AS ret_user,
  CAST(return_amount AS BIGINT) AS ret_uc
FROM (
  SELECT *,
    MIN(std_dt) OVER (PARTITION BY spin_contents) AS actual_launch,
    MAX(std_dt) OVER (PARTITION BY spin_contents) AS max_std_dt
  FROM pubgm_mart.spin_kpi_stat
  WHERE country = 'JP'
    AND launch_spin_date >= ADD_MONTHS(CURRENT_DATE(), -36)
)
QUALIFY ROW_NUMBER() OVER (PARTITION BY spin_contents ORDER BY std_dt DESC) = 1
ORDER BY actual_launch DESC
"""

# ── SQL: 전체 상자 기본 통계 (crate_kpi_info) ──────────────────────────────
# n_per_launch = (crate, launch_chest_date) 그룹별 row 수
# 가장 row 수가 많은 launch_chest_date = 실제 판매 시작일 (D+61부터 파이프라인이 날짜를 밀어서 기록하는 아티팩트 제거)
SQL_BOX_ALL = f"""
SELECT
  crate AS box_id,
  CAST(launch_chest_date AS STRING) AS launch,
  CAST(max_std_dt AS STRING) AS end_date,
  CAST(std_dt AS STRING) AS last_std_dt,
  DATEDIFF(DATE(max_std_dt), launch_chest_date) AS duration,
  CAST(total_user   AS BIGINT) AS total_user,
  CAST(total_amount AS BIGINT) AS total_uc,
  CAST(nbu_user     AS BIGINT) AS npu_user,
  CAST(nbu_amount   AS BIGINT) AS npu_uc,
  CAST(return_user  AS BIGINT) AS ret_user,
  CAST(return_amount AS BIGINT) AS ret_uc
FROM (
  SELECT *,
    COUNT(*) OVER (PARTITION BY crate, launch_chest_date) AS n_per_launch,
    MAX(std_dt) OVER (PARTITION BY crate) AS max_std_dt
  FROM pubgm_mart.crate_kpi_info
  WHERE country = 'JP'
    AND crate NOT LIKE '%KR'
    AND launch_chest_date >= ADD_MONTHS(CURRENT_DATE(), -36)
)
QUALIFY ROW_NUMBER() OVER (
  PARTITION BY crate
  ORDER BY n_per_launch DESC, launch_chest_date ASC, std_dt DESC
) = 1
ORDER BY launch_chest_date DESC
"""

def _dn_buyer_cols():
    """buyer_dn CTE: uc_dn0~uc_dn60 + uc_total 컬럼 생성"""
    cols = ['    SUM(uc_used) AS uc_total']
    for n in range(61):
        cols.append(f'    SUM(CASE WHEN day_n <= {n} THEN uc_used ELSE 0 END) AS uc_dn{n}')
    return ',\n'.join(cols)

def _dn_pass_cols():
    """classified/combos에서 uc 컬럼 목록 (uc_total, uc_dn0..60)"""
    return 'uc_total, ' + ', '.join(f'uc_dn{n}' for n in range(61))

def _dn_union_all(id_col):
    """최종 UNION ALL: total + D+0~D+60"""
    parts = [
        f"SELECT {id_col}, method, pay_level, vopenid, 'total' AS d_preset, uc_total AS uc_val FROM combos"
    ]
    for n in range(61):
        parts.append(
            f"SELECT {id_col}, method, pay_level, vopenid, '{n}' AS d_preset, uc_dn{n} FROM combos WHERE first_day <= {n}"
        )
    return '\n  UNION ALL\n  '.join(parts)

def make_spin_level_sql(spin_ids, cutoff, snap_dates=None, min_launch=None, max_eff_end=None):
    """판매기간/3m/1m × Total/D+0~D+60 집계 — (spin_id, method, d_preset, pay_level, users, total_uc) 반환
    레벨: purchasing_power_user 스냅샷 기반
    snap_dates: 리터럴 날짜 집합 (Delta Lake 파티션 프루닝 최적화)
    min_launch: 배치 내 최소 launch_spin_date — 파티션 프루닝 하한선
    max_eff_end: 배치 내 최대 eff_end — 파티션 프루닝 상한선 (종료 상품은 cutoff보다 훨씬 이전)
    """
    ids_str    = ','.join(str(i) for i in spin_ids)
    min_launch = min_launch or '2020-01-01'
    max_eff_end = max_eff_end or cutoff
    dn_cols  = _dn_buyer_cols()
    dn_pass  = _dn_pass_cols()
    dn_union = _dn_union_all('spin_id')
    if snap_dates:
        date_lit = ', '.join(f"DATE('{d}')" for d in sorted(snap_dates))
        lv_date_filter = f"p.std_dt IN ({date_lit})"
    else:
        lv_date_filter = """p.std_dt IN (
      SELECT DISTINCT launch_spin_date FROM spin_meta UNION ALL
      SELECT DISTINCT prev3m_end       FROM spin_meta UNION ALL
      SELECT DISTINCT prev1m_end       FROM spin_meta
    )"""
    return f"""
WITH spin_meta AS (
  SELECT
    CAST(spin_contents AS STRING) AS spin_id,
    DATE(actual_launch) AS launch_spin_date,
    LEAST(DATE(max_std_dt), DATE('{cutoff}')) AS eff_end,
    LAST_DAY(ADD_MONTHS(DATE(actual_launch), -3)) AS prev3m_end,
    LAST_DAY(ADD_MONTHS(DATE(actual_launch), -1)) AS prev1m_end
  FROM (
    SELECT *,
      MIN(std_dt) OVER (PARTITION BY spin_contents) AS actual_launch,
      MAX(std_dt) OVER (PARTITION BY spin_contents) AS max_std_dt
    FROM pubgm_mart.spin_kpi_stat
    WHERE country = 'JP' AND spin_contents IN ({ids_str})
  )
  QUALIFY ROW_NUMBER() OVER (PARTITION BY spin_contents ORDER BY std_dt DESC) = 1
),
jp_users AS (
  SELECT DISTINCT CAST(vopenid AS STRING) AS vopenid
  FROM pubgm_mart.purchasing_power_user
  WHERE country = 'JP' AND std_dt = DATE('{cutoff}')
),
spin_txns AS (
  SELECT sm.spin_id, CAST(s.vopenid AS STRING) AS vopenid,
    DATEDIFF(s.std_dt, sm.launch_spin_date) AS day_n,
    ABS(COALESCE(CAST(s.realchg AS DOUBLE), 0)) AS uc_used
  FROM pubgm_mart.user_uc_chg_subreason_info s
  JOIN spin_meta sm ON CAST(s.subreason AS BIGINT) = CAST(sm.spin_id AS BIGINT)
    AND s.std_dt >= sm.launch_spin_date AND s.std_dt <= sm.eff_end
  JOIN jp_users ON CAST(s.vopenid AS STRING) = jp_users.vopenid
  WHERE s.addorreduce = 0
    AND s.std_dt >= DATE('{min_launch}')
    AND s.std_dt <= DATE('{max_eff_end}')
    AND CAST(s.subreason AS BIGINT) IN ({ids_str})
),
buyer_dn AS (
  SELECT spin_id, vopenid,
    MIN(day_n) AS first_day,
{dn_cols}
  FROM spin_txns
  GROUP BY spin_id, vopenid
),
buyers AS (
  SELECT DISTINCT vopenid FROM buyer_dn
),
lv_snap AS (
  SELECT CAST(p.vopenid AS STRING) AS vopenid, p.std_dt,
    COALESCE(p.pay_amt_duringLast30days_tag, 'Non-Paid') AS lv
  FROM pubgm_mart.purchasing_power_user p
  LEFT SEMI JOIN buyers b ON CAST(p.vopenid AS STRING) = b.vopenid
  WHERE p.country = 'JP'
    AND {lv_date_filter}
),
classified AS (
  SELECT bd.spin_id, bd.vopenid, bd.first_day,
    {dn_pass},
    COALESCE(lp.lv, 'Non-Paid') AS lv_period,
    COALESCE(l3.lv, 'Non-Paid') AS lv_3m,
    COALESCE(l1.lv, 'Non-Paid') AS lv_1m
  FROM buyer_dn bd
  JOIN spin_meta sm ON bd.spin_id = sm.spin_id
  LEFT JOIN lv_snap lp ON bd.vopenid = lp.vopenid AND lp.std_dt = sm.launch_spin_date
  LEFT JOIN lv_snap l3 ON bd.vopenid = l3.vopenid AND l3.std_dt = sm.prev3m_end
  LEFT JOIN lv_snap l1 ON bd.vopenid = l1.vopenid AND l1.std_dt = sm.prev1m_end
),
combos AS (
  SELECT spin_id, vopenid, first_day, 'period' AS method, lv_period AS pay_level, {dn_pass} FROM classified
  UNION ALL
  SELECT spin_id, vopenid, first_day, '3m',    lv_3m,     {dn_pass} FROM classified
  UNION ALL
  SELECT spin_id, vopenid, first_day, '1m',    lv_1m,     {dn_pass} FROM classified
)
SELECT spin_id, method, d_preset, pay_level,
  COUNT(*) AS users,
  CAST(SUM(uc_val) AS BIGINT) AS total_uc
FROM (
  {dn_union}
)
GROUP BY spin_id, method, d_preset, pay_level
ORDER BY spin_id, method, d_preset, pay_level
"""

def make_box_level_sql(box_ids, cutoff, snap_dates=None, min_launch=None, max_eff_end=None):
    """판매기간/3m/1m × Total/D+0~D+60 집계 — (box_id, method, d_preset, pay_level, users, total_uc) 반환
    최적화:
    - box_txns(user_use_uc_info) 3회 스캔 (기존 63회 → method UNION ALL 3회만)
    - buyer_dn에서 D+0~D+60 누적합을 ARRAY로 패킹
    - LATERAL VIEW EXPLODE(SEQUENCE(-1,60))로 d_preset 확장 (UNION ALL 63개 제거)
    - 최종 결과: ~1,488 rows/box — raw user rows 전송 없음
    """
    def _sql_str(s):
        if "'" not in s:
            return f"'{s}'"
        parts = s.split("'")
        return "CONCAT(" + ",char(39),".join(f"'{p}'" for p in parts) + ")"
    ids_str     = ','.join(_sql_str(b) for b in box_ids)
    min_launch  = min_launch or '2020-01-01'
    max_eff_end = max_eff_end or cutoff
    dn_cols     = ',\n'.join(
        f'    SUM(CASE WHEN day_n <= {n} THEN uc_day ELSE 0 END) AS c{n}'
        for n in range(61)
    )
    cum_array   = 'ARRAY(' + ', '.join(f'c{n}' for n in range(61)) + ')'
    if snap_dates:
        date_lit = ', '.join(f"DATE('{d}')" for d in sorted(snap_dates))
        lv_date_filter = f"p.std_dt IN ({date_lit})"
    else:
        lv_date_filter = """p.std_dt IN (
      SELECT DISTINCT launch_chest_date FROM box_meta UNION ALL
      SELECT DISTINCT prev3m_end        FROM box_meta UNION ALL
      SELECT DISTINCT prev1m_end        FROM box_meta
    )"""
    return f"""
WITH box_meta AS (
  SELECT
    crate AS box_id,
    launch_chest_date,
    LEAST(DATE(max_std_dt), DATE('{cutoff}'))   AS eff_end,
    LAST_DAY(ADD_MONTHS(launch_chest_date, -3)) AS prev3m_end,
    LAST_DAY(ADD_MONTHS(launch_chest_date, -1)) AS prev1m_end
  FROM (
    SELECT *,
      COUNT(*) OVER (PARTITION BY crate, launch_chest_date) AS n_per_launch,
      MAX(std_dt) OVER (PARTITION BY crate) AS max_std_dt
    FROM pubgm_mart.crate_kpi_info
    WHERE country = 'JP' AND crate IN ({ids_str})
  )
  QUALIFY ROW_NUMBER() OVER (
    PARTITION BY crate
    ORDER BY n_per_launch DESC, launch_chest_date ASC, std_dt DESC
  ) = 1
),
jp_users AS (
  SELECT DISTINCT CAST(vopenid AS STRING) AS vopenid
  FROM pubgm_mart.purchasing_power_user
  WHERE country = 'JP' AND std_dt = DATE('{cutoff}')
),
box_txns AS (
  SELECT bm.box_id, CAST(u.vopenid AS STRING) AS vopenid,
    DATEDIFF(u.std_dt, bm.launch_chest_date) AS day_n,
    SUM(CAST(u.amount AS BIGINT)) AS uc_day
  FROM pubgm_mart.user_use_uc_info u
  JOIN box_meta bm ON u.product = bm.box_id
    AND u.std_dt >= bm.launch_chest_date AND u.std_dt <= bm.eff_end
  JOIN jp_users ON CAST(u.vopenid AS STRING) = jp_users.vopenid
  WHERE u.country = 'JP'
    AND u.std_dt >= DATE('{min_launch}')
    AND u.std_dt <= DATE('{max_eff_end}')
  GROUP BY bm.box_id, CAST(u.vopenid AS STRING), DATEDIFF(u.std_dt, bm.launch_chest_date)
),
buyer_dn AS (
  SELECT box_id, vopenid,
    MIN(day_n) AS first_day,
    SUM(uc_day) AS uc_total,
{dn_cols}
  FROM box_txns
  GROUP BY box_id, vopenid
),
buyers AS (
  SELECT DISTINCT vopenid FROM buyer_dn
),
lv_snap AS (
  SELECT CAST(p.vopenid AS STRING) AS vopenid, p.std_dt,
    COALESCE(p.pay_amt_duringLast30days_tag, 'Non-Paid') AS lv
  FROM pubgm_mart.purchasing_power_user p
  LEFT SEMI JOIN buyers b ON CAST(p.vopenid AS STRING) = b.vopenid
  WHERE p.country = 'JP'
    AND {lv_date_filter}
),
classified AS (
  SELECT bd.box_id, bd.vopenid, bd.first_day, bd.uc_total,
    {cum_array} AS cum_uc,
    COALESCE(lp.lv, 'Non-Paid') AS lv_period,
    COALESCE(l3.lv, 'Non-Paid') AS lv_3m,
    COALESCE(l1.lv, 'Non-Paid') AS lv_1m
  FROM buyer_dn bd
  JOIN box_meta bm ON bd.box_id = bm.box_id
  LEFT JOIN lv_snap lp ON bd.vopenid = lp.vopenid AND lp.std_dt = bm.launch_chest_date
  LEFT JOIN lv_snap l3 ON bd.vopenid = l3.vopenid AND l3.std_dt = bm.prev3m_end
  LEFT JOIN lv_snap l1 ON bd.vopenid = l1.vopenid AND l1.std_dt = bm.prev1m_end
),
combos AS (
  SELECT box_id, vopenid, first_day, uc_total, cum_uc, 'period' AS method, lv_period AS pay_level FROM classified
  UNION ALL
  SELECT box_id, vopenid, first_day, uc_total, cum_uc, '3m',    lv_3m                            FROM classified
  UNION ALL
  SELECT box_id, vopenid, first_day, uc_total, cum_uc, '1m',    lv_1m                            FROM classified
)
SELECT box_id, method,
  CASE WHEN d_n = -1 THEN 'total' ELSE CAST(d_n AS STRING) END AS d_preset,
  pay_level,
  COUNT(*) AS users,
  CAST(SUM(CASE WHEN d_n = -1 THEN uc_total ELSE cum_uc[d_n] END) AS BIGINT) AS total_uc
FROM combos
LATERAL VIEW EXPLODE(SEQUENCE(-1, 60)) t AS d_n
WHERE d_n = -1 OR first_day <= d_n
GROUP BY box_id, method, d_preset, pay_level
ORDER BY box_id, method, d_preset, pay_level
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
        for row in ws.iter_rows(min_row=3, values_only=True):
            if len(row) >= 4 and row[2] and row[3]:
                try:
                    names[str(int(float(row[2])))] = str(row[3])
                except (ValueError, TypeError):
                    continue
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
    """쿼리 결과를 {id: {method: {d_preset: {pay_level: {u, uc}}}}} 로 변환"""
    result = {}
    for r in rows:
        pid    = r[id_key]
        method = r['method']
        dp     = r['d_preset']
        lv     = r['pay_level']
        if pid not in result:
            result[pid] = {}
        if method not in result[pid]:
            result[pid][method] = {}
        if dp not in result[pid][method]:
            result[pid][method][dp] = {}
        result[pid][method][dp][lv] = {
            'u': int(r['users'] or 0),
            'uc': int(r['total_uc'] or 0)
        }
    return result

def build_level_from_raw_rows(rows, id_key='box_id'):
    """raw level rows → {id: {method: {d_preset: {pay_level: {'u':..,'uc':..}}}}}
    Input: [{'<id_key>','vopenid','day_n','uc_day','lv_period','lv_3m','lv_1m'}, ...]
    spin/box 공용. D+n 집계는 Python에서 처리 (SQL 62컬럼 + ucpurchase_individual 제거)
    """
    buyers = {}
    for r in rows:
        bid  = r[id_key]
        vid  = r['vopenid']
        dn   = int(r['day_n'] or 0)
        uc   = int(r['uc_day'] or 0)
        key  = (bid, vid)
        if key not in buyers:
            buyers[key] = {
                'days': {},
                'lvs': (r['lv_period'] or 'Non-Paid',
                        r['lv_3m']    or 'Non-Paid',
                        r['lv_1m']    or 'Non-Paid')
            }
        buyers[key]['days'][dn] = buyers[key]['days'].get(dn, 0) + uc

    result = {}
    for (bid, vid), info in buyers.items():
        if bid not in result:
            result[bid] = {}

        days      = info['days']
        lv_period, lv_3m, lv_1m = info['lvs']
        first_day = min(days.keys()) if days else 0
        uc_total  = sum(days.values())

        # cumulative UC: cum[n] = sum of uc for day_n in [0..n]
        cum = [0] * 61
        running = 0
        for n in range(61):
            running += days.get(n, 0)
            cum[n] = running

        for method, lv in (('period', lv_period), ('3m', lv_3m), ('1m', lv_1m)):
            if method not in result[bid]:
                result[bid][method] = {}
            m = result[bid][method]

            # total (all days)
            if 'total' not in m:
                m['total'] = {}
            entry = m['total'].setdefault(lv, {'u': 0, 'uc': 0})
            entry['u']  += 1
            entry['uc'] += uc_total

            # D+n: buyer must have bought by day n
            for n in range(61):
                if first_day <= n:
                    dp = str(n)
                    if dp not in m:
                        m[dp] = {}
                    entry = m[dp].setdefault(lv, {'u': 0, 'uc': 0})
                    entry['u']  += 1
                    entry['uc'] += cum[n]

    return result


def make_spin_npu_sql(spin_ids, cutoff):
    """spin_kpi_stat에서 D+n별 nbu_user/return_user 쿼리 (실제 launch_spin_date 기준)"""
    ids_str = ','.join(str(i) for i in spin_ids)
    return f"""
WITH real_launch AS (
  SELECT spin_contents, MIN(std_dt) AS launch_spin_date
  FROM pubgm_mart.spin_kpi_stat
  WHERE country = 'JP' AND spin_contents IN ({ids_str})
  GROUP BY spin_contents
)
SELECT
  CAST(k.spin_contents AS STRING) AS spin_id,
  DATEDIFF(k.std_dt, rl.launch_spin_date) AS d_preset,
  CAST(k.nbu_user AS BIGINT) AS npu_u,
  CAST(k.nbu_amount AS BIGINT) AS npu_uc,
  CAST(k.return_user AS BIGINT) AS ret_u,
  CAST(k.return_amount AS BIGINT) AS ret_uc
FROM pubgm_mart.spin_kpi_stat k
JOIN real_launch rl ON k.spin_contents = rl.spin_contents AND k.launch_spin_date = rl.launch_spin_date
WHERE k.country = 'JP'
  AND k.spin_contents IN ({ids_str})
  AND k.std_dt <= DATE('{cutoff}')
  AND DATEDIFF(k.std_dt, rl.launch_spin_date) BETWEEN 0 AND 60
QUALIFY ROW_NUMBER() OVER (
  PARTITION BY k.spin_contents, DATEDIFF(k.std_dt, rl.launch_spin_date)
  ORDER BY k.std_dt DESC
) = 1
ORDER BY spin_id, d_preset
"""

def make_box_npu_sql(box_ids, cutoff):
    """crate_kpi_info에서 D+n별 nbu_user/return_user 쿼리 (실제 launch_chest_date 기준)"""
    def _sql_str(s):
        if "'" not in s:
            return f"'{s}'"
        parts = s.split("'")
        return "CONCAT(" + ",char(39),".join(f"'{p}'" for p in parts) + ")"
    ids_str = ','.join(_sql_str(b) for b in box_ids)
    return f"""
WITH real_launch AS (
  SELECT crate, launch_chest_date
  FROM (
    SELECT crate, launch_chest_date,
      COUNT(*) OVER (PARTITION BY crate, launch_chest_date) AS n_per_launch
    FROM pubgm_mart.crate_kpi_info
    WHERE country = 'JP' AND crate IN ({ids_str})
  )
  QUALIFY ROW_NUMBER() OVER (
    PARTITION BY crate ORDER BY n_per_launch DESC, launch_chest_date ASC
  ) = 1
)
SELECT
  k.crate AS box_id,
  DATEDIFF(k.std_dt, rl.launch_chest_date) AS d_preset,
  CAST(k.nbu_user AS BIGINT) AS npu_u,
  CAST(k.nbu_amount AS BIGINT) AS npu_uc,
  CAST(k.return_user AS BIGINT) AS ret_u,
  CAST(k.return_amount AS BIGINT) AS ret_uc
FROM pubgm_mart.crate_kpi_info k
JOIN real_launch rl ON k.crate = rl.crate AND k.launch_chest_date = rl.launch_chest_date
WHERE k.country = 'JP'
  AND k.crate IN ({ids_str})
  AND k.std_dt <= DATE('{cutoff}')
  AND DATEDIFF(k.std_dt, rl.launch_chest_date) BETWEEN 0 AND 60
QUALIFY ROW_NUMBER() OVER (
  PARTITION BY k.crate, DATEDIFF(k.std_dt, rl.launch_chest_date)
  ORDER BY k.std_dt DESC
) = 1
ORDER BY box_id, d_preset
"""

def build_npu_map(rows, id_key='spin_id'):
    """D+n NBU/Return 쿼리 결과를 {id: {npu_dn: {dp:[u,uc]}, ret_dn: {dp:[u,uc]}}} 로 변환"""
    result = {}
    for r in rows:
        pid = r[id_key]
        dp  = str(r['d_preset'])
        if pid not in result:
            result[pid] = {'npu_dn': {}, 'ret_dn': {}}
        result[pid]['npu_dn'][dp] = [int(r['npu_u'] or 0), int(r['npu_uc'] or 0)]
        result[pid]['ret_dn'][dp]  = [int(r['ret_u'] or 0), int(r['ret_uc'] or 0)]
    return result

ACTIVE_THRESHOLD = (date.today() - timedelta(days=3)).isoformat()

def _month_end(y, m):
    """월 말일 반환"""
    if m == 12: return date(y, 12, 31)
    return date(y, m + 1, 1) - timedelta(days=1)

def _prev_month_end(d, months_before):
    """d 기준 months_before 개월 전 월 말일"""
    m, y = d.month - months_before, d.year
    while m <= 0: m += 12; y -= 1
    return _month_end(y, m)

def _snap_dates_for_launches(launch_dates):
    """launch_dates(str or date iterable) → {period, 3m, 1m} 스냅샷 날짜 집합 (리터럴 최적화)"""
    dates = set()
    for ld in launch_dates:
        d = date.fromisoformat(str(ld)[:10]) if ld else None
        if not d: continue
        dates.add(d.isoformat())
        dates.add(_prev_month_end(d, 3).isoformat())
        dates.add(_prev_month_end(d, 1).isoformat())
    return dates

def is_active(last_std_dt_str):
    """spin_kpi_stat std_dt(최신 데이터 날짜)가 3일 이내면 판매 중으로 판단
    (데이터 2일 지연 + 1일 여유)"""
    try:
        return last_std_dt_str >= ACTIVE_THRESHOLD
    except:
        return False

def main():
    # --quick: 기본 통계만 뽑고 HTML 생성 (레벨 쿼리 스킵, 구조 확인용)
    quick = '--quick' in sys.argv
    boxes_only = '--boxes-only' in sys.argv

    def int_arg(name, default):
        prefix = f'--{name}='
        for arg in sys.argv:
            if arg.startswith(prefix):
                return int(arg.split('=', 1)[1])
        return int(os.environ.get(name.upper().replace('-', '_'), default))

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

    spin_level = {}
    box_level  = {}
    spin_npu   = {}
    box_npu    = {}

    if quick:
        print("--quick mode: skipping pay level queries")
    else:
        # 판매 중 vs 완료 분리
        active_spin_ids = [r['spin_id'] for r in raw_spins if is_active(r['last_std_dt'])]
        done_spin_ids   = [r['spin_id'] for r in raw_spins
                           if not is_active(r['last_std_dt']) and r['spin_id'] not in cache['spins']]
        active_box_ids  = [r['box_id'] for r in raw_boxes if is_active(r['last_std_dt'])]
        done_box_ids    = [r['box_id'] for r in raw_boxes
                           if not is_active(r['last_std_dt']) and r['box_id'] not in cache['boxes']]

        print(f"Active spins: {len(active_spin_ids)}, New completed spins: {len(done_spin_ids)}")
        print(f"Active boxes: {len(active_box_ids)}, New completed boxes: {len(done_box_ids)}")

        SPIN_BATCH = int_arg('spin-batch', 20)  # spin: 배치당 날짜범위 최소화로 파티션 프루닝 효율 향상
        BOX_BATCH  = int_arg('box-batch', 20)   # box: user_use_uc_info + purchasing_power_user 스냅샷

        spin_launch   = {r['spin_id']: r.get('launch') for r in raw_spins}
        spin_eff_end  = {r['spin_id']: min(str(r['end_date']), DATA_CUTOFF) if r.get('end_date') else DATA_CUTOFF for r in raw_spins}
        box_launch    = {r['box_id']:  r.get('launch') for r in raw_boxes}
        box_eff_end   = {r['box_id']:  min(str(r['end_date']), DATA_CUTOFF) if r.get('end_date') else DATA_CUTOFF for r in raw_boxes}

        query_spin_ids = [] if boxes_only else active_spin_ids + done_spin_ids
        if boxes_only:
            print("--boxes-only mode: using cached spin level/NPU data")
        if query_spin_ids:
            print(f"Querying spin pay level ({len(query_spin_ids)} products)...")
            all_spin_ids_set = set(done_spin_ids + active_spin_ids)
            for i in range(0, len(query_spin_ids), SPIN_BATCH):
                batch = query_spin_ids[i:i+SPIN_BATCH]
                print(f"  spin batch {i//SPIN_BATCH+1}/{(len(query_spin_ids)-1)//SPIN_BATCH+1} ({len(batch)} products)")
                snap_dates = _snap_dates_for_launches(spin_launch.get(s) for s in batch)
                batch_spin_launches = [spin_launch.get(s) for s in batch if spin_launch.get(s)]
                min_launch = min(batch_spin_launches) if batch_spin_launches else DATA_CUTOFF
                batch_spin_eff_ends = [spin_eff_end.get(s) for s in batch if spin_eff_end.get(s)]
                max_eff_end = max(batch_spin_eff_ends) if batch_spin_eff_ends else DATA_CUTOFF
                rows = run_query(token, make_spin_level_sql(batch, DATA_CUTOFF, snap_dates, min_launch, max_eff_end), f'spin_level_b{i//SPIN_BATCH+1}')
                batch_lv = build_level_map(rows, 'spin_id')
                spin_level.update(batch_lv)
                npu_rows = run_query(token, make_spin_npu_sql(batch, DATA_CUTOFF), f'spin_npu_b{i//SPIN_BATCH+1}')
                batch_npu = build_npu_map(npu_rows, 'spin_id')
                spin_npu.update(batch_npu)
                for sid in batch:
                    if sid in all_spin_ids_set:
                        cache['spins'][sid] = {
                            'level':  spin_level.get(sid, {}),
                            'npu_dn': spin_npu.get(sid, {}).get('npu_dn', {}),
                            'ret_dn': spin_npu.get(sid, {}).get('ret_dn', {}),
                        }
                save_cache(cache)

        query_box_ids = active_box_ids + done_box_ids
        if query_box_ids:
            print(f"Querying box pay level ({len(query_box_ids)} products)...")
            all_box_ids_set = set(done_box_ids + active_box_ids)
            for i in range(0, len(query_box_ids), BOX_BATCH):
                batch = query_box_ids[i:i+BOX_BATCH]
                print(f"  box batch {i//BOX_BATCH+1}/{(len(query_box_ids)-1)//BOX_BATCH+1} ({len(batch)} products)")
                snap_dates = _snap_dates_for_launches(box_launch.get(b) for b in batch)
                batch_launches = [box_launch.get(b) for b in batch if box_launch.get(b)]
                min_launch = min(batch_launches) if batch_launches else DATA_CUTOFF
                batch_box_eff_ends = [box_eff_end.get(b) for b in batch if box_eff_end.get(b)]
                max_eff_end = max(batch_box_eff_ends) if batch_box_eff_ends else DATA_CUTOFF
                rows = run_query(token, make_box_level_sql(batch, DATA_CUTOFF, snap_dates, min_launch, max_eff_end), f'box_level_b{i//BOX_BATCH+1}')
                batch_lv = build_level_map(rows, 'box_id')
                box_level.update(batch_lv)
                npu_rows = run_query(token, make_box_npu_sql(batch, DATA_CUTOFF), f'box_npu_b{i//BOX_BATCH+1}')
                batch_npu = build_npu_map(npu_rows, 'box_id')
                box_npu.update(batch_npu)
                for bid in batch:
                    if bid in all_box_ids_set:
                        cache['boxes'][bid] = {
                            'level':  box_level.get(bid, {}),
                            'npu_dn': box_npu.get(bid, {}).get('npu_dn', {}),
                            'ret_dn': box_npu.get(bid, {}).get('ret_dn', {}),
                        }
                save_cache(cache)

    # 캐시에서 완료 상품 데이터 로드
    for sid, cached in cache['spins'].items():
        if sid not in spin_level:
            spin_level[sid] = cached.get('level', {}) if isinstance(cached, dict) else cached
        if sid not in spin_npu:
            spin_npu[sid] = {'npu_dn': cached.get('npu_dn', {}), 'ret_dn': cached.get('ret_dn', {})} if isinstance(cached, dict) else {}
    for bid, cached in cache['boxes'].items():
        if bid not in box_level:
            box_level[bid] = cached.get('level', {}) if isinstance(cached, dict) else cached
        if bid not in box_npu:
            box_npu[bid] = {'npu_dn': cached.get('npu_dn', {}), 'ret_dn': cached.get('ret_dn', {})} if isinstance(cached, dict) else {}

    # ── JSON 데이터 구축 ──────────────────────────────────────────────────────
    def make_basic(r, id_key):
        return {
            'id':       r[id_key],
            'launch':   r['launch'],
            'end':      r['end_date'],
            'days':     int(r['duration'] or 0),
            'active':   is_active(r['last_std_dt']),
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
        d['name'] = SPIN_NAME_OVERRIDES.get(sid, spin_names.get(sid, sid))
        d['spin_type'] = r.get('spin_type', '')
        d['vtuber']       = int(sid) in VTUBER_SPIN_IDS if sid.isdigit() else False
        d['cat_anime']    = int(sid) in ANIME_SPIN_IDS if sid.isdigit() else False
        d['cat_xsuit']    = sid.endswith('236002')
        d['cat_gilt']     = sid.endswith('073013') or sid.endswith('073014')
        d['cat_sportcar'] = False
        d['cat_mummy']    = False
        d['level'] = spin_level.get(sid, {})
        d['npu_dn'] = spin_npu.get(sid, {}).get('npu_dn', {})
        d['ret_dn'] = spin_npu.get(sid, {}).get('ret_dn', {})
        spin_data.append(d)

    box_data = []
    for r in raw_boxes:
        bid = r['box_id']
        d = make_basic(r, 'box_id')
        d['name'] = BOX_NAME_OVERRIDES.get(bid, bid)
        d['vtuber']       = bid in VTUBER_BOX_NAMES
        d['cat_anime']    = bid in ANIME_BOX_NAMES
        d['cat_xsuit']    = False
        d['cat_gilt']     = False
        d['cat_sportcar'] = bid.startswith('跑车宝箱')
        d['cat_mummy']    = bid in MUMMY_BOX_NAMES
        d['level'] = box_level.get(bid, {})
        d['npu_dn'] = box_npu.get(bid, {}).get('npu_dn', {})
        d['ret_dn'] = box_npu.get(bid, {}).get('ret_dn', {})
        box_data.append(d)

    def compact_level(lv_data):
        """{u:N,uc:M} → [N,M] for smaller JSON output (cache stays unchanged)"""
        if not lv_data:
            return lv_data
        out = {}
        for method, dp_map in lv_data.items():
            out[method] = {}
            for dp, lv_map in dp_map.items():
                out[method][dp] = {
                    lv: [v['u'], v['uc']] if isinstance(v, dict) else v
                    for lv, v in lv_map.items()
                }
        return out

    for d in spin_data:
        d['level'] = compact_level(d['level'])
    for d in box_data:
        d['level'] = compact_level(d['level'])

    gen_date = datetime.now().strftime('%Y-%m-%d %H:%M')
    tpl = os.path.join(os.path.dirname(__file__), 'template.html')
    dist_dir = os.path.join(os.path.dirname(__file__), 'dist')
    os.makedirs(dist_dir, exist_ok=True)
    out = os.path.join(dist_dir, 'index.html')

    with open(tpl) as f:
        html = f.read()

    html = html.replace('%%SPIN_DATA%%', json.dumps(spin_data, ensure_ascii=False, separators=(',', ':')))
    html = html.replace('%%BOX_DATA%%',  json.dumps(box_data,  ensure_ascii=False, separators=(',', ':')))
    html = html.replace('%%GEN_DATE%%',  gen_date)
    html = html.replace('%%CUTOFF%%',    DATA_CUTOFF)

    with open(out, 'w') as f:
        f.write(html)

    print(f"Generated {out} (spins:{len(spin_data)}, boxes:{len(box_data)})")

if __name__ == '__main__':
    main()
