import openpyxl
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── 공통 스타일 ──────────────────────────────────────────────
def border(thin=True):
    s = Side(style='thin' if thin else 'medium')
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(text, bg='1F4E79', fg='FFFFFF', sz=12, bold=True, wrap=False):
    return {
        'value': text,
        'font': Font(name='맑은 고딕', size=sz, bold=bold, color=fg),
        'fill': PatternFill('solid', fgColor=bg),
        'alignment': Alignment(horizontal='center', vertical='center',
                               wrap_text=wrap),
        'border': border()
    }

def cell(text, bg=None, bold=False, align='center', wrap=False, sz=10, fg='000000'):
    f = PatternFill('solid', fgColor=bg) if bg else PatternFill()
    return {
        'value': text,
        'font': Font(name='맑은 고딕', size=sz, bold=bold, color=fg),
        'fill': f,
        'alignment': Alignment(horizontal=align, vertical='center',
                               wrap_text=wrap),
        'border': border()
    }

def apply(ws, row, col, d):
    c = ws.cell(row=row, column=col, value=d['value'])
    c.font = d['font']
    c.fill = d['fill']
    c.alignment = d['alignment']
    c.border = d['border']

def set_col_width(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def set_row_height(ws, heights):
    for r, h in heights.items():
        ws.row_dimensions[r].height = h

# ═══════════════════════════════════════════════════════════════
# Sheet 1 — 표지 / 현황
# ═══════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = '1.표지_현황'
set_col_width(ws1, [4, 18, 28, 18, 18, 18, 18])
ws1.merge_cells('B1:G1')
c = ws1.cell(row=1, column=2,
    value='GKO1808 소켓 내경 타원 발생 개선대책서')
c.font = Font(name='맑은 고딕', size=18, bold=True, color='FFFFFF')
c.fill = PatternFill('solid', fgColor='1F4E79')
c.alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[1].height = 45

ws1.merge_cells('B2:G2')
c = ws1.cell(row=2, column=2,
    value='2죠우 척(2-Jaw Chuck) 클램핑 변형 → 내경 Y방향 타원화 원인 분석 및 대책')
c.font = Font(name='맑은 고딕', size=11, color='FFFFFF')
c.fill = PatternFill('solid', fgColor='2E75B6')
c.alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[2].height = 28

# 기본 정보 테이블
info = [
    ['문서번호', 'GKO1808-QI-002', '작성일', '2026-05-15', '버전', 'Rev.1'],
    ['협력사', '부경하이텍', '품번', 'GKO1808', '품명', '소켓(Socket)'],
    ['내경공차', '+0.05 / 0 mm', '측정방법', '삼차원측정기(CMM)', '문서구분', '개선대책서'],
    ['원인확정', '2죠우 척 Y방향 클램핑 탄성변형', '담당부서', '품질관리팀', '승인', ''],
]
for r, row_data in enumerate(info, 4):
    for c_idx in range(0, 6, 2):
        apply(ws1, r, 2, hdr(row_data[c_idx], bg='BDD7EE', fg='1F4E79', sz=10))
        d = cell(row_data[c_idx+1], bold=(c_idx==0))
        apply(ws1, r, 3 if c_idx==0 else (5 if c_idx==2 else 7), d)
    ws1.row_dimensions[r].height = 22

# 공백
ws1.row_dimensions[8].height = 10

# 부적합 현황 헤더
ws1.merge_cells('B9:G9')
apply(ws1, 9, 2, hdr('■ 부적합 현황 및 CMM 측정 결과', bg='1F4E79', sz=12))
ws1.row_dimensions[9].height = 28

headers = ['항목', '내용', '측정방향', 'CMM 편차', '공차', '판정']
for i, h in enumerate(headers, 2):
    apply(ws1, 10, i, hdr(h, bg='2E75B6', sz=10))
ws1.row_dimensions[10].height = 22

rows = [
    ['불량유형', '소켓 내경 타원 발생 (진원도 불량)', 'Y방향 (수직, 조우 방향)', '0.05 ~ 0.08 mm', '+0.05 / 0', '❌ 불량'],
    ['X방향 편차', '조우 직각방향 — Poisson 수축', 'X방향 (수평)', '0.05 ~ 0.08 mm', '+0.05 / 0', '⚠ 경계'],
    ['Y방향 편차', '조우 방향 — 탄성 복원 후 팽창', 'Y방향 (수직)', 'X보다 크게 초과', '+0.05 / 0', '❌ 초과'],
    ['클램핑방식', 'CNC 선반 2죠우 척', '—', '—', '—', '원인'],
    ['측정방법', '삼차원 측정기 (CMM)', '4방향 이상', '—', '진원도 ≤0.03', '기준'],
]
fills = ['DEEAF1', 'FFF2CC', 'FFE0E0', 'DEEAF1', 'DEEAF1']
for r, (row_data, bg) in enumerate(zip(rows, fills), 11):
    for c_idx, val in enumerate(row_data, 2):
        clr = 'FFE0E0' if val == '❌ 불량' or val == '❌ 초과' else (
              'FFF2CC' if val == '⚠ 경계' else bg)
        bold = (c_idx == 2) or val.startswith('❌')
        apply(ws1, r, c_idx, cell(val, bg=clr, bold=bold,
              align='left' if c_idx == 3 else 'center'))
    ws1.row_dimensions[r].height = 22


# ═══════════════════════════════════════════════════════════════
# Sheet 2 — 원인 분석
# ═══════════════════════════════════════════════════════════════
ws2 = wb.create_sheet('2.원인분석')
set_col_width(ws2, [4, 22, 38, 20, 14])
ws2.merge_cells('B1:E1')
apply(ws2, 1, 2, hdr('■ 2죠우 척 클램핑 변형 메커니즘 — 원인 분석', bg='1F4E79', sz=13))
ws2.row_dimensions[1].height = 36

# 메커니즘 설명
mech = [
    ('STEP 1', '2죠우 척 Y방향 클램핑',
     '상·하 2개 조우가 소재를 Y방향(수직)으로 압축\n→ 소재 OD_Y 감소, 내경 ID_Y도 압축되어 감소',
     '탄성 압축 발생', '1F4E79'),
    ('STEP 2', 'Poisson 효과 (X방향 팽창)',
     'Y방향 압축에 의한 포아송 반응\n→ X방향(수평)은 구속 없으므로 소재 팽창\n→ ID_X 증가',
     'X방향 팽창', '2E75B6'),
    ('STEP 3', '클램핑 상태에서 내경 보링 가공',
     '변형된(Y압축·X팽창) 상태의 소재를 기준으로 원형 보링\n→ 공구 궤적은 원형이나 소재는 이미 변형 중',
     '가공 기준 오류', 'C55A11'),
    ('STEP 4', '척 해제 → 탄성 복원',
     'Y방향: 압축 해제 → 탄성 복원 → ID_Y 증가 (팽창)\n  ※ 공차 상한 +0.05mm 초과 → 불량\nX방향: 팽창 해제 → 수축 → ID_X 감소',
     'Y방향 내경 과대', 'C00000'),
    ('결과', '타원형 내경 불량 확정',
     'ID_Y > ID_X → 타원 발생\nY방향 내경 = 공차 상한 초과 → GKO1808 부적합\nX방향 내경 = 하한 미달 가능성',
     '공차 이탈', 'C00000'),
]

headers2 = ['단계', '현상', '상세 설명', '영향', '']
for i, h in enumerate(headers2, 2):
    apply(ws2, 2, i, hdr(h, bg='2E75B6', sz=10))
ws2.row_dimensions[2].height = 22

for r, (step, phen, desc, effect, color) in enumerate(mech, 3):
    apply(ws2, r, 2, hdr(step, bg=color, sz=10))
    apply(ws2, r, 3, cell(phen, bold=True, bg='F2F2F2'))
    apply(ws2, r, 4, cell(desc, align='left', wrap=True, bg='FFFFFF'))
    apply(ws2, r, 5, cell(effect, bold=True,
          bg='FFE0E0' if color=='C00000' else 'FFF2CC'))
    ws2.row_dimensions[r].height = 55

ws2.row_dimensions[8].height = 15

# 변형 방향 요약표
ws2.merge_cells('B9:E9')
apply(ws2, 9, 2, hdr('■ Y방향 클램핑 시 내경 변화 요약', bg='1F4E79', sz=11))
ws2.row_dimensions[9].height = 28

for i, h in enumerate(['구분', '클램핑 중 (가공 시)', '척 해제 후 (완성품)', '판정'], 2):
    apply(ws2, 10, i, hdr(h, bg='2E75B6', sz=10))
ws2.row_dimensions[10].height = 22

summary = [
    ['Y방향 내경\n(조우 압축 방향)', '압축 → ID_Y 감소\n(공구는 목표치로 가공)', '탄성 복원 → ID_Y 증가\n→ 목표치 + ΔY', '❌ 공차 상한 초과'],
    ['X방향 내경\n(조우 직각 방향)', 'Poisson 팽창 → ID_X 증가\n(공구는 목표치로 가공)', '수축 → ID_X 감소\n→ 목표치 - ΔX', '⚠ 공차 하한 주의'],
    ['진원도\n(타원 편차)', 'Y>X 또는 X>Y 변형 상태', 'Y방향 > X방향\n타원 형상 고착', '❌ 진원도 불량'],
]
fills2 = ['FFE0E0', 'FFF2CC', 'FFE0E0']
for r, (row_data, bg) in enumerate(zip(summary, fills2), 11):
    for c_idx, val in enumerate(row_data, 2):
        bold = '❌' in val or '⚠' in val
        apply(ws2, r, c_idx, cell(val, bg=bg, bold=bold, wrap=True,
              align='center'))
    ws2.row_dimensions[r].height = 45


# ═══════════════════════════════════════════════════════════════
# Sheet 3 — 개선 대책
# ═══════════════════════════════════════════════════════════════
ws3 = wb.create_sheet('3.개선대책')
set_col_width(ws3, [4, 14, 30, 20, 12, 12, 14])
ws3.merge_cells('B1:G1')
apply(ws3, 1, 2, hdr('■ 단계별 개선 대책', bg='1F4E79', sz=13))
ws3.row_dimensions[1].height = 36

for i, h in enumerate(['단계', '대책명', '세부 내용', '기대효과', '비용', '기간', '담당'], 2):
    apply(ws3, 2, i, hdr(h, bg='2E75B6', sz=10))
ws3.row_dimensions[2].height = 22

actions = [
    ('즉시\n(당일)', '정삭 클램핑력\n50% 감소',
     '• 황삭: 정상 클램핑\n• 정삭: 클램핑 토크 50%로 감소\n  → 탄성 변형량 감소 → 복원량 감소',
     'Y방향 편차\n30~50% 감소', '없음', '즉시', '생산팀', 'E2EFDA'),
    ('즉시\n(당일)', '스프링 패스\n(Spring Pass) 추가',
     '• 정삭 후 절삭깊이 0으로 공구 재통과\n• 가공 중 탄성변형 여유재료 제거\n  → 진원도 개선',
     '진원도 개선\n0.03mm 이하 목표', '없음', '즉시', '생산팀', 'E2EFDA'),
    ('1~2주', '소프트 조우\n(Soft Jaw) 제작',
     '• 소재 OD와 동일 반경으로 조우 내면 가공\n• 접촉면적 최대화 → 단위압력 감소\n• 접촉각 90° 이상 확보',
     '클램핑 편차\n균등화', '소', '1~2주', '생산기술', 'FFF2CC'),
    ('1~2주', 'CNC 보정\n가공 적용',
     '• 정삭 후 척 해제 → CMM 측정\n• Y방향 복원량(ΔY) 측정\n• 정삭 목표치 = 공칭 - ΔY 로 보정\n  예) ΔY=0.04mm → 목표치 -0.04 설정',
     '보정 후\n공차 만족', '없음', '즉시\n~1주', '생산기술', 'FFF2CC'),
    ('1개월', '3죠우 척으로\n교체 ★핵심',
     '• 120° 균등 3방향 클램핑\n• Y방향 단독 압축 → 3방향 균등 압축으로 변경\n• 탄성 변형 방향성 소멸 → 타원 발생 근본 해소\n• 원형 소재에 가장 적합한 척 방식',
     '타원 불량\n근본 해소\nCpk≥1.33', '중', '1개월', '생산기술', 'DEEAF1'),
    ('3개월', '콜렛 척\n도입 검토',
     '• 360° 균등 면 클램핑\n• 런아웃 ±0.003mm 이하\n• OD 치수 편차 허용 좁음 → 소재 관리 필요',
     '최고 정밀도\n진원도 확보', '대', '3개월', '생산기술', 'DEEAF1'),
]

for r, (stage, name, desc, effect, cost, period, dept, bg) in enumerate(actions, 3):
    apply(ws3, r, 2, cell(stage, bg='1F4E79', fg='FFFFFF', bold=True, wrap=True))
    apply(ws3, r, 3, cell(name, bg=bg, bold=True, wrap=True))
    apply(ws3, r, 4, cell(desc, bg='FFFFFF', align='left', wrap=True))
    apply(ws3, r, 5, cell(effect, bg=bg, wrap=True))
    apply(ws3, r, 6, cell(cost, bg='FFF2CC' if cost=='없음' else 'FFE0E0' if cost=='대' else 'FFFFFF'))
    apply(ws3, r, 7, cell(period, wrap=True))
    apply(ws3, r, 8, cell(dept))
    ws3.row_dimensions[r].height = 70


# ═══════════════════════════════════════════════════════════════
# Sheet 4 — 절삭조건 & 측정기준
# ═══════════════════════════════════════════════════════════════
ws4 = wb.create_sheet('4.절삭조건_측정기준')
set_col_width(ws4, [4, 22, 22, 22, 18])
ws4.merge_cells('B1:F1')
apply(ws4, 1, 2, hdr('■ 최적 절삭 조건 및 측정 관리 기준', bg='1F4E79', sz=13))
ws4.row_dimensions[1].height = 36

# 절삭조건
ws4.merge_cells('B2:F2')
apply(ws4, 2, 2, hdr('[ 절삭 조건 개선안 ]', bg='2E75B6', sz=11))
ws4.row_dimensions[2].height = 26

for i, h in enumerate(['항목', '현행', '황삭 개선', '정삭 개선', '비고'], 2):
    apply(ws4, 3, i, hdr(h, bg='BDD7EE', fg='1F4E79', sz=10))
ws4.row_dimensions[3].height = 22

cutting = [
    ['클램핑 토크', '—', '100% (정상)', '50% 이하로 감소', '핵심 개선'],
    ['절삭 깊이', '—', '0.3 ~ 0.5 mm', '0.05 ~ 0.10 mm', '진원도 확보'],
    ['이송 속도', '—', '정상', '감속 적용', '절삭력 최소화'],
    ['스프링 패스', '미실시', '—', '절삭깊이 0, 1회 추가', '필수 적용'],
    ['공구 돌출량', '—', '3D 이하', '3D 이하', 'D=공구경'],
    ['절삭유', '—', '충분 공급', '충분 공급', '열변형 방지'],
]
for r, row_data in enumerate(cutting, 4):
    bgs = ['F2F2F2', 'FFFFFF', 'E2EFDA', 'E2EFDA', 'FFFFFF']
    for c_idx, (val, bg) in enumerate(zip(row_data, bgs), 2):
        bold = c_idx == 2
        apply(ws4, r, c_idx, cell(val, bg=bg, bold=bold, wrap=True))
    ws4.row_dimensions[r].height = 30

ws4.row_dimensions[10].height = 15

# 측정기준
ws4.merge_cells('B11:F11')
apply(ws4, 11, 2, hdr('[ CMM 측정 관리 기준 — 개선 ]', bg='2E75B6', sz=11))
ws4.row_dimensions[11].height = 26

for i, h in enumerate(['측정항목', '현행', '개선 기준', '측정시점', '판정기준'], 2):
    apply(ws4, 12, i, hdr(h, bg='BDD7EE', fg='1F4E79', sz=10))
ws4.row_dimensions[12].height = 22

measure = [
    ['내경 측정 방향', '1방향', '4방향 이상\n(0°/45°/90°/135°)', '공정 중 + 완성품', '전 방향 +0.05/0 만족'],
    ['진원도 계산', '미실시', '최대내경 - 최소내경\n= 진원도', '완성품 전수', '≤ 0.03 mm'],
    ['Y방향 복원량(ΔY)', '미관리', 'ΔY = 완성내경Y - 목표\n가공 보정값으로 활용', '로트 초물', 'ΔY ≤ 0.02 mm 목표'],
    ['측정 빈도', '샘플', '전수 CMM 측정', '완성품', '불합격 즉시 피드백'],
    ['척 해제 전·후 비교', '미실시', '해제 전/후 측정 비교\n→ 탄성복원량 관리', '로트별 초물', '복원량 기록·관리'],
]
for r, row_data in enumerate(measure, 13):
    bgs = ['F2F2F2', 'FFE0E0', 'E2EFDA', 'FFFFFF', 'DEEAF1']
    for c_idx, (val, bg) in enumerate(zip(row_data, bgs), 2):
        apply(ws4, r, c_idx, cell(val, bg=bg, bold=(c_idx==2), wrap=True,
              align='left' if c_idx in [3,4,5,6] else 'center'))
    ws4.row_dimensions[r].height = 42


# ═══════════════════════════════════════════════════════════════
# Sheet 5 — 실행계획 & 효과확인
# ═══════════════════════════════════════════════════════════════
ws5 = wb.create_sheet('5.실행계획_효과확인')
set_col_width(ws5, [4, 20, 16, 14, 14, 14, 14, 14])
ws5.merge_cells('B1:I1')
apply(ws5, 1, 2, hdr('■ 실행 계획 (Action Plan) 및 효과 확인', bg='1F4E79', sz=13))
ws5.row_dimensions[1].height = 36

# 실행계획 Gantt
ws5.merge_cells('B2:I2')
apply(ws5, 2, 2, hdr('[ 주차별 실행 계획 ]', bg='2E75B6', sz=11))
ws5.row_dimensions[2].height = 26

for i, h in enumerate(['대책 항목', '담당', '1주차', '2주차', '3주차', '4주차', '2개월', '3개월'], 2):
    apply(ws5, 3, i, hdr(h, bg='BDD7EE', fg='1F4E79', sz=10))
ws5.row_dimensions[3].height = 22

gantt = [
    ['정삭 클램핑력 50% 감소', '생산팀', '●', '●', '●', '●', '●', '●'],
    ['스프링 패스 추가', '생산팀', '●', '●', '●', '●', '●', '●'],
    ['기존 재공품 CMM 전수 선별', '품질팀', '●', '', '', '', '', ''],
    ['소프트 조우 제작·적용', '생산기술', '○', '●', '●', '', '', ''],
    ['CNC 보정값 설정·적용', '생산기술', '○', '●', '●', '', '', ''],
    ['CMM 4방향 측정 기준 적용', '품질팀', '●', '●', '●', '●', '●', '●'],
    ['3죠우 척 발주·교체', '생산기술', '', '', '○', '●', '', ''],
    ['작업표준서·FMEA 개정', '품질팀', '', '○', '●', '', '', ''],
    ['콜렛 척 도입 검토', '생산기술', '', '', '', '', '○', '●'],
    ['Cpk 관리체계 구축', '품질팀', '', '', '', '', '○', '●'],
]
for r, row_data in enumerate(gantt, 4):
    for c_idx, val in enumerate(row_data, 2):
        bg = 'E2EFDA' if val == '●' else ('FFF2CC' if val == '○' else 'FFFFFF')
        bold = val in ['●', '○']
        apply(ws5, r, c_idx, cell(val, bg=bg if c_idx > 3 else (
              'F2F2F2' if c_idx == 2 else 'FFFFFF'), bold=bold))
    ws5.row_dimensions[r].height = 22

ws5.row_dimensions[14].height = 15

# 효과 확인
ws5.merge_cells('B15:I15')
apply(ws5, 15, 2, hdr('[ 개선 효과 확인 지표 ]', bg='2E75B6', sz=11))
ws5.row_dimensions[15].height = 26

for i, h in enumerate(['관리 지표', '현행', '즉시 대책 후', '소프트조우 후', '3죠우 교체 후', '목표', '확인방법', '확인주기'], 2):
    apply(ws5, 16, i, hdr(h, bg='BDD7EE', fg='1F4E79', sz=10))
ws5.row_dimensions[16].height = 22

effects = [
    ['Y방향 내경 편차', '+0.05 초과', '+0.03~0.04', '+0.02~0.03', '0.01 이하', '+0.05/0 내', 'CMM', '전수'],
    ['진원도 (X-Y 편차)', '0.05~0.08mm', '0.04~0.06mm', '0.02~0.04mm', '0.01~0.02mm', '≤0.03mm', 'CMM', '전수'],
    ['내경 공차 합격률', '불량 다수', '70% 이상', '90% 이상', '99% 이상', '100%', 'CMM', '전수'],
    ['공정능력 Cpk', '< 1.0', '≥ 1.0', '≥ 1.17', '≥ 1.33', '≥ 1.33', 'SPC', '월 1회'],
    ['고객 클레임', '발생', '0건 목표', '0건 목표', '0건', '0건', '이력관리', '월 1회'],
]
for r, row_data in enumerate(effects, 17):
    bgs = ['F2F2F2', 'FFE0E0', 'FFF2CC', 'FFF2CC', 'E2EFDA', 'DEEAF1', 'FFFFFF', 'FFFFFF']
    for c_idx, (val, bg) in enumerate(zip(row_data, bgs), 2):
        bold = c_idx in [2, 7]
        apply(ws5, r, c_idx, cell(val, bg=bg, bold=bold, wrap=True))
    ws5.row_dimensions[r].height = 28


# ═══════════════════════════════════════════════════════════════
# Sheet 6 — 승인 & 재발방지
# ═══════════════════════════════════════════════════════════════
ws6 = wb.create_sheet('6.재발방지_승인')
set_col_width(ws6, [4, 24, 32, 16, 12])
ws6.merge_cells('B1:F1')
apply(ws6, 1, 2, hdr('■ 재발 방지 대책 및 결론', bg='1F4E79', sz=13))
ws6.row_dimensions[1].height = 36

# 핵심 요약
ws6.merge_cells('B2:F2')
apply(ws6, 2, 2, hdr('[ 핵심 원인 및 결론 ]', bg='C00000', sz=11))
ws6.row_dimensions[2].height = 26

conclusions = [
    '확정 원인: CNC 선반 2죠우 척이 Y방향(수직)으로 소재를 압축 클램핑',
    '변형 기전: Y방향 압축 → 가공 후 척 해제 시 탄성 복원 → Y방향 내경 증가(팽창)',
    '불량 결과: Y방향 내경 = 공차 상한 +0.05mm 초과 → GKO1808 내경 부적합',
    '근본 해결: 2죠우 척 → 3죠우 척 교체로 120° 균등 클램핑 확보',
    '즉시 조치: 정삭 클램핑력 50% 감소 + 스프링 패스 + CNC 보정 가공 적용',
]
for r, text in enumerate(conclusions, 3):
    ws6.merge_cells(f'B{r}:F{r}')
    bg = 'FFE0E0' if '확정 원인' in text or '불량' in text else (
         'E2EFDA' if '근본' in text or '즉시' in text else 'FFF2CC')
    apply(ws6, r, 2, cell(text, bg=bg, bold=True, align='left', sz=11))
    ws6.row_dimensions[r].height = 26

ws6.row_dimensions[8].height = 15

# 재발방지
ws6.merge_cells('B9:F9')
apply(ws6, 9, 2, hdr('[ 재발 방지 관리 계획 ]', bg='2E75B6', sz=11))
ws6.row_dimensions[9].height = 26

for i, h in enumerate(['관리 항목', '개정/추가 내용', '적용 시점', '담당'], 2):
    apply(ws6, 10, i, hdr(h, bg='BDD7EE', fg='1F4E79', sz=10))
ws6.row_dimensions[10].height = 22

prevent = [
    ['작업표준서', '정삭 클램핑 토크 기준값 명기\n2단 가공법(황삭/정삭) 절차 추가\n스프링 패스 필수 실시 명기', '1주 이내', '생산기술'],
    ['검사기준서', 'CMM 4방향 측정 및 진원도 기록\n내경 + 진원도 동시 판정 기준 추가', '1주 이내', '품질팀'],
    ['공정 FMEA', '2죠우 척 탄성변형 고장모드 추가\n탄성복원 → 내경 과대 영향 반영', '2주 이내', '품질팀'],
    ['관리 계획서', '진원도 공정능력(Cpk) 관리 항목 추가\n척 조우 마모 점검 주기 명기', '1개월', '품질팀'],
    ['수평 전개', '2죠우 척 사용 전 품종 동일 검토\n소프트 조우/3죠우 교체 확대 적용', '1개월', '생산기술'],
]
for r, row_data in enumerate(prevent, 11):
    bgs = ['F2F2F2', 'FFFFFF', 'FFF2CC', 'FFFFFF']
    for c_idx, (val, bg) in enumerate(zip(row_data, bgs), 2):
        apply(ws6, r, c_idx, cell(val, bg=bg, bold=(c_idx==2),
              align='left', wrap=True))
    ws6.row_dimensions[r].height = 45

ws6.row_dimensions[16].height = 20

# 승인란
ws6.merge_cells('B17:F17')
apply(ws6, 17, 2, hdr('[ 검토 및 승인 ]', bg='1F4E79', sz=11))
ws6.row_dimensions[17].height = 26

for i, h in enumerate(['구분', '작성', '검토', '승인', '비고'], 2):
    apply(ws6, 18, i, hdr(h, bg='2E75B6', sz=10))
ws6.row_dimensions[18].height = 22

for r, label in enumerate(['성명', '서명', '일자'], 19):
    vals = [label, '', '', '', '2026.05.15' if label == '일자' else '']
    for c_idx, val in enumerate(vals, 2):
        apply(ws6, r, c_idx, cell(val, bg='F2F2F2' if c_idx == 2 else 'FFFFFF',
              bold=(c_idx == 2)))
    ws6.row_dimensions[r].height = 30

# ── 저장 ────────────────────────────────────────────────────────
path = '/home/user/quality-analysis/GKO1808_소켓내경타원발생_2죠우척_개선대책서_Rev1.xlsx'
wb.save(path)
print(f'저장 완료: {path}')
