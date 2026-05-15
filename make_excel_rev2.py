import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── 공통 스타일 ──────────────────────────────────────────────
def bd():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)

def med_bd():
    sm = Side(style='medium')
    st = Side(style='thin')
    return Border(left=sm, right=sm, top=sm, bottom=sm)

def hdr(text, bg='1F4E79', fg='FFFFFF', sz=11, bold=True, wrap=True):
    return dict(value=text,
        font=Font(name='맑은 고딕', size=sz, bold=bold, color=fg),
        fill=PatternFill('solid', fgColor=bg),
        alignment=Alignment(horizontal='center', vertical='center', wrap_text=wrap),
        border=bd())

def cell(text, bg=None, bold=False, align='center', wrap=True, sz=10, fg='000000'):
    return dict(value=text,
        font=Font(name='맑은 고딕', size=sz, bold=bold, color=fg),
        fill=PatternFill('solid', fgColor=bg) if bg else PatternFill(),
        alignment=Alignment(horizontal=align, vertical='center', wrap_text=wrap),
        border=bd())

def ap(ws, r, c, d):
    x = ws.cell(row=r, column=c, value=d['value'])
    for k in ('font','fill','alignment','border'):
        setattr(x, k, d[k])

def widths(ws, w_list):
    for i, w in enumerate(w_list, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def rh(ws, d):
    for r, h in d.items():
        ws.row_dimensions[r].height = h

# ═══════════════════════════════════════════════════════════════════
# Sheet 1 — 표지 / 동영상 검증 결과
# ═══════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = '1.표지_검증결과'
widths(ws1, [3, 20, 26, 18, 18, 14, 14])

# 제목
ws1.merge_cells('B1:G1')
ap(ws1,1,2, hdr('GKO1808 소켓 내경 타원 발생 개선대책서  Rev.2',
                bg='1F4E79', sz=16))
ws1.merge_cells('B2:G2')
ap(ws1,2,2, hdr('2죠우 유압 척(2-Jaw Hydraulic Chuck) — 동영상 현장 검증 반영',
                bg='2E75B6', sz=11))
rh(ws1,{1:42,2:26})

# 문서정보
info = [
    ('문서번호','GKO1808-QI-003','작성일','2026-05-15','버전','Rev.2'),
    ('협력사','부경하이텍','품번','GKO1808','품명','소켓(Socket)'),
    ('내경공차','+0.05 / 0 mm','척 종류','2죠우 유압 척','측정','삼차원(CMM)'),
    ('개정 사유','동영상 현장 검증 결과 반영 (척 종류 수정·클램핑력 동일 확인)','','','',''),
]
for r,row in enumerate(info,4):
    ap(ws1,r,2, hdr(row[0],bg='BDD7EE',fg='1F4E79',sz=10))
    if r==7:
        ws1.merge_cells(f'C{r}:G{r}')
        ap(ws1,r,3, cell(row[1],bold=True))
    else:
        ap(ws1,r,3, cell(row[1],bold=True))
        ap(ws1,r,4, hdr(row[2],bg='BDD7EE',fg='1F4E79',sz=10))
        ap(ws1,r,5, cell(row[3]))
        ap(ws1,r,6, hdr(row[4],bg='BDD7EE',fg='1F4E79',sz=10))
        ap(ws1,r,7, cell(row[5]))
    rh(ws1,{r:22})

rh(ws1,{8:12})

# ── 동영상 검증 결과 ──
ws1.merge_cells('B9:G9')
ap(ws1,9,2, hdr('■ 동영상 현장 검증 결과 (체크리스트)',bg='C00000',sz=12))
rh(ws1,{9:28})

hdrs=['확인 항목','개선대책서 내용','동영상 확인 결과','일치 여부','조치 필요']
for i,h in enumerate(hdrs,2):
    ap(ws1,10,i+1, hdr(h,bg='2E75B6',sz=10))
ap(ws1,10,2, hdr('No.',bg='2E75B6',sz=10))
rh(ws1,{10:22})

checks = [
    ('1','조우 개수','2죠우 척','2죠우 확인 ✔','✅ 일치','없음'),
    ('2','클램핑 방향','Y방향(수직) 압축','수직 방향 확인 ✔','✅ 일치','없음'),
    ('3','클램핑 방식','유압 척 권고 대책 있었음',
     '이미 유압 척 사용 중\n→ 기존 대책 오류',
     '❌ 대책 오류','대책서 수정 완료\n(유압→3죠우로 변경)'),
    ('4','정삭 클램핑력','정삭 시 50% 감소 필요',
     '황삭=정삭 동일 클램핑\n→ 미개선 상태',
     '⚠ 핵심 문제','즉시 유압압력 감소\n(정삭 시 50% 이하)'),
    ('5','스프링 패스','정삭 후 1회 추가','확인 불명확','❓ 미확인','현장 확인 필요'),
    ('6','가공 중 진동','채터 진동 점검','확인 불명확','❓ 미확인','현장 확인 필요'),
    ('7','척 해제 후 측정','CMM 측정 필수','미실시 확인','❌ 미실시','즉시 전수 측정 추가'),
    ('8','척킹 위치','편심 없이 균일 파지','동영상 참조\n(육안 확인 필요)','❓ 확인 필요','척킹 상태 재확인'),
]
fill_map = {'✅ 일치':'E2EFDA','❌ 대책 오류':'FFE0E0','⚠ 핵심 문제':'FFE0E0',
            '❓ 미확인':'FFF2CC','❌ 미실시':'FFE0E0','❓ 확인 필요':'FFF2CC'}
for r,(no,item,doc,vid,match,action) in enumerate(checks,11):
    row_bg = fill_map.get(match,'FFFFFF')
    ap(ws1,r,2, cell(no,bg='F2F2F2',bold=True))
    ap(ws1,r,3, cell(item,bg='F2F2F2',bold=True))
    ap(ws1,r,4, cell(doc,bg='FFFFFF',align='left'))
    ap(ws1,r,5, cell(vid,bg='FFF9F0' if '오류' in vid or '동일' in vid else 'FFFFFF',
                     align='left',bold='오류' in vid or '동일' in vid))
    ap(ws1,r,6, cell(match,bg=row_bg,bold=True))
    ap(ws1,r,7, cell(action,bg=row_bg if action!='없음' else 'FFFFFF',
                     align='left',bold=action!='없음'))
    rh(ws1,{r:42})

# 검증 총평
rh(ws1,{19:12})
ws1.merge_cells('B20:G20')
ap(ws1,20,2, hdr('▶ 검증 총평: 유압 척 이미 사용 중(대책 수정) + 정삭 클램핑력 황삭 동일(핵심 미개선) + 척 해제 후 측정 없음 → Rev.2 즉시 적용 필요',
                bg='C00000',fg='FFFFFF',sz=10))
rh(ws1,{20:30})


# ═══════════════════════════════════════════════════════════════════
# Sheet 2 — 원인 분석 (수정)
# ═══════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet('2.원인분석_수정')
widths(ws2, [3,16,30,22,16])
ws2.merge_cells('B1:F1')
ap(ws2,1,2, hdr('■ 2죠우 유압 척 — Y방향 클램핑 탄성복원 메커니즘 (확정)',bg='1F4E79',sz=13))
rh(ws2,{1:36})

ws2.merge_cells('B2:F2')
ap(ws2,2,2, hdr('[ 핵심: 유압 척이라도 2죠우 구조이면 Y방향 탄성변형 동일하게 발생 ]',
                bg='C00000',sz=11))
rh(ws2,{2:28})

mechs = [
    ('STEP 1','유압 2죠우 척\nY방향 클램핑',
     '상·하 유압 조우가 소재를 Y방향(수직)으로 균일 압축\n'
     '→ 유압이므로 반복 클램핑력은 일정하나\n'
     '→ 2죠우 구조적 문제(Y방향 단독 압축)는 동일하게 존재',
     '탄성 압축\n(Y방향)','1F4E79'),
    ('STEP 2','Poisson 효과\nX방향 팽창',
     'Y방향 압축 → 포아송 효과로 X방향(수평) 팽창\n'
     'X방향 내경 증가 / Y방향 내경 감소 (가공 중 상태)',
     'X방향 팽창','2E75B6'),
    ('STEP 3','황삭=정삭\n동일 클램핑력\n【동영상 확인】',
     '⚠ 동영상 검증 결과: 정삭 시에도 황삭과 동일한 유압 클램핑력 적용\n'
     '→ 정삭 가공 중 최대 탄성 변형 상태 유지\n'
     '→ 탄성 복원량 최대화 → 불량 심화',
     '【핵심 문제】\n복원량 최대','C00000'),
    ('STEP 4','척 해제\n탄성 복원\n【동영상 확인】',
     'Y방향: 압축 해제 → 탄성 복원 → ID_Y 증가 (팽창)\n'
     '  → 공차 상한 +0.05mm 초과 → ❌ 불량\n'
     'X방향: 팽창 해제 → 수축 → ID_X 감소\n'
     '  → 공차 하한 0mm 미달 가능성\n'
     '⚠ 척 해제 후 CMM 측정 없음(동영상 확인) → 탄성복원량 미파악',
     'Y↑ 과대\nX↓ 과소','C00000'),
    ('결과','타원 불량\n확정',
     'ID_Y > ID_X → 타원 발생\n'
     'Y방향 내경 공차 상한 +0.05mm 초과 → GKO1808 부적합\n'
     '정삭 클램핑력 동일 + 해제 후 측정 없음 → 불량 지속 구조',
     '공차 이탈\n검출 불가','C00000'),
]
for i,h in enumerate(['단계','현상','상세 설명 (동영상 검증 반영)','영향',''],2):
    ap(ws2,3,i, hdr(h,bg='2E75B6',sz=10))
rh(ws2,{3:22})
for r,(step,phen,desc,eff,col) in enumerate(mechs,4):
    ap(ws2,r,2, hdr(step,bg=col,sz=10))
    ap(ws2,r,3, cell(phen,bg='F2F2F2',bold=True))
    ap(ws2,r,4, cell(desc,align='left',wrap=True,bg='FFFFFF'))
    ap(ws2,r,5, cell(eff,bold=True,
                     bg='FFE0E0' if col=='C00000' else 'FFF2CC'))
    rh(ws2,{r:65})

rh(ws2,{9:14})
# 변화 요약
ws2.merge_cells('B10:F10')
ap(ws2,10,2, hdr('■ Y방향 유압 클램핑 → 내경 변화 요약 (동영상 검증 확정)',bg='1F4E79',sz=11))
rh(ws2,{10:28})
for i,h in enumerate(['구분','클램핑 중 (가공 시)','척 해제 후 (완성품)','판정'],2):
    ap(ws2,11,i, hdr(h,bg='2E75B6',sz=10))
rh(ws2,{11:22})
summary=[
    ('Y방향 내경\n(조우 압축방향)',
     'Y방향 압축 → ID_Y 감소\n(보링공구는 목표치로 가공)',
     '탄성 복원 → ID_Y 증가\n목표치 + ΔY\n정삭 클램핑=황삭이므로 ΔY 최대',
     '❌ 공차 상한\n+0.05 초과'),
    ('X방향 내경\n(조우 직각방향)',
     'Poisson 팽창 → ID_X 증가\n(보링공구는 목표치로 가공)',
     '수축 → ID_X 감소\n목표치 - ΔX',
     '⚠ 공차 하한\n0mm 미달 주의'),
    ('진원도\n(타원 편차)',
     'Y·X 불균등 변형 상태',
     'Y > X → 타원 고착\nCMM X/Y편차 0.05~0.08mm 측정',
     '❌ 진원도 불량'),
]
for r,(a,b,c,d) in enumerate(summary,12):
    for ci,(val,bg) in enumerate(zip([a,b,c,d],
                ['F2F2F2','FFF9F0','FFE0E0','FFE0E0']),2):
        ap(ws2,r,ci, cell(val,bg=bg,bold=('❌' in val or '⚠' in val),wrap=True))
    rh(ws2,{r:50})


# ═══════════════════════════════════════════════════════════════════
# Sheet 3 — 개선 대책 Rev.2
# ═══════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet('3.개선대책_Rev2')
widths(ws3, [3,14,12,30,20,10,12,12])
ws3.merge_cells('B1:I1')
ap(ws3,1,2, hdr('■ 개선 대책 Rev.2 — 동영상 검증 결과 반영',bg='1F4E79',sz=13))
rh(ws3,{1:36})

ws3.merge_cells('B2:I2')
ap(ws3,2,2, hdr(
    '【주요 개정】 ① 유압 척 교체→3죠우 유압 척으로 수정  '
    '② 정삭 유압압력 감소 최우선  ③ 척 해제 후 CMM 전수측정 추가',
    bg='C00000',sz=10))
rh(ws3,{2:28})

for i,h in enumerate(['우선','단계','대책명','세부 내용 (개정)','기대 효과','비용','기간','담당'],2):
    ap(ws3,3,i, hdr(h,bg='2E75B6',sz=10))
rh(ws3,{3:22})

actions=[
    ('★★★','즉시\n(당일)','정삭 유압압력\n50% 감소\n【신규·최우선】',
     '• 동영상 확인: 현재 황삭=정삭 동일 유압 클램핑\n'
     '• 유압 척 압력 조절 밸브로 정삭 시 압력 50% 감소\n'
     '  예) 황삭 40bar → 정삭 20bar 이하\n'
     '• 유압 척 특성상 압력 조절만으로 즉시 적용 가능\n'
     '• 클램핑력 감소 → 탄성변형량 감소 → 복원량 감소',
     'Y방향 편차\n40~60% 감소','없음','즉시','생산팀',
     'FFE0E0'),
    ('★★★','즉시\n(당일)','척 해제 후\nCMM 전수 측정\n【신규·최우선】',
     '• 동영상 확인: 현재 척 해제 후 CMM 측정 없음\n'
     '• 모든 완성품 척 해제 후 X·Y·45°·135° 4방향 CMM 측정\n'
     '• 탄성 복원량(ΔY) 기록 → 가공 보정값 산출\n'
     '• 내경 + 진원도 동시 판정 (진원도 ≤0.03mm)',
     '불량 유출\n즉시 차단','없음','즉시','품질팀',
     'FFE0E0'),
    ('★★★','즉시\n(당일)','스프링 패스\n추가',
     '• 정삭 후 절삭깊이 0으로 공구 1회 재통과\n'
     '• 가공 중 탄성변형 잔류 재료 제거\n'
     '• 진원도 추가 개선 효과 (0.01~0.02mm 개선)',
     '진원도 개선\n0.03mm 이하','없음','즉시','생산팀',
     'FFE0E0'),
    ('★★','즉시\n(당일)','CNC 보정\n가공 적용',
     '• 정삭 후 척 해제 → CMM으로 Y방향 복원량(ΔY) 측정\n'
     '• CNC 프로그램 정삭 목표치 = 공칭 - ΔY 로 보정\n'
     '  예) ΔY=0.04mm 측정 시 → 정삭 지름 -0.04mm 보정\n'
     '• 복원량이 안정화될 때까지 매 초물 측정 후 보정',
     '공차 만족률\n즉시 향상','없음','즉시','생산기술',
     'FFF2CC'),
    ('★★','1~2주','전용 소프트 조우\n제작 적용',
     '• 소재 OD와 동일 반경으로 조우 내면 정밀 보링\n'
     '• 접촉각 120° 이상 → 단위압력 분산 → 변형 감소\n'
     '• 유압 척에 소프트 조우 체결하여 적용\n'
     '• 동영상 척킹 위치 확인 후 편심 보정 반영',
     '변형량\n30~50% 감소','소','1~2주','생산기술',
     'FFF2CC'),
    ('★★★','1개월','3죠우 유압 척\n교체\n【기존 대책 수정】',
     '• 기존 대책의 "유압 척 도입" → 이미 유압 척 사용 중이므로 수정\n'
     '• 현재: 2죠우 유압 척 → 변경: 3죠우 유압 척\n'
     '• 120° 균등 3방향 클램핑 → Y방향 단독 압축 해소\n'
     '• 유압 방식 유지로 반복 정밀도 확보\n'
     '• 타원 발생 근본 원인 완전 해소',
     '타원 불량\n근본 해소\nCpk≥1.33','중','1개월','생산기술',
     'E2EFDA'),
    ('★','3개월','콜렛 척 도입\n검토',
     '• 360° 균등 면 클램핑으로 최고 진원도 확보\n'
     '• OD 치수 편차 ±0.1mm 이내 소재 관리 필요\n'
     '• 3죠우 교체 후 Cpk 미달 시 검토',
     '최고 정밀도','대','3개월','생산기술',
     'DEEAF1'),
]
for r,(pri,stage,name,desc,eff,cost,period,dept,bg) in enumerate(actions,4):
    ap(ws3,r,2, cell(pri,bg='C00000' if '★★★' in pri else '2E75B6' if '★★' in pri else 'BDD7EE',
                     fg='FFFFFF' if '★' in pri else '000000',bold=True))
    ap(ws3,r,3, cell(stage,bg='1F4E79',fg='FFFFFF',bold=True))
    ap(ws3,r,4, cell(name,bg=bg,bold=True))
    ap(ws3,r,5, cell(desc,bg='FFFFFF',align='left'))
    ap(ws3,r,6, cell(eff,bg=bg,bold=True))
    ap(ws3,r,7, cell(cost,bg='E2EFDA' if cost=='없음' else 'FFE0E0' if cost=='대' else 'FFFFFF'))
    ap(ws3,r,8, cell(period))
    ap(ws3,r,9, cell(dept))
    rh(ws3,{r:75})


# ═══════════════════════════════════════════════════════════════════
# Sheet 4 — 절삭조건 / 유압압력 기준
# ═══════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet('4.유압압력_절삭조건')
widths(ws4, [3,22,20,22,20,16])
ws4.merge_cells('B1:G1')
ap(ws4,1,2, hdr('■ 유압 척 압력 관리 기준 및 절삭 조건',bg='1F4E79',sz=13))
rh(ws4,{1:36})

ws4.merge_cells('B2:G2')
ap(ws4,2,2, hdr('[ 유압 척 압력 설정 기준 — 핵심 개선 (동영상 확인: 황삭=정삭 동일 → 즉시 변경) ]',
                bg='C00000',sz=11))
rh(ws4,{2:28})

for i,h in enumerate(['항목','현행 (문제)','황삭 기준','정삭 기준 (개선)','비고'],2):
    ap(ws4,3,i, hdr(h,bg='2E75B6',sz=10))
rh(ws4,{3:22})

pressure=[
    ('유압 클램핑 압력','황삭=정삭 동일\n(동영상 확인)','40bar (예시)\n최대 파지력','20bar 이하\n(황삭의 50% 이하)','★ 최우선 개선\n압력조절밸브로 즉시 적용'),
    ('클램핑 목적','소재 고정','재료 제거 중 고정','탄성변형 최소화\n+ 소재 고정','정삭 시 저압 필수'),
    ('탄성 변형량','최대 (불량)','최대 허용','목표: 50% 이하\n감소','CMM으로 ΔY 확인'),
    ('척 해제 후\n복원량(ΔY)','미측정\n(동영상 확인)','—','측정·기록 필수\n목표: ΔY≤0.02mm','초물 매 로트 측정'),
    ('스프링 패스','미확인\n(동영상 불명)','미적용','정삭 후 1회 필수','절삭깊이 0으로 통과'),
]
for r,(item,curr,rough,finish,note) in enumerate(pressure,4):
    ap(ws4,r,2, cell(item,bg='F2F2F2',bold=True))
    ap(ws4,r,3, cell(curr,bg='FFE0E0',bold=True))
    ap(ws4,r,4, cell(rough,bg='FFFFFF'))
    ap(ws4,r,5, cell(finish,bg='E2EFDA',bold=True))
    ap(ws4,r,6, cell(note,bg='FFF2CC',align='left'))
    rh(ws4,{r:40})

rh(ws4,{9:14})
ws4.merge_cells('B10:G10')
ap(ws4,10,2, hdr('[ 절삭 조건 최적화 기준 ]',bg='2E75B6',sz=11))
rh(ws4,{10:26})

for i,h in enumerate(['항목','현행','황삭','정삭','스프링 패스'],2):
    ap(ws4,11,i, hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
rh(ws4,{11:22})

cutting=[
    ('유압 압력','황삭=정삭 동일','40bar (정상)','20bar 이하','—'),
    ('절삭 깊이','—','0.3~0.5mm','0.05~0.10mm','0mm (공구만 통과)'),
    ('이송 속도','—','정상','감속 적용','최저 이송'),
    ('주축 회전수','—','정상','동일 또는 증속','동일'),
    ('절삭유 공급','—','충분','충분','충분'),
    ('통과 횟수','—','1회','1회','1회 추가 필수'),
]
for r,row in enumerate(cutting,12):
    bgs=['F2F2F2','FFE0E0','FFFFFF','E2EFDA','DEEAF1']
    for ci,(val,bg) in enumerate(zip(row,bgs),2):
        ap(ws4,r,ci, cell(val,bg=bg,bold=(ci==2)))
    rh(ws4,{r:28})


# ═══════════════════════════════════════════════════════════════════
# Sheet 5 — CMM 측정 기준
# ═══════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet('5.CMM측정기준')
widths(ws5, [3,20,22,22,16,14])
ws5.merge_cells('B1:G1')
ap(ws5,1,2, hdr('■ CMM 측정 관리 기준 (동영상 검증: 척 해제 후 측정 없음 → 즉시 추가)',
                bg='1F4E79',sz=12))
rh(ws5,{1:36})

ws5.merge_cells('B2:G2')
ap(ws5,2,2, hdr('척 해제 후 CMM 전수 측정 미실시 확인 → 탄성복원량 파악 불가 → 즉시 측정 체계 수립',
                bg='C00000',sz=10))
rh(ws5,{2:26})

# 측정 흐름
ws5.merge_cells('B3:G3')
ap(ws5,3,2, hdr('[ 개선 측정 프로세스 ]',bg='2E75B6',sz=11))
rh(ws5,{3:26})

flows=[
    ('①','소재 투입','척킹 편심 여부 육안 확인\n(동영상 항목8 현장 재확인)','생산','—'),
    ('②','황삭 가공','정상 유압압력(40bar) 클램핑','생산','—'),
    ('③','공정 내 측정①','황삭 후 내경 확인 (공차여유 0.1mm 이상)','품질','매 20개'),
    ('④','정삭 클램핑','유압압력 50% 감소 (20bar 이하) 재클램핑','생산','★필수'),
    ('⑤','정삭 가공','절삭깊이 0.05~0.10mm','생산','—'),
    ('⑥','스프링 패스','절삭깊이 0으로 1회 추가 통과','생산','★필수'),
    ('⑦','척 해제\n【동영상 미실시 확인】','척 해제 후 CMM 측정 즉시 실시\n(기존 미실시 → 즉시 추가)','품질','★전수'),
    ('⑧','CMM 4방향 측정','0°/45°/90°/135° 4방향 내경 측정\n진원도 = 최대-최소 내경','품질','전수'),
    ('⑨','탄성복원량 기록','ΔY = CMM Y방향 - 목표치 기록\n보정값으로 CNC 피드백','품질','전수'),
    ('⑩','합·불 판정','내경 +0.05/0 AND 진원도 ≤0.03mm','품질','전수'),
]
for i,h in enumerate(['순서','단계','세부 내용','담당','빈도'],2):
    ap(ws5,4,i, hdr(h,bg='2E75B6',sz=10))
rh(ws5,{4:22})
for r,(no,stage,desc,dept,freq) in enumerate(flows,5):
    bg='FFE0E0' if '★' in freq or '★' in desc else 'FFFFFF'
    ap(ws5,r,2, cell(no,bg='1F4E79',fg='FFFFFF',bold=True))
    ap(ws5,r,3, cell(stage,bg='F2F2F2',bold=True))
    ap(ws5,r,4, cell(desc,bg=bg,align='left',bold='동영상' in stage))
    ap(ws5,r,5, cell(dept,bg='FFFFFF'))
    ap(ws5,r,6, cell(freq,bg='FFE0E0' if '★' in freq else 'E2EFDA' if freq=='전수' else 'FFFFFF',
                     bold='★' in freq))
    rh(ws5,{r:38})

rh(ws5,{15:14})
ws5.merge_cells('B16:G16')
ap(ws5,16,2, hdr('[ CMM 판정 기준 ]',bg='2E75B6',sz=11))
rh(ws5,{16:26})
for i,h in enumerate(['측정항목','측정방향','현행 기준','개선 기준','판정'],2):
    ap(ws5,17,i, hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
rh(ws5,{17:22})
criteria=[
    ('내경 치수','0°/45°/90°/135°','1방향만','4방향 전수','전방향 +0.05/0 만족'),
    ('진원도','최대-최소 내경 차','미관리','≤ 0.03mm','초과 시 즉시 불합격'),
    ('탄성복원량 ΔY','Y방향 기준','미측정','≤ 0.02mm 목표\n초과 시 CNC 보정','로트 초물 전수 기록'),
    ('척킹 편심','육안 확인','미실시','매 척킹 시 확인','편심 시 재척킹'),
]
for r,row in enumerate(criteria,18):
    bgs=['F2F2F2','FFFFFF','FFE0E0','E2EFDA','DEEAF1']
    for ci,(val,bg) in enumerate(zip(row,bgs),2):
        ap(ws5,r,ci, cell(val,bg=bg,bold=(ci==2 or ci==6),wrap=True))
    rh(ws5,{r:36})


# ═══════════════════════════════════════════════════════════════════
# Sheet 6 — 실행계획
# ═══════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet('6.실행계획_효과')
widths(ws6, [3,22,14,12,12,12,12,12,12])
ws6.merge_cells('B1:J1')
ap(ws6,1,2, hdr('■ 실행 계획 및 효과 확인 (Rev.2)',bg='1F4E79',sz=13))
rh(ws6,{1:36})

ws6.merge_cells('B2:J2')
ap(ws6,2,2, hdr('[ 주차별 실행 계획 ]',bg='2E75B6',sz=11))
rh(ws6,{2:26})

for i,h in enumerate(['대책 항목','담당','1주','2주','3주','4주','2개월','3개월','비고'],2):
    ap(ws6,3,i, hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
rh(ws6,{3:22})

gantt=[
    ('정삭 유압압력 50% 감소 적용','생산팀',  '●','●','●','●','●','●','★즉시'),
    ('척 해제 후 CMM 전수 측정 추가','품질팀', '●','●','●','●','●','●','★즉시'),
    ('스프링 패스 추가','생산팀',              '●','●','●','●','●','●','★즉시'),
    ('CNC 보정값 산출·적용','생산기술',        '●','●','●','●','●','●',''),
    ('재공품 전수 CMM 선별','품질팀',          '●','','','','','',''),
    ('소프트 조우 제작·적용','생산기술',        '○','●','●','','','',''),
    ('척킹 위치 편심 재확인','생산팀',         '●','●','','','','','동영상 항목8'),
    ('3죠우 유압 척 발주·교체','생산기술',     '','','○','●','','',''),
    ('작업표준·FMEA 개정','품질팀',           '','○','●','','','',''),
    ('CMM 측정 기준서 개정','품질팀',         '●','●','','','','',''),
    ('Cpk 관리체계 구축','품질팀',            '','','','','○','●',''),
]
for r,(item,dept,*weeks) in enumerate(gantt,4):
    ap(ws6,r,2, cell(item,bg='F2F2F2' if '★' not in weeks[-1] else 'FFE0E0',
                     bold='★' in (weeks[-1] if weeks else '')))
    ap(ws6,r,3, cell(dept))
    for ci,(w,col) in enumerate(zip(weeks[:6],range(4,10))):
        bg='E2EFDA' if w=='●' else 'FFF2CC' if w=='○' else 'FFFFFF'
        ap(ws6,r,col, cell(w,bg=bg,bold=w in ['●','○']))
    ap(ws6,r,10, cell(weeks[-1] if len(weeks)==8 else '',
                      bg='FFE0E0' if weeks and '★' in str(weeks[-1]) else 'FFFFFF',
                      bold=True))
    rh(ws6,{r:22})

rh(ws6,{15:14})
ws6.merge_cells('B16:J16')
ap(ws6,16,2, hdr('[ 개선 효과 확인 지표 ]',bg='2E75B6',sz=11))
rh(ws6,{16:26})
for i,h in enumerate(['지표','현행','즉시 대책 후','소프트조우','3죠우 교체','목표','확인방법'],2):
    ap(ws6,17,i, hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
rh(ws6,{17:22})
effects=[
    ('Y방향 내경 편차','+0.05 초과','+0.02~0.03','+0.01~0.02','0.01 이하','+0.05/0 내','CMM 전수'),
    ('진원도 편차','0.05~0.08mm','0.03~0.05mm','0.02~0.03mm','0.01mm 이하','≤0.03mm','CMM 전수'),
    ('내경 합격률','불량 다수','80% 이상','93% 이상','99% 이상','100%','CMM'),
    ('공정능력 Cpk','< 1.0','≥ 1.0','≥ 1.17','≥ 1.33','≥ 1.33','SPC'),
    ('탄성복원량 ΔY','미측정','≤0.03mm','≤0.02mm','≤0.01mm','≤0.02mm','CMM'),
]
for r,row in enumerate(effects,18):
    bgs=['F2F2F2','FFE0E0','FFF2CC','FFF2CC','E2EFDA','DEEAF1','FFFFFF']
    for ci,(val,bg) in enumerate(zip(row,bgs),2):
        ap(ws6,r,ci, cell(val,bg=bg,bold=(ci==2),wrap=True))
    rh(ws6,{r:28})


# ═══════════════════════════════════════════════════════════════════
# Sheet 7 — 재발방지 / 승인
# ═══════════════════════════════════════════════════════════════════
ws7 = wb.create_sheet('7.재발방지_승인')
widths(ws7, [3,22,32,14,12])
ws7.merge_cells('B1:F1')
ap(ws7,1,2, hdr('■ 재발 방지 대책 및 결론 (Rev.2)',bg='1F4E79',sz=13))
rh(ws7,{1:36})

ws7.merge_cells('B2:F2')
ap(ws7,2,2, hdr('[ 동영상 검증 기반 최종 결론 ]',bg='C00000',sz=11))
rh(ws7,{2:26})

concl=[
    '확정 원인: 2죠우 유압 척이 Y방향(수직)으로 소재를 압축 클램핑',
    '변형 기전: Y방향 압축 → 가공 후 척 해제 시 탄성 복원 → Y방향 내경 증가(팽창) → +0.05 초과',
    '가중 원인①: 정삭 시에도 황삭과 동일 유압압력 적용 → 탄성변형 최대 → 복원량 최대 (동영상 확인)',
    '가중 원인②: 척 해제 후 CMM 측정 없음 → 탄성복원 불량 검출 불가 (동영상 확인)',
    '대책 수정: 기존 "유압 척 교체" → "3죠우 유압 척 교체"로 정정 (이미 유압 척 사용 중)',
    '근본 해결: 3죠우 유압 척으로 교체 → 120° 균등 클램핑 → 방향성 탄성변형 해소',
]
bg_map={'확정':'FFE0E0','변형':'FFE0E0','가중 원인①':'FFE0E0','가중 원인②':'FFE0E0',
        '대책 수정':'FFF2CC','근본':'E2EFDA'}
for r,text in enumerate(concl,3):
    ws7.merge_cells(f'B{r}:F{r}')
    key=text[:4]
    bg=next((v for k,v in bg_map.items() if text.startswith(k)),'FFF9F0')
    ap(ws7,r,2, cell(text,bg=bg,bold=True,align='left',sz=11))
    rh(ws7,{r:26})

rh(ws7,{9:14})
ws7.merge_cells('B10:F10')
ap(ws7,10,2, hdr('[ 재발 방지 관리 계획 ]',bg='2E75B6',sz=11))
rh(ws7,{10:26})
for i,h in enumerate(['관리 항목','개정 내용','시점','담당'],2):
    ap(ws7,11,i, hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
rh(ws7,{11:22})
prevent=[
    ('작업표준서',
     '정삭 유압압력 기준값 명기 (황삭 50% 이하)\n'
     '2단 클램핑법(황삭/정삭 압력 분리) 절차 추가\n'
     '스프링 패스 필수 실시 명기\n척킹 편심 확인 절차 추가','1주 이내','생산기술'),
    ('검사기준서',
     'CMM 4방향 측정 필수 명기\n척 해제 후 측정 필수 추가\n'
     '진원도 ≤0.03mm 판정 기준 신설\n탄성복원량 ΔY 기록 양식 추가','1주 이내','품질팀'),
    ('공정 FMEA',
     '2죠우 유압 척 탄성변형 고장모드 추가\n'
     '정삭 클램핑력 동일 → 복원량 최대 고장모드 추가\n'
     '검출 대책: CMM 척 해제 후 전수 측정','2주 이내','품질팀'),
    ('유압 설비 관리',
     '정삭 유압 압력 설정값 라벨 부착\n압력 조절밸브 봉인(임의 변경 방지)\n유압압력 일상 점검 항목 추가','1주 이내','설비팀'),
    ('수평 전개',
     '2죠우 척 사용 전 품종 동일 검토 실시\n유압압력 정삭/황삭 구분 여부 전수 확인','1개월','품질팀'),
]
for r,row in enumerate(prevent,12):
    bgs=['F2F2F2','FFFFFF','FFF2CC','FFFFFF']
    for ci,(val,bg) in enumerate(zip(row,bgs),2):
        ap(ws7,r,ci, cell(val,bg=bg,bold=(ci==2),align='left',wrap=True))
    rh(ws7,{r:55})

rh(ws7,{17:14})
ws7.merge_cells('B18:F18')
ap(ws7,18,2, hdr('[ 검토 및 승인 ]',bg='1F4E79',sz=11))
rh(ws7,{18:26})
for i,h in enumerate(['구분','작성','검토','승인','비고'],2):
    ap(ws7,19,i, hdr(h,bg='2E75B6',sz=10))
rh(ws7,{19:22})
for r,label in enumerate(['성명','서명','일자'],20):
    for ci,val in enumerate([label,'','','',
                              '2026.05.15' if label=='일자' else ''],2):
        ap(ws7,r,ci, cell(val,bg='F2F2F2' if ci==2 else 'FFFFFF',bold=(ci==2)))
    rh(ws7,{r:28})

# ── 저장 ─────────────────────────────────────────────────────────
path='/home/user/quality-analysis/GKO1808_소켓내경타원발생_개선대책서_Rev2_동영상검증반영.xlsx'
wb.save(path)
print(f'저장 완료: {path}')
