import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '2죠우척_변형개선'

# 열 너비
cols = [2, 14, 16, 16, 16, 16, 16, 14]
for i, w in enumerate(cols, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

def bd(t='thin'):
    s = Side(style=t)
    return Border(left=s, right=s, top=s, bottom=s)

def ap(ws, r, c, val, bg=None, fg='000000', sz=10, bold=False,
       align='center', wrap=True, merge=None, border=True, num=None, italic=False):
    if merge:
        ws.merge_cells(merge)
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = Font(name='맑은 고딕', size=sz, bold=bold, color=fg, italic=italic)
    cell.fill = PatternFill('solid', fgColor=bg) if bg else PatternFill()
    cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    if border:
        cell.border = bd()
    if num:
        cell.number_format = num
    return cell

def rh(r, h): ws.row_dimensions[r].height = h

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 타이틀
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ap(ws,1,2,'OUTER TIE ROD END — 구중심높이 불량 / 2죠우척 탄성변형 개선대책',
   bg='1F4E79',fg='FFFFFF',sz=13,bold=True,merge='B1:H1'); rh(1,36)
ap(ws,2,2,'품번: GKO1808계열  재질: S45C Q&T  척: 2죠우 유압척  작성: 2026-05-15',
   bg='2E75B6',fg='FFFFFF',sz=10,merge='B2:H2'); rh(2,22)

rh(3,8)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ① 불량 현황
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ap(ws,4,2,'① 불량 현황',bg='C00000',fg='FFFFFF',sz=11,bold=True,merge='B4:H4'); rh(4,26)

# 헤더
for c,h in zip(range(2,9),['항목','SPEC','실측값','편차','공차폭','초과량','판정']):
    ap(ws,5,c,h,bg='404040',fg='FFFFFF',bold=True,sz=10); rh(5,22)

rows5 = [
    ('구중심높이','11.07 +0.07/0\n(11.07~11.14)','11.34 mm','+0.20 mm','0.07 mm','+0.20 mm (약3배)','❌ NG'),
    ('ø31.5 내경','ø31.5 +0.15/0','—','—','0.15 mm','—','측정 필요'),
    ('ø26.08 내경','ø26.08 +0.07/0','—','—','0.07 mm','—','측정 필요'),
    ('진원도 ø0.8','ø0.8 A 이하','—','—','—','—','측정 필요'),
]
bgs6=['FFE0E0','FFF2CC','FFF2CC','FFF2CC']
for r,(a,b,c,d,e,f,g) in enumerate(rows5,6):
    bg=bgs6[r-6]
    for ci,v in enumerate([a,b,c,d,e,f,g],2):
        bold = (ci==8 and '❌' in str(v))
        fg2 = 'C00000' if '❌' in str(v) else '000000'
        ap(ws,r,ci,v,bg=bg,bold=bold,fg=fg2,sz=10,align='center' if ci!=3 else 'center'); rh(r,32)

rh(12,8)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ② 변형 메커니즘
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ap(ws,13,2,'② 2죠우척 탄성변형 → 구중심높이 증가 메커니즘',bg='C00000',fg='FFFFFF',sz=11,bold=True,merge='B13:H13'); rh(13,26)

mech = [
    ('STEP 1','2죠우척 Y방향(수직) 클램핑',
     '상·하 조우가 소켓 OD를 Y방향으로 압축\n→ 소켓이 Y방향으로 납작해지며 X방향 팽창',
     '탄성 압축 발생','1F4E79','FFFFFF'),
    ('STEP 2','가공 중 변형 상태 유지',
     '압축 변형 상태에서 구면 보링·선삭 가공\n→ 공구는 변형된 형상 기준으로 절삭\n※ S45C Q&T: 탄성계수 높아 복원력 큼',
     '변형 기준 가공','2E75B6','FFFFFF'),
    ('STEP 3','척 해제 → 탄성 복원',
     'Y방향 압축 해제 → 소켓 높이 방향 팽창\n→ 구중심높이 증가\n실측: 11.07 기준 → 11.34 (+0.27mm 가공 中 변형)',
     '구중심높이 증가\n+0.20mm 초과','C00000','FFFFFF'),
    ('가중 요인','S45C Q&T 재질 특성',
     '• Q&T 후 항복강도 高 → 탄성 변형량 ↑\n• 열처리 잔류응력 → 가공 후 추가 변형 가능\n• 2죠우 구조 + 고강도재 조합 → 변형 최대화',
     '복원량 최대','C55A11','FFFFFF'),
]
for c,h in zip(range(2,7),['단계','현상','상세 내용','영향','']):
    ap(ws,14,c,h,bg='404040',fg='FFFFFF',bold=True); rh(14,22)
ws.merge_cells('F14:H14')
ap(ws,14,6,'영향',bg='404040',fg='FFFFFF',bold=True)

for r,(step,phen,desc,eff,hbg,hfg) in enumerate(mech,15):
    ap(ws,r,2,step,bg=hbg,fg=hfg,bold=True,sz=9)
    ap(ws,r,3,phen,bg='F2F2F2',bold=True,sz=10)
    ws.merge_cells(f'D{r}:F{r}')
    ap(ws,r,4,desc,bg='FFFFFF',align='left',sz=10)
    ws.merge_cells(f'G{r}:H{r}')
    ap(ws,r,7,eff,bg='FFE0E0' if hbg=='C00000' else 'FFF2CC',bold=True,sz=10)
    rh(r,52)

rh(19,8)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ③ 변형량 추정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ap(ws,20,2,'③ 변형량 추정 및 CNC 보정 기준',bg='375623',fg='FFFFFF',sz=11,bold=True,merge='B20:H20'); rh(20,26)

for c,h in zip(range(2,9),['항목','수식/근거','추정값','비고','','','']):
    if c<=4:
        ap(ws,21,c,h,bg='404040',fg='FFFFFF',bold=True)
    rh(21,22)
ws.merge_cells('E21:H21')
ap(ws,21,5,'',bg='404040')

calc = [
    ('실측 초과량','실측(11.34) - 공차상한(11.14)','+0.20 mm','즉각 대응 필요'),
    ('척 해제 복원량 추정','실측(11.34) - 정삭목표(11.07) = 클램핑 중 변형량','약 0.27 mm','S45C Q&T 고강도로 복원량 큼'),
    ('초기 CNC 보정값','정삭 목표 = 공칭 - 복원량\n= 11.07 - 0.20 = 10.87mm','10.87 mm (초기)','클램핑력 감소 후 재측정·재산출'),
    ('클램핑력 감소 후 재산출','정삭 유압 50% 감소 → 복원량 감소 예상','0.10~0.15mm 예상','초물 3개 가공 후 CMM 실측 확인'),
    ('목표 복원량','유압 감소 + 소프트조우 적용 후','≤ 0.05 mm','구중심높이 11.07~11.14 안정화'),
]
for r,(item,basis,val,note) in enumerate(calc,22):
    ap(ws,r,2,item,bg='F2F2F2',bold=True,sz=10)
    ws.merge_cells(f'C{r}:D{r}')
    ap(ws,r,3,basis,bg='FFFFFF',align='left',sz=10)
    ap(ws,r,5,val,bg='E2EFDA',bold=True,sz=10)
    ws.merge_cells(f'F{r}:H{r}')
    ap(ws,r,6,note,bg='FFFFFF',align='left',sz=10)
    rh(r,32)

rh(27,8)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ④ 개선 대책
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ap(ws,28,2,'④ 개선 대책 (우선순위 순)',bg='C00000',fg='FFFFFF',sz=11,bold=True,merge='B28:H28'); rh(28,26)

for c,h in zip(range(2,9),['순위','단계','대책','세부 내용','효과','비용','기간']):
    ap(ws,29,c,h,bg='404040',fg='FFFFFF',bold=True); rh(29,22)

acts = [
    ('★★★','즉시','정삭 유압압력\n50% 감소',
     '황삭: 정상 유압 유지\n정삭: 50% 이하로 압력 감소\n(예: 40bar→18~20bar)\n→ 탄성변형량 감소 → 복원량 감소',
     '복원량 50%↓','없음','즉시','FFE0E0'),
    ('★★★','즉시','CNC 보정\n정삭 목표 하향',
     '정삭 목표 높이 = 11.07 - 복원량\n초기 보정: 10.87mm 가공\n→ 척 해제 후 CMM 측정 → 보정값 확정\n※ 클램핑력 감소 후 재산출 필수',
     '구중심높이\n공차 내 진입','없음','즉시','FFE0E0'),
    ('★★★','즉시','척 해제 후\n구중심높이 측정',
     '현재 척 해제 후 측정 없음\n→ 정삭 후 척 해제 → CMM 측정 추가\n구중심높이 + ø내경 동시 측정\n복원량(ΔH) 기록 → 보정 피드백',
     '불량 유출\n즉시 차단','없음','즉시','FFE0E0'),
    ('★★','즉시','스프링 패스\n추가',
     '정삭 후 절삭깊이 0으로 1회 재통과\n→ 탄성변형 잔류재료 제거\n→ 진원도 + 높이 정밀도 동시 개선',
     '진원도 개선','없음','즉시','FFE0E0'),
    ('★★','1~2주','소프트 조우\n제작',
     'OD 곡률 맞춤 소프트 조우 제작\n접촉각 120° 이상 확보\n→ 단위 클램핑압 분산\n→ 탄성변형 30~50% 감소',
     '변형량 30~50%↓','소','1~2주','FFF2CC'),
    ('★★★','1개월','3죠우 유압척\n교체 【근본】',
     '2죠우 → 3죠우 유압척으로 교체\n120° 균등 3방향 클램핑\nY방향 단독 압축 해소\n→ 타원·높이 변형 근본 해결\n목표: 구중심높이 Cpk≥1.33',
     '근본 해결\nCpk≥1.33','중','1개월','E2EFDA'),
    ('★','검토','Q&T 공정 순서\n재검토',
     '현재: 가공 후 Q&T 또는 Q&T 후 가공?\n→ Q&T 후 가공 시 잔류응력 추가 변형 가능\n→ 가공 공정 전 응력 이완 검토\n(저온 어닐링 150~200℃)',
     '잔류응력↓','소~중','검토','DEEAF1'),
]
for r,(pri,stage,name,desc,eff,cost,period,bg) in enumerate(acts,30):
    pbg='C00000' if '★★★' in pri else '2E75B6' if '★★' in pri else 'BDD7EE'
    pfg='FFFFFF'
    ap(ws,r,2,pri,bg=pbg,fg=pfg,bold=True,sz=9)
    ap(ws,r,3,stage,bg='1F4E79',fg='FFFFFF',bold=True,sz=9)
    ap(ws,r,4,name,bg=bg,bold=True,sz=10)
    ap(ws,r,5,desc,bg='FFFFFF',align='left',sz=9)
    ap(ws,r,6,eff,bg=bg,bold=True,sz=10)
    ap(ws,r,7,cost,bg='E2EFDA' if cost=='없음' else 'FFF2CC',sz=10)
    ap(ws,r,8,period,sz=10)
    rh(r,60)

rh(37,8)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ⑤ 측정 관리 기준
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ap(ws,38,2,'⑤ 측정·관리 기준 (개선)',bg='375623',fg='FFFFFF',sz=11,bold=True,merge='B38:H38'); rh(38,26)

for c,h in zip(range(2,9),['항목','현행','개선','판정 기준','빈도','','담당']):
    if c not in [7]:
        ap(ws,39,c,h,bg='404040',fg='FFFFFF',bold=True)
    rh(39,22)
ws.merge_cells('G39:G39')

meas=[
    ('구중심높이','미측정 또는 샘플','척 해제 후 CMM 전수','11.07~11.14mm','전수','','품질팀'),
    ('복원량 ΔH','미관리','정삭 전·후 측정\nΔH = 측정 - 목표','≤ 0.05mm 목표','초물 매 로트','','품질팀'),
    ('ø31.5 내경','—','CMM 측정','31.5~31.65mm','전수','','품질팀'),
    ('ø26.08 내경','—','CMM 측정','26.08~26.15mm','전수','','품질팀'),
    ('진원도','—','CMM','ø0.8 A 이하','전수','','품질팀'),
    ('유압 압력','미관리','정삭/황삭 구분 기록','황삭 40bar\n정삭 20bar 이하','일상 점검','','설비팀'),
]
for r,row in enumerate(meas,40):
    bgs=['F2F2F2','FFE0E0','E2EFDA','DEEAF1','FFF2CC','FFFFFF','FFFFFF']
    for ci,(v,bg) in enumerate(zip(row,bgs),2):
        if ci==7: continue
        ap(ws,r,ci,v,bg=bg,bold=(ci==2),sz=9,align='center' if ci!=4 else 'center')
    ap(ws,r,7,row[6],bg='FFFFFF',sz=9)
    rh(r,32)

rh(46,8)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ⑥ 효과 예측
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ap(ws,47,2,'⑥ 단계별 개선 효과 예측',bg='1F4E79',fg='FFFFFF',sz=11,bold=True,merge='B47:H47'); rh(47,26)

for c,h in zip(range(2,9),['지표','현행','즉시 대책 후','소프트조우 후','3죠우 교체 후','목표','확인']):
    ap(ws,48,c,h,bg='404040',fg='FFFFFF',bold=True); rh(48,22)

effs=[
    ('구중심높이','11.34mm (NG)','11.15~11.20mm','11.10~11.16mm','11.07~11.13mm','11.07~11.14','CMM전수'),
    ('복원량 ΔH','0.27mm 추정','0.10~0.15mm','0.05~0.08mm','0.01~0.03mm','≤0.05mm','CMM'),
    ('구중심높이 Cpk','<0 (NG)','≈0.5','≈1.0','≥1.33','≥1.33','SPC'),
    ('X-Y 타원편차','0.05~0.08mm','0.03~0.05mm','0.02~0.03mm','≤0.01mm','≤0.03mm','CMM'),
    ('합격률','불량','70% 이상','90% 이상','99% 이상','100%','전수검사'),
]
bgs_e=['F2F2F2','FFE0E0','FFF2CC','FFF2CC','E2EFDA','DEEAF1','FFFFFF']
for r,row in enumerate(effs,49):
    for ci,(v,bg) in enumerate(zip(row,bgs_e),2):
        bold = ci==2 or (ci==3 and 'NG' in str(v))
        fg2='C00000' if 'NG' in str(v) and ci==3 else '000000'
        ap(ws,r,ci,v,bg=bg,bold=bold,fg=fg2,sz=10)
    rh(r,24)

rh(54,8)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 하단 결론
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ap(ws,55,2,
   '▶ 결론: 2죠우척 Y방향 압축 → 척 해제 시 탄성복원 → 구중심높이 증가 (+0.20mm 초과). '
   '즉시: 정삭 유압 50%↓ + CNC 보정(-0.20mm) + 척 해제 후 CMM 전수측정. '
   '근본: 3죠우 유압척 교체. S45C Q&T 고강도재 특성상 복원량이 크므로 보정값 정밀 관리 필수.',
   bg='1F4E79',fg='FFFFFF',sz=10,bold=True,merge='B55:H55',align='left'); rh(55,38)

path = '/home/user/quality-analysis/OUTER_TIE_ROD_END_구중심높이_2죠우척_변형개선.xlsx'
wb.save(path)
print(f'저장: {path}')
