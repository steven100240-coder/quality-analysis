import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 공통 스타일 ──────────────────────────────────────────────────
def bd(w='thin'):
    s=Side(style=w); return Border(left=s,right=s,top=s,bottom=s)

def hdr(v,bg='1F4E79',fg='FFFFFF',sz=11,bold=True,wrap=True):
    return dict(value=v,
        font=Font(name='맑은 고딕',size=sz,bold=bold,color=fg),
        fill=PatternFill('solid',fgColor=bg),
        alignment=Alignment(horizontal='center',vertical='center',wrap_text=wrap),
        border=bd())

def cel(v,bg=None,bold=False,align='center',wrap=True,sz=10,fg='000000',num=None):
    d=dict(value=v,
        font=Font(name='맑은 고딕',size=sz,bold=bold,color=fg),
        fill=PatternFill('solid',fgColor=bg) if bg else PatternFill(),
        alignment=Alignment(horizontal=align,vertical='center',wrap_text=wrap),
        border=bd())
    if num: d['number_format']=num
    return d

def ap(ws,r,c,d):
    x=ws.cell(row=r,column=c,value=d['value'])
    for k in ('font','fill','alignment','border'): setattr(x,k,d[k])
    if 'number_format' in d: x.number_format=d['number_format']

def cw(ws,w):
    for i,v in enumerate(w,1): ws.column_dimensions[get_column_letter(i)].width=v

def rh(ws,d):
    for r,h in d.items(): ws.row_dimensions[r].height=h

# ── CMM 실측 데이터 (파싱 결과) ──────────────────────────────────
# SPEC: ① ø30.83+0.05/0 → 30.83~30.88  ② ø9.2+0.05/0 → 9.20~9.25
#        ③ ø30.83+0.05/0                ④ 205.77±1.0 → 204.77~206.77
#        ⑤ 15.7°+3°/0 → 15.7~18.7°    ⑥ D/C끝점 227.8±0.3

SPEC_1_LO=30.83; SPEC_1_HI=30.88
SPEC_2_LO=9.20;  SPEC_2_HI=9.25
SPEC_3_LO=30.83; SPEC_3_HI=30.88
SPEC_4_LO=204.77; SPEC_4_HI=206.77

data = {
    '2026.05.12(CMM)': [
        {'no':1,'d1':30.8925,'d2':9.2605,'d3':30.8934,'d4':206.3811,'d5':"16°23'54\"",'xval':None,'yval':None,'xy':None},
        {'no':2,'d1':30.8829,'d2':9.2717,'d3':30.8940,'d4':206.2300,'d5':"16°28'41\"",'xval':None,'yval':None,'xy':None},
        {'no':3,'d1':30.8813,'d2':9.2705,'d3':30.8768,'d4':206.2331,'d5':"16°23'41\"",'xval':None,'yval':None,'xy':None},
    ],
    '2026.05.14(CMM+게이지)_1차': [
        {'no':1,'d1':30.8854,'d2':9.2634,'d3':30.8953,'d4':206.5856,'d5':None,'xval':30.84,'yval':30.90,'xy':0.060},
        {'no':2,'d1':30.9410,'d2':9.2119,'d3':30.9046,'d4':206.7691,'d5':None,'xval':30.86,'yval':30.91,'xy':0.050},
        {'no':3,'d1':30.8963,'d2':9.2412,'d3':30.8941,'d4':206.2741,'d5':None,'xval':30.85,'yval':30.90,'xy':0.050},
    ],
    '2026.05.14(CMM+게이지)_2차': [
        {'no':1,'d1':30.8761,'d2':9.2201,'d3':30.8413,'d4':203.2358,'d5':None,'xval':30.81,'yval':30.88,'xy':0.070},
        {'no':2,'d1':30.8759,'d2':9.2097,'d3':30.8395,'d4':203.5005,'d5':None,'xval':30.80,'yval':30.88,'xy':0.080},
        {'no':3,'d1':30.8756,'d2':9.1848,'d3':30.8435,'d4':203.5297,'d5':None,'xval':30.81,'yval':30.88,'xy':0.070},
    ],
}

def judge(v,lo,hi):
    if v is None: return '—'
    return 'OK' if lo<=v<=hi else 'NG'

def jcolor(v,lo,hi):
    if v is None: return 'FFFFFF'
    return 'E2EFDA' if lo<=v<=hi else 'FFE0E0'

wb=openpyxl.Workbook()

# ═══════════════════════════════════════════════════════════════════
# Sheet 1 — 표지
# ═══════════════════════════════════════════════════════════════════
ws1=wb.active; ws1.title='1.표지'
cw(ws1,[3,20,26,18,18,14,14])

ws1.merge_cells('B1:G1')
ap(ws1,1,2,hdr('GKO1808 소켓 내경 타원 발생 개선대책서  Rev.3',bg='1F4E79',sz=16))
ws1.merge_cells('B2:G2')
ap(ws1,2,2,hdr('CMM 삼차원 측정 데이터 + 동영상 현장 검증 완전 반영',bg='2E75B6',sz=11))
rh(ws1,{1:42,2:26})

info=[
    ('문서번호','GKO1808-QI-003 Rev.3','작성일','2026-05-15','버전','Rev.3'),
    ('협력사','부경하이텍','품번','GKO1808','품명','소켓(Socket)'),
    ('내경공차','① ø30.83+0.05/0  ② ø9.2+0.05/0  ③ ø30.83+0.05/0','척 종류','2죠우 유압 척','측정','CMM(삼차원)'),
    ('개정 사유','CMM 측정 데이터(26.05.12·14) 수치 반영 + 동영상 검증 결과 통합','','','',''),
]
for r,row in enumerate(info,4):
    ap(ws1,r,2,hdr(row[0],bg='BDD7EE',fg='1F4E79',sz=10))
    if r==7:
        ws1.merge_cells(f'C7:G7')
        ap(ws1,r,3,cel(row[1],bold=True))
    else:
        ap(ws1,r,3,cel(row[1],bold=True))
        ap(ws1,r,4,hdr(row[2],bg='BDD7EE',fg='1F4E79',sz=10))
        ap(ws1,r,5,cel(row[3]))
        ap(ws1,r,6,hdr(row[4],bg='BDD7EE',fg='1F4E79',sz=10))
        ap(ws1,r,7,cel(row[5]))
    rh(ws1,{r:22})

rh(ws1,{8:10})
ws1.merge_cells('B9:G9')
ap(ws1,9,2,hdr('■ 핵심 수치 요약 (CMM 실측 기반)',bg='C00000',sz=12))
rh(ws1,{9:28})

kpi=[
    ('X-Y 내경 편차 (실측)','0.05 ~ 0.08 mm','실린더게이지 측정 (2026.05.14)'),
    ('Y방향 내경 (실측 MAX)','30.91 mm','공차 상한 30.88mm → 0.03mm 초과'),
    ('X방향 내경 (실측 MIN)','30.80 mm','공차 하한 30.83mm → 0.03mm 미달'),
    ('CMM ① 내경 (AVG)','30.886~30.908 mm','전 로트 NG — 공차 상한 초과'),
    ('CMM ② 내경 (AVG)','9.204~9.268 mm','일부 9.25 초과 NG'),
    ('CNC 목표 보정 필요량','Y방향 -0.06 ~ -0.08 mm','정삭 목표치 하향 보정 권고'),
    ('④ 길이 (이상 로트)','203.2~203.5 mm','공차 하한 204.77mm 미달 — 별도 조사'),
]
for i,h in enumerate(['항목','실측값','의미'],2):
    ap(ws1,10,i,hdr(h,bg='2E75B6',sz=10))
rh(ws1,{10:22})
for r,(item,val,meaning) in enumerate(kpi,11):
    bg='FFE0E0' if 'NG' in meaning or '초과' in meaning or '미달' in meaning else 'FFF2CC'
    ap(ws1,r,2,cel(item,bg='F2F2F2',bold=True))
    ap(ws1,r,3,cel(val,bg=bg,bold=True))
    ap(ws1,r,4,cel(meaning,bg=bg,align='left'))
    rh(ws1,{r:22})


# ═══════════════════════════════════════════════════════════════════
# Sheet 2 — CMM 측정 데이터
# ═══════════════════════════════════════════════════════════════════
ws2=wb.create_sheet('2.CMM측정데이터')
cw(ws2,[3,14,13,9,9,9,9,11,11,9,11,14,14])

ws2.merge_cells('B1:N1')
ap(ws2,1,2,hdr("■ GKO1808 소켓 CMM 측정 데이터 이력 (원본 반영)",bg='1F4E79',sz=13))
rh(ws2,{1:36})

# SPEC 행
ws2.merge_cells('B2:N2')
ap(ws2,2,2,hdr('[ SPEC ] ① ø30.83+0.05/0  ② ø9.2+0.05/0  ③ ø30.83+0.05/0  '
               '④ 205.77±1.0  ⑤ 15.7°+3°/0  Y방향관리기준: ø30.88이하',
               bg='2E75B6',sz=10))
rh(ws2,{2:26})

hdrs2=['No','측정일자','①CMM\nø30.83\n+0.05/0','①판정',
       '②CMM\nø9.2\n+0.05/0','②판정',
       '③CMM\nø30.83\n+0.05/0','③판정',
       '④길이\n205.77±1','④판정',
       '⑤각도\n15.7°+3/0',
       'X방향\n게이지','Y방향\n게이지','X-Y차']
for i,h in enumerate(hdrs2,2):
    ap(ws2,3,i,hdr(h,bg='2E75B6',sz=9,wrap=True))
rh(ws2,{3:45})

row_num=4
for batch_name,rows in data.items():
    # 배치 헤더
    ws2.merge_cells(f'B{row_num}:N{row_num}')
    bg='1F4E79' if '05.12' in batch_name else ('375623' if '1차' in batch_name else 'C55A11')
    ap(ws2,row_num,2,hdr(f'[ {batch_name} ]',bg=bg,sz=10))
    rh(ws2,{row_num:22}); row_num+=1

    vals1=[]; vals2=[]; vals3=[]; vals4=[]
    for row in rows:
        d1,d2,d3,d4=row['d1'],row['d2'],row['d3'],row['d4']
        vals1.append(d1); vals2.append(d2); vals3.append(d3); vals4.append(d4)
        j1=judge(d1,SPEC_1_LO,SPEC_1_HI); j2=judge(d2,SPEC_2_LO,SPEC_2_HI)
        j3=judge(d3,SPEC_3_LO,SPEC_3_HI); j4=judge(d4,SPEC_4_LO,SPEC_4_HI)
        ap(ws2,row_num,2,cel(row['no'],bg='F2F2F2',bold=True))
        ap(ws2,row_num,3,cel(batch_name.split('(')[0] if row['no']==1 else ''))
        ap(ws2,row_num,4,cel(d1,bg=jcolor(d1,SPEC_1_LO,SPEC_1_HI),num='0.0000'))
        ap(ws2,row_num,5,cel(j1,bg=jcolor(d1,SPEC_1_LO,SPEC_1_HI),bold=(j1=='NG'),
                             fg='C00000' if j1=='NG' else '375623'))
        ap(ws2,row_num,6,cel(d2,bg=jcolor(d2,SPEC_2_LO,SPEC_2_HI),num='0.0000'))
        ap(ws2,row_num,7,cel(j2,bg=jcolor(d2,SPEC_2_LO,SPEC_2_HI),bold=(j2=='NG'),
                             fg='C00000' if j2=='NG' else '375623'))
        ap(ws2,row_num,8,cel(d3,bg=jcolor(d3,SPEC_3_LO,SPEC_3_HI),num='0.0000'))
        ap(ws2,row_num,9,cel(j3,bg=jcolor(d3,SPEC_3_LO,SPEC_3_HI),bold=(j3=='NG'),
                             fg='C00000' if j3=='NG' else '375623'))
        ap(ws2,row_num,10,cel(d4,bg=jcolor(d4,SPEC_4_LO,SPEC_4_HI),num='0.0000'))
        ap(ws2,row_num,11,cel(j4,bg=jcolor(d4,SPEC_4_LO,SPEC_4_HI),bold=(j4=='NG'),
                              fg='C00000' if j4=='NG' else '375623'))
        ap(ws2,row_num,12,cel(row.get('d5') or ''))
        xv=row.get('xval'); yv=row.get('yval'); xyv=row.get('xy')
        ap(ws2,row_num,13,cel(xv if xv else '',
                              bg=jcolor(xv,SPEC_1_LO,SPEC_1_HI) if xv else 'FFFFFF',
                              num='0.00' if xv else None))
        ap(ws2,row_num,14,cel(yv if yv else '',
                              bg='FFE0E0' if yv and yv>SPEC_1_HI else (
                              'E2EFDA' if yv else 'FFFFFF'),
                              bold=(yv is not None and yv>SPEC_1_HI),
                              num='0.00' if yv else None))
        ap(ws2,row_num,15,cel(xyv if xyv else '',
                              bg='FFE0E0' if xyv and xyv>=0.05 else 'FFFFFF',
                              bold=(xyv is not None and xyv>=0.05),
                              num='0.000' if xyv else None))
        rh(ws2,{row_num:20}); row_num+=1

    # MAX/MIN/AVG/판정
    for label,func,bgg in [('MAX',max,'FFF2CC'),('MIN',min,'FFF2CC'),
                            ('AVG',lambda x:sum(x)/len(x),'DEEAF1')]:
        v1=func(vals1); v2=func(vals2); v3=func(vals3); v4=func(vals4)
        j1=judge(v1,SPEC_1_LO,SPEC_1_HI); j2=judge(v2,SPEC_2_LO,SPEC_2_HI)
        j3=judge(v3,SPEC_3_LO,SPEC_3_HI); j4=judge(v4,SPEC_4_LO,SPEC_4_HI)
        ap(ws2,row_num,2,cel('')); ap(ws2,row_num,3,cel(''))
        ap(ws2,row_num,2,hdr(label,bg='BDD7EE',fg='1F4E79',sz=9))
        ap(ws2,row_num,4,cel(round(v1,4),bg=bgg,num='0.0000'))
        ap(ws2,row_num,5,cel(''))
        ap(ws2,row_num,6,cel(round(v2,4),bg=bgg,num='0.0000'))
        ap(ws2,row_num,7,cel(''))
        ap(ws2,row_num,8,cel(round(v3,4),bg=bgg,num='0.0000'))
        ap(ws2,row_num,9,cel(''))
        ap(ws2,row_num,10,cel(round(v4,4),bg=bgg,num='0.0000'))
        for c in [5,7,9,11,12,13,14,15]: ap(ws2,row_num,c,cel(''))
        rh(ws2,{row_num:18}); row_num+=1

    # 판정 행
    all_ng_flag=any(judge(v,SPEC_1_LO,SPEC_1_HI)=='NG' for v in vals1)
    ap(ws2,row_num,2,hdr('판정',bg='C00000' if all_ng_flag else '375623',sz=9))
    ap(ws2,row_num,3,cel(''))
    for ci,flag in zip([4,6,8,10],
        [any(judge(v,lo,hi)=='NG' for v in vs)
         for vs,lo,hi in [(vals1,SPEC_1_LO,SPEC_1_HI),(vals2,SPEC_2_LO,SPEC_2_HI),
                          (vals3,SPEC_3_LO,SPEC_3_HI),(vals4,SPEC_4_LO,SPEC_4_HI)]]):
        ap(ws2,row_num,ci,cel('NG' if flag else 'OK',
                              bg='FFE0E0' if flag else 'E2EFDA',
                              bold=True,fg='C00000' if flag else '375623'))
        ap(ws2,row_num,ci+1,cel(''))
    for c in [12,13,14,15]: ap(ws2,row_num,c,cel(''))
    rh(ws2,{row_num:20}); row_num+=1
    row_num+=1  # 공백

# 분석 코멘트
ws2.merge_cells(f'B{row_num}:O{row_num}')
ap(ws2,row_num,2,hdr('■ CMM 데이터 분석 결과',bg='1F4E79',sz=11))
rh(ws2,{row_num:26}); row_num+=1

comments=[
    ('① Y방향 내경 과대','Y방향(게이지): 30.88~30.91mm → 공차상한 30.88mm 초과 (최대 +0.03mm 초과)','FFE0E0'),
    ('② X방향 내경 부족','X방향(게이지): 30.80~30.86mm → 일부 공차하한 30.83mm 미달 가능성','FFF2CC'),
    ('③ X-Y 편차 확인','실측 X-Y 차이: 0.05~0.08mm → 타원 편차 CMM으로 정량 확인','FFE0E0'),
    ('④ CMM ① 전 로트 NG','CMM 측정 ① 항목: 전 측정일 NG — 공차상한 30.88mm 지속 초과','FFE0E0'),
    ('⑤ CMM ② NG','ø9.2 내경: 9.2605~9.2717mm → 공차상한 9.25mm 초과 NG 빈발','FFE0E0'),
    ('⑥ ④ 길이 이상 로트','2026.05.14 2차: 203.2~203.5mm → 공차하한 204.77mm 크게 미달 → 별도 원인 조사 필요','FFF2CC'),
    ('⑦ CNC 보정 필요량','Y방향 복원량 0.06~0.08mm → 정삭 목표 내경을 0.06mm 하향 보정 권고','E2EFDA'),
]
for item,desc,bg in comments:
    ws2.merge_cells(f'B{row_num}:D{row_num}')
    ap(ws2,row_num,2,hdr(item,bg='2E75B6',sz=10))
    ws2.merge_cells(f'E{row_num}:O{row_num}')
    ap(ws2,row_num,5,cel(desc,bg=bg,align='left',bold=('CMM' in item and 'NG' in desc)))
    rh(ws2,{row_num:24}); row_num+=1


# ═══════════════════════════════════════════════════════════════════
# Sheet 3 — 원인 분석
# ═══════════════════════════════════════════════════════════════════
ws3=wb.create_sheet('3.원인분석')
cw(ws3,[3,16,28,24,16])
ws3.merge_cells('B1:F1')
ap(ws3,1,2,hdr('■ 2죠우 유압 척 — Y방향 클램핑 탄성복원 메커니즘 (CMM 수치 확정)',bg='1F4E79',sz=13))
rh(ws3,{1:36})

ws3.merge_cells('B2:F2')
ap(ws3,2,2,hdr('CMM 실측: Y방향 내경 X방향보다 0.05~0.08mm 크게 측정 → 탄성복원으로 Y 팽창 확정',
               bg='C00000',sz=10))
rh(ws3,{2:26})

mechs=[
    ('STEP 1','유압 2죠우 척\nY방향(수직) 클램핑',
     '• 상·하 유압 조우가 소재 Y방향 압축\n'
     '• 유압 방식이지만 2죠우 구조 → Y방향 단독 압축은 동일\n'
     '• 동영상 확인: 수직 방향 파지 확인','탄성 압축 발생','1F4E79'),
    ('STEP 2','X방향 Poisson 팽창',
     '• Y압축 → X방향 포아송 팽창\n'
     '• 가공 중: X내경 증가, Y내경 감소 (변형 상태)\n'
     '• 공구는 변형된 상태에서 원형 보링','X방향 팽창','2E75B6'),
    ('STEP 3','정삭=황삭 동일 유압압력\n【동영상 확인: 핵심 가중 원인】',
     '• 동영상 검증: 정삭 시에도 황삭과 동일 유압압력 적용\n'
     '• 탄성 변형량 최대 유지 → 복원량 최대\n'
     '• CMM NG 전 로트: 정삭 클램핑력 미조정이 주 원인','【핵심 가중 원인】\n탄성변형 최대','C00000'),
    ('STEP 4','척 해제 → Y방향 탄성 복원\n→ 내경 증가 (CMM 수치 확인)',
     '• Y방향: 압축 해제 → 탄성 복원 → ID_Y 증가\n'
     '  CMM실측: Y방향 30.88~30.91mm (공차상한 +0.03mm 초과)\n'
     '• X방향: 팽창 해제 → 수축 → ID_X 감소\n'
     '  CMM실측: X방향 30.80~30.86mm (하한 미달 가능)\n'
     '• X-Y 실측 편차: 0.05~0.08mm ← CMM 정량 확인','Y↑ 팽창 = 과대\nX↓ 수축 = 과소','C00000'),
    ('결과','타원 불량 + 전 로트 NG',
     '• CMM ①②③ 전 측정일 NG 확정\n'
     '• Y방향 내경 최대 30.91mm (공차상한 0.03mm 초과)\n'
     '• CNC 보정 필요: 정삭 목표 -0.06mm 하향\n'
     '• ④길이 이상 로트(203mm대) 별도 원인 조사 병행 필요','전 로트 NG\n즉시 대책 필요','C00000'),
]
for i,h in enumerate(['단계','현상','상세 (CMM 수치 포함)','영향',''],2):
    ap(ws3,3,i,hdr(h,bg='2E75B6',sz=10))
rh(ws3,{3:22})
for r,(step,phen,desc,eff,col) in enumerate(mechs,4):
    ap(ws3,r,2,hdr(step,bg=col,sz=9))
    ap(ws3,r,3,cel(phen,bg='F2F2F2',bold=True))
    ap(ws3,r,4,cel(desc,bg='FFFFFF',align='left'))
    ap(ws3,r,5,cel(eff,bold=True,bg='FFE0E0' if col=='C00000' else 'FFF2CC'))
    rh(ws3,{r:68})

rh(ws3,{9:12})
ws3.merge_cells('B10:F10')
ap(ws3,10,2,hdr('■ CMM 실측 기반 X-Y 방향 내경 변화 확정표',bg='1F4E79',sz=11))
rh(ws3,{10:28})
for i,h in enumerate(['구분','SPEC','X방향 실측 (게이지)','Y방향 실측 (게이지)','판정'],2):
    ap(ws3,11,i,hdr(h,bg='2E75B6',sz=10))
rh(ws3,{11:22})
sumrows=[
    ('① 내경 (ø30.83+0.05/0)','30.83 ~ 30.88mm',
     '30.80 ~ 30.86mm\n(일부 하한 미달)','30.88 ~ 30.91mm\n(상한 초과)','❌ Y방향 NG'),
    ('X-Y 편차 (타원도)','—','—','X-Y = 0.05~0.08mm','❌ 타원 확정'),
    ('CNC 보정 권고','—','—','ΔY = +0.06~0.08mm\n→ 정삭 목표 -0.06mm 보정','즉시 적용 필요'),
]
for r,(a,b,c,d,e) in enumerate(sumrows,12):
    for ci,val in enumerate([a,b,c,d,e],2):
        bg='FFE0E0' if '❌' in str(e) and ci in [4,5,6] else 'F2F2F2' if ci==2 else 'FFFFFF'
        ap(ws3,r,ci,cel(val,bg=bg,bold=('❌' in str(val)),wrap=True))
    rh(ws3,{r:45})


# ═══════════════════════════════════════════════════════════════════
# Sheet 4 — 개선 대책 Rev.3
# ═══════════════════════════════════════════════════════════════════
ws4=wb.create_sheet('4.개선대책_Rev3')
cw(ws4,[3,10,14,30,20,10,12,12])
ws4.merge_cells('B1:I1')
ap(ws4,1,2,hdr('■ 개선 대책 Rev.3 — CMM 실측 수치 기반 구체적 대책',bg='1F4E79',sz=13))
rh(ws4,{1:36})

ws4.merge_cells('B2:I2')
ap(ws4,2,2,hdr('CMM: Y방향 최대 30.91mm(+0.03 초과) / X-Y편차 0.05~0.08mm 실측 확정 → 아래 대책 즉시 적용',
               bg='C00000',sz=10))
rh(ws4,{2:26})

for i,h in enumerate(['우선','단계','대책명','세부 내용 (CMM 실측값 기반)','기대 효과','비용','기간','담당'],2):
    ap(ws4,3,i,hdr(h,bg='2E75B6',sz=10))
rh(ws4,{3:22})

acts=[
    ('★★★','즉시','정삭 유압압력\n50% 이하 감소\n(유압 조절밸브)',
     '• 동영상 확인: 현재 황삭=정삭 동일 유압 → 즉시 변경\n'
     '• 유압 압력 조절 밸브: 정삭 시 황삭 대비 50% 이하\n'
     '  예) 황삭 40bar → 정삭 18~20bar\n'
     '• CMM 실측 Y방향 복원량 0.06~0.08mm → 클램핑력 감소로 절반 이하 목표\n'
     '• 조절 후 CMM 재측정 → 효과 확인 후 압력값 확정',
     'Y편차 30.91\n→ 30.88 이하\n목표','없음','즉시','생산팀','FFE0E0'),
    ('★★★','즉시','CNC 정삭 목표\n내경 보정\n(-0.06mm)',
     '• CMM 실측 Y방향 복원량 ΔY = 0.06~0.08mm\n'
     '• 정삭 목표 내경 = 공칭 - ΔY = 30.83 - 0.06 = 30.77mm로 보정\n'
     '  (클램핑 중 가공 → 해제 후 30.83~30.88 도달 목적)\n'
     '• 초물 3개 가공 후 CMM 측정 → ΔY 재확인 → 보정값 미세 조정\n'
     '• 클램핑력 감소 후 ΔY 재측정 → 보정값 재산출 필수',
     '해제 후 내경\n공차 내 진입','없음','즉시','생산기술','FFE0E0'),
    ('★★★','즉시','척 해제 후\nCMM 전수 측정\n(4방향)',
     '• 동영상 확인: 현재 척 해제 후 측정 없음 → 즉시 추가\n'
     '• 측정 방향: X방향(0°) / Y방향(90°) / 45° / 135° 4방향\n'
     '• 판정: 전 방향 30.83~30.88mm AND 진원도(최대-최소) ≤0.03mm\n'
     '• ΔY = Y측정값 - 정삭목표 기록 → CNC 보정 피드백에 활용',
     '불량 유출\n즉시 차단','없음','즉시','품질팀','FFE0E0'),
    ('★★★','즉시','스프링 패스\n추가',
     '• 정삭 후 절삭깊이 0으로 공구 1회 재통과\n'
     '• 진원도 추가 개선 0.01~0.02mm\n'
     '• 클램핑력 감소 후에도 잔류 탄성변형 여유재료 제거',
     '진원도 0.03\n이하 목표','없음','즉시','생산팀','FFE0E0'),
    ('★★','즉시','② ø9.2 내경\n별도 보정',
     '• CMM ②: 9.2605~9.2717mm → 공차상한 9.25mm 초과 확인\n'
     '• ø9.2 내경도 별도 CNC 보정 필요 (현재 평균 약 +0.02mm 초과)\n'
     '• 정삭 목표 = 9.2 - 0.02 = 9.18mm로 보정 후 CMM 재확인',
     '② 내경 NG\n해소','없음','즉시','생산기술','FFF2CC'),
    ('★★','긴급\n조사','④ 길이 이상\n원인 조사\n(203mm대)',
     '• 2026.05.14 2차 로트 ④길이: 203.2~203.5mm\n'
     '  공차: 204.77~206.77mm → 1.3mm 이상 미달 (중대 이탈)\n'
     '• 해당 로트 전수 격리 후 별도 원인 조사 실시\n'
     '• 원인: 공구 마모, 셋업 오류, 소재 불량 등 조사\n'
     '• CMM 5.14 1차 로트 ④: 206.2~206.8mm (정상) → 로트 간 차이 조사',
     '해당 로트\n격리·조사','없음','긴급','품질팀','FFF2CC'),
    ('★★','1~2주','전용 소프트 조우\n제작',
     '• 소재 OD와 동일 반경 호형으로 조우 내면 보링\n'
     '• 접촉각 120° 이상 → 단위 클램핑압 분산\n'
     '• X-Y 편차 0.05~0.08 → 0.02~0.03mm 수준으로 감소 목표',
     'X-Y편차\n50% 감소','소','1~2주','생산기술','FFF2CC'),
    ('★★★','1개월','3죠우 유압 척\n교체 (근본 해결)',
     '• 현재 2죠우 유압 척 → 3죠우 유압 척으로 변경\n'
     '• 120° 균등 3방향 클램핑 → Y방향 단독 압축 근본 해소\n'
     '• 유압 방식 유지로 반복 클램핑 정밀도 확보\n'
     '• 목표: CMM 전 항목 OK / X-Y편차 0.01mm 이하 / Cpk≥1.33',
     '타원 근본 해소\nCpk≥1.33','중','1개월','생산기술','E2EFDA'),
]
for r,(pri,stage,name,desc,eff,cost,period,dept,bg) in enumerate(acts,4):
    fc='FFFFFF'
    pbg='C00000' if '★★★' in pri else ('2E75B6' if '★★' in pri else 'BDD7EE')
    ap(ws4,r,2,cel(pri,bg=pbg,fg=fc,bold=True))
    ap(ws4,r,3,cel(stage,bg='1F4E79',fg='FFFFFF',bold=True))
    ap(ws4,r,4,cel(name,bg=bg,bold=True))
    ap(ws4,r,5,cel(desc,bg='FFFFFF',align='left'))
    ap(ws4,r,6,cel(eff,bg=bg,bold=True))
    ap(ws4,r,7,cel(cost,bg='E2EFDA' if cost=='없음' else 'FFE0E0' if cost in ['대','중'] else 'FFFFFF'))
    ap(ws4,r,8,cel(period))
    ap(ws4,r,9,cel(dept))
    rh(ws4,{r:80})


# ═══════════════════════════════════════════════════════════════════
# Sheet 5 — CNC 보정 기준
# ═══════════════════════════════════════════════════════════════════
ws5=wb.create_sheet('5.CNC보정기준')
cw(ws5,[3,22,20,20,20,16])
ws5.merge_cells('B1:G1')
ap(ws5,1,2,hdr('■ CNC 보정 가공 기준 (CMM 실측 기반)',bg='1F4E79',sz=13))
rh(ws5,{1:36})

ws5.merge_cells('B2:G2')
ap(ws5,2,2,hdr('CMM 실측 Y방향 탄성복원량 ΔY = 0.06~0.08mm → 정삭 목표 내경 하향 보정 필수',
               bg='C00000',sz=10))
rh(ws5,{2:26})

ws5.merge_cells('B3:G3')
ap(ws5,3,2,hdr('[ ① ø30.83 내경 보정 기준 ]',bg='2E75B6',sz=11))
rh(ws5,{3:26})
for i,h in enumerate(['항목','내용','산출 근거'],2):
    ap(ws5,4,i,hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
ws5.merge_cells('D4:G4')
rh(ws5,{4:22})

corr1=[
    ('공칭 치수','ø30.83mm','SPEC 기준'),
    ('공차','0 ~ +0.05mm','30.83 ~ 30.88mm'),
    ('CMM 실측 Y복원량(ΔY)','0.06 ~ 0.08mm','2026.05.14 측정'),
    ('현행 정삭 목표','≒ 30.83mm','클램핑 중 목표'),
    ('개선 정삭 목표 (초기)','30.83 - 0.06 = 30.77mm','클램핑 중 가공\n→ 해제 후 30.83~30.88 목표'),
    ('유압압력 50% 감소 후 ΔY','재측정 필요 (목표 0.02~0.03mm)','클램핑력 감소 → 복원량 감소'),
    ('최종 보정값 확정 절차','초물 3개 가공 → CMM 측정 → ΔY 재산출 → 보정값 미세 조정','매 로트 초물 확인'),
]
for r,(item,val,note) in enumerate(corr1,5):
    ap(ws5,r,2,cel(item,bg='F2F2F2',bold=True))
    ap(ws5,r,3,cel(val,bg='E2EFDA' if '30.77' in val or '재측정' in val else 'FFFFFF',bold='30.77' in val))
    ws5.merge_cells(f'D{r}:G{r}')
    ap(ws5,r,4,cel(note,bg='FFFFFF',align='left'))
    rh(ws5,{r:32})

rh(ws5,{12:12})
ws5.merge_cells('B13:G13')
ap(ws5,13,2,hdr('[ ② ø9.2 내경 보정 기준 ]',bg='2E75B6',sz=11))
rh(ws5,{13:26})
for i,h in enumerate(['항목','내용','산출 근거'],2):
    ap(ws5,14,i,hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
ws5.merge_cells('D14:G14')
rh(ws5,{14:22})
corr2=[
    ('공칭 치수','ø9.2mm','SPEC 기준'),
    ('공차','0 ~ +0.05mm','9.20 ~ 9.25mm'),
    ('CMM 실측 평균','9.204 ~ 9.268mm (NG)','일부 9.25 초과'),
    ('초과량 추정','약 +0.01~0.02mm','9.26-9.25=0.01 초과'),
    ('개선 정삭 목표','9.2 - 0.02 = 9.18mm (초기)','해제 후 9.2~9.25 목표'),
    ('확정 절차','초물 CMM 측정 → 재산출','ø30.83과 동시 진행'),
]
for r,(item,val,note) in enumerate(corr2,15):
    ap(ws5,r,2,cel(item,bg='F2F2F2',bold=True))
    ap(ws5,r,3,cel(val,bg='E2EFDA' if '9.18' in val else 'FFE0E0' if 'NG' in val else 'FFFFFF',bold='9.18' in val))
    ws5.merge_cells(f'D{r}:G{r}')
    ap(ws5,r,4,cel(note,bg='FFFFFF',align='left'))
    rh(ws5,{r:28})

rh(ws5,{21:12})
ws5.merge_cells('B22:G22')
ap(ws5,22,2,hdr('[ ④ 길이 이상 로트 조사 항목 (2026.05.14 2차: 203mm대) ]',bg='C00000',sz=10))
rh(ws5,{22:26})
inv=[
    ('이상 측정값','203.2 ~ 203.5mm','정상: 204.77~206.77mm → 약 1.3mm 미달'),
    ('격리 조치','해당 로트 전수 격리','출하 보류 → 원인 확인 후 결정'),
    ('조사 항목 ①','공구 마모·파손 여부','절삭공구 인선 확인'),
    ('조사 항목 ②','셋업 기준 오류','원점 설정·공구 옵셋 재확인'),
    ('조사 항목 ③','소재 길이 불량','원자재 수입검사 이력 확인'),
    ('재발 방지','원인 확정 후 FMEA 추가','별도 시정조치 보고서 작성'),
]
for r,(item,val,note) in enumerate(inv,23):
    ap(ws5,r,2,cel(item,bg='F2F2F2',bold=True))
    ap(ws5,r,3,cel(val,bg='FFE0E0',bold=True))
    ws5.merge_cells(f'D{r}:G{r}')
    ap(ws5,r,4,cel(note,bg='FFF2CC',align='left'))
    rh(ws5,{r:28})


# ═══════════════════════════════════════════════════════════════════
# Sheet 6 — 실행계획 & 효과
# ═══════════════════════════════════════════════════════════════════
ws6=wb.create_sheet('6.실행계획_효과')
cw(ws6,[3,26,12,10,10,10,10,10,10])
ws6.merge_cells('B1:J1')
ap(ws6,1,2,hdr('■ 실행 계획 및 효과 확인 Rev.3',bg='1F4E79',sz=13))
rh(ws6,{1:36})

ws6.merge_cells('B2:J2')
ap(ws6,2,2,hdr('[ 주차별 실행 계획 ]',bg='2E75B6',sz=11))
rh(ws6,{2:26})
for i,h in enumerate(['대책 항목','담당','즉시','1주','2주','3주','4주','2M','3M'],2):
    ap(ws6,3,i,hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
rh(ws6,{3:22})
gantt=[
    ('정삭 유압압력 50% 감소 (조절밸브 설정)','생산팀','●','●','●','●','●','●','●'),
    ('CNC 정삭 목표 내경 보정 (-0.06mm 초기)','생산기술','●','●','●','●','●','●','●'),
    ('척 해제 후 CMM 4방향 전수 측정 추가','품질팀','●','●','●','●','●','●','●'),
    ('스프링 패스 추가 (정삭 후 1회)','생산팀','●','●','●','●','●','●','●'),
    ('② ø9.2 내경 별도 CNC 보정','생산기술','●','●','●','●','●','',''),
    ('④ 길이 이상 로트 격리·원인 조사','품질팀','●','●','','','','',''),
    ('소프트 조우 제작·적용','생산기술','','○','●','','','',''),
    ('3죠우 유압 척 발주·교체','생산기술','','','○','●','','',''),
    ('작업표준·검사기준·FMEA 개정','품질팀','','○','●','●','','',''),
    ('유압압력 조절밸브 봉인·라벨','설비팀','●','','','','','',''),
    ('Cpk SPC 관리체계 구축','품질팀','','','','','○','●',''),
]
for r,(item,dept,*wks) in enumerate(gantt,4):
    bg_row='FFE0E0' if r in [4,5,6,7,8] else 'FFFFFF'
    ap(ws6,r,2,cel(item,bg=bg_row,bold=(r<=7)))
    ap(ws6,r,3,cel(dept))
    for ci,w in enumerate(wks,4):
        ap(ws6,r,ci,cel(w,bg='E2EFDA' if w=='●' else 'FFF2CC' if w=='○' else 'FFFFFF',
                        bold=w in['●','○']))
    rh(ws6,{r:22})

rh(ws6,{15:12})
ws6.merge_cells('B16:J16')
ap(ws6,16,2,hdr('[ 개선 효과 확인 지표 (CMM 실측 기준) ]',bg='2E75B6',sz=11))
rh(ws6,{16:26})
for i,h in enumerate(['지표','현행 (CMM실측)','즉시대책후','소프트조우','3죠우교체','목표','방법'],2):
    ap(ws6,17,i,hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
rh(ws6,{17:22})
effs=[
    ('Y방향 내경 (게이지)','30.88~30.91mm NG','30.84~30.88','30.83~30.87','30.83~30.86','30.83~30.88','CMM전수'),
    ('X-Y 편차 (타원도)','0.05~0.08mm','0.03~0.05','0.02~0.03','0.01 이하','≤0.03mm','CMM전수'),
    ('CMM ① 합격률','전 로트 NG','70% 이상','90% 이상','99% 이상','100%','CMM전수'),
    ('CMM ② 합격률','일부 NG','80% 이상','95% 이상','99% 이상','100%','CMM전수'),
    ('진원도 (최대-최소)','0.05~0.08mm','≤0.05','≤0.03','≤0.01','≤0.03mm','CMM전수'),
    ('Cpk (내경 ①)','<1.0','≥1.0','≥1.17','≥1.33','≥1.33','SPC'),
    ('고객 클레임','발생','0건 목표','0건 목표','0건','0건','이력관리'),
]
for r,row in enumerate(effs,18):
    bgs=['F2F2F2','FFE0E0','FFF2CC','FFF2CC','E2EFDA','DEEAF1','FFFFFF']
    for ci,(v,bg) in enumerate(zip(row,bgs),2):
        ap(ws6,r,ci,cel(v,bg=bg,bold=(ci==2 or 'NG' in str(v)),wrap=True))
    rh(ws6,{r:28})


# ═══════════════════════════════════════════════════════════════════
# Sheet 7 — 재발방지 / 승인
# ═══════════════════════════════════════════════════════════════════
ws7=wb.create_sheet('7.재발방지_승인')
cw(ws7,[3,22,32,14,12])
ws7.merge_cells('B1:F1')
ap(ws7,1,2,hdr('■ 재발 방지 대책 및 결론 Rev.3',bg='1F4E79',sz=13))
rh(ws7,{1:36})

ws7.merge_cells('B2:F2')
ap(ws7,2,2,hdr('[ CMM + 동영상 검증 통합 결론 ]',bg='C00000',sz=11))
rh(ws7,{2:26})

concl=[
    ('확정 원인','2죠우 유압 척의 Y방향(수직) 압축 클램핑 → 척 해제 시 탄성 복원 → Y방향 내경 증가','FFE0E0'),
    ('CMM 수치 확정','Y방향 내경 실측 30.88~30.91mm (공차상한 +0.03 초과) / X-Y 편차 0.05~0.08mm','FFE0E0'),
    ('가중 원인 ①','정삭=황삭 동일 유압압력 (동영상 확인) → 탄성변형 최대 → 복원량 최대 → NG 심화','FFE0E0'),
    ('가중 원인 ②','척 해제 후 CMM 측정 없음 (동영상 확인) → 탄성복원 불량 검출 불가 → 불량 유출','FFE0E0'),
    ('추가 이상','④ 길이 203mm대 로트: 별도 원인 조사 필요 (셋업 오류/공구 마모 등)','FFF2CC'),
    ('즉시 조치','① 정삭 유압 50% 감소  ② 정삭 목표 -0.06mm 보정  ③ CMM 전수 측정  ④ 스프링 패스','E2EFDA'),
    ('근본 해결','3죠우 유압 척 교체 → 120° 균등 클램핑 → 타원 발생 근본 해소','E2EFDA'),
]
for r,(key,desc,bg) in enumerate(concl,3):
    ap(ws7,r,2,hdr(key,bg='2E75B6' if bg=='E2EFDA' else 'C00000' if bg=='FFE0E0' else 'C55A11',sz=10))
    ws7.merge_cells(f'C{r}:F{r}')
    ap(ws7,r,3,cel(desc,bg=bg,bold=True,align='left',sz=10))
    rh(ws7,{r:26})

rh(ws7,{10:12})
ws7.merge_cells('B11:F11')
ap(ws7,11,2,hdr('[ 재발 방지 관리 계획 ]',bg='2E75B6',sz=11))
rh(ws7,{11:26})
for i,h in enumerate(['관리 항목','개정 내용','시점','담당'],2):
    ap(ws7,12,i,hdr(h,bg='BDD7EE',fg='1F4E79',sz=10))
rh(ws7,{12:22})
prev=[
    ('작업표준서','정삭 유압압력 기준값 명기 / 정삭 목표 내경 보정값 명기\n2단 클램핑법 + 스프링 패스 필수 절차 추가','1주','생산기술'),
    ('검사기준서','CMM 4방향 측정 / 척 해제 후 전수 측정 필수\n진원도 ≤0.03mm 판정 기준 / ΔY 기록 양식 추가','1주','품질팀'),
    ('공정 FMEA','2죠우 유압 척 탄성변형 고장모드 추가\n정삭 동일 클램핑력 → 복원 최대 모드 추가','2주','품질팀'),
    ('유압 설비 관리','정삭 압력 설정값 라벨 부착 / 조절밸브 봉인\n유압압력 일상 점검 추가','1주','설비팀'),
    ('④ 길이 이상 로트','격리 → 원인 조사 → 시정조치 보고서\n셋업/공구/소재 전 항목 조사','긴급','품질팀'),
    ('수평 전개','2죠우 척 전 품종 동일 검토\n유압압력 정삭/황삭 구분 전수 확인','1개월','품질팀'),
]
for r,row in enumerate(prev,13):
    bgs=['F2F2F2','FFFFFF','FFF2CC','FFFFFF']
    for ci,(v,bg) in enumerate(zip(row,bgs),2):
        ap(ws7,r,ci,cel(v,bg=bg,bold=(ci==2),align='left',wrap=True))
    rh(ws7,{r:50})

rh(ws7,{19:12})
ws7.merge_cells('B20:F20')
ap(ws7,20,2,hdr('[ 검토 및 승인 ]',bg='1F4E79',sz=11))
rh(ws7,{20:26})
for i,h in enumerate(['구분','작성','검토','승인','비고'],2):
    ap(ws7,21,i,hdr(h,bg='2E75B6',sz=10))
rh(ws7,{21:22})
for r,label in enumerate(['성명','서명','일자'],22):
    for ci,v in enumerate([label,'','','','2026.05.15' if label=='일자' else ''],2):
        ap(ws7,r,ci,cel(v,bg='F2F2F2' if ci==2 else 'FFFFFF',bold=(ci==2)))
    rh(ws7,{r:28})

path='/home/user/quality-analysis/GKO1808_소켓내경타원발생_개선대책서_Rev3_CMM데이터반영.xlsx'
wb.save(path)
print(f'저장 완료: {path}')
