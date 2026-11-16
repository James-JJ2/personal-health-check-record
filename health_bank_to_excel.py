"""健康存摺 XML → Excel 統計程式
用法: python health_bank_to_excel.py 健康存摺.xml 輸出.xlsx
產生分頁: 說明、統計、門診就醫、門診醫令明細、牙科就醫、牙科醫令明細、疫苗接種、檢驗檢查結果、影像報告
"""
import sys, xml.etree.ElementTree as ET, re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
src, dst = sys.argv[1], sys.argv[2]
root = ET.parse(src).getroot()
b = root.find('bdata')
def g(e, t):
    x = e.find(t); return (x.text or '').strip() if x is not None and x.text else ''
def d(s):
    if re.fullmatch(r'\d{8}', s or ''):
        try: return datetime.strptime(s, '%Y%m%d').date()
        except: return s
    return s or None
def num(s):
    try: return float(s) if '.' in s else int(s)
    except: return s or None
wb = Workbook(); F='Arial'
hdrfill = PatternFill('solid', fgColor='1F4E78')
def sheet(name, headers, rows, widths=None):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for r in rows: ws.append(r)
    for c in ws[1]:
        c.font = Font(name=F, bold=True, color='FFFFFF'); c.fill = hdrfill
        c.alignment = Alignment(wrap_text=True, vertical='center')
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=F)
            if hasattr(c.value, 'year'): c.number_format = 'yyyy-mm-dd'
    ws.freeze_panes = 'A2'
    if rows: ws.auto_filter.ref = ws.dimensions
    for i, h in enumerate(headers, 1):
        w = (widths or {}).get(h, max(10, min(40, len(h)*2+4)))
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws
# 說明
ws = wb.active; ws.title = '說明'
info = [('資料來源', '健保署「健康存摺」XML 匯出'),
        ('身分證號（遮蔽）', g(b,'b1.1')),
        ('下載日期', d(g(b,'b1.2'))),
        ('聲明', re.sub(r'<[^>]+>', ' ', (g(b,'r0').replace('&lt;','<').replace('&gt;','>'))).strip()),
        ('住院資料 (r2)', g(b,'r2') or ''), ('r4', g(b,'r4')), ('r5', g(b,'r5')),
        ('備註', '欄位名稱依健康存摺格式推定；空白次診斷欄已合併。日期已轉為 Excel 日期格式。')]
for k,v in info: ws.append([k,v])
for row in ws.iter_rows():
    row[0].font = Font(name=F, bold=True); row[1].font = Font(name=F)
    row[1].alignment = Alignment(wrap_text=True, vertical='top')
    if hasattr(row[1].value,'year'): row[1].number_format='yyyy-mm-dd'
ws.column_dimensions['A'].width = 20; ws.column_dimensions['B'].width = 100
# 門診
vis, det = [], []
for i, r in enumerate(b.findall('r1'), 1):
    sec = []
    for k in range(14, 28, 2):
        c, n = g(r, f'r1.{k}'), g(r, f'r1.{k+1}')
        if c: sec.append(f'{c} {n}')
    vis.append([i, d(g(r,'r1.5')), d(g(r,'r1.6')), g(r,'r1.4'), g(r,'r1.3'), g(r,'r1.2'), g(r,'r1.7'),
                g(r,'r1.8'), g(r,'r1.9'), '；'.join(sec), num(g(r,'r1.12')), num(g(r,'r1.13'))])
    for x in r.findall('r1_1'):
        det.append([i, d(g(r,'r1.5') or g(r,'r1.6')), g(r,'r1.4'), g(x,'r1_1.1'), g(x,'r1_1.2'), num(g(x,'r1_1.3')), num(g(x,'r1_1.4'))])
sheet('門診就醫', ['編號','就醫日期','藥局調劑日期','醫事機構','機構代碼','分區業務組','就醫序號','主診斷代碼','主診斷','次診斷','部分負擔(元)','健保支付點數'],
      vis, {'次診斷':70,'主診斷':28,'醫事機構':14})
sheet('門診醫令明細', ['就醫編號','日期','醫事機構','醫令代碼','醫令/藥品名稱','總量','給藥日份'], det, {'醫令/藥品名稱':60})
# 牙科
dv, dd = [], []
for i, r in enumerate(b.findall('r3'), 1):
    sec = [f"{g(r,f'r3.{k}')} {g(r,f'r3.{k+1}')}" for k in range(13,23,2) if g(r,f'r3.{k}')]
    sp = [f"{g(r,f'r3.{k}')} {g(r,f'r3.{k+1}')}" for k in (23,25) if g(r,f'r3.{k}')]
    dv.append([i, d(g(r,'r3.5')), g(r,'r3.4'), g(r,'r3.3'), g(r,'r3.6'), g(r,'r3.7'), g(r,'r3.8'),
               '；'.join(sec), g(r,'r3.9'), g(r,'r3.10'), '；'.join(sp), num(g(r,'r3.11')), num(g(r,'r3.12'))])
    for x in r.findall('r3_1'):
        dd.append([i, d(g(r,'r3.5')), g(r,'r3.4'), g(x,'r3_1.1'), g(x,'r3_1.2'), num(g(x,'r3_1.3')), g(x,'r3_1.4'), g(x,'r3_1.5'), num(g(x,'r3_1.6'))])
sheet('牙科就醫', ['編號','就醫日期','醫事機構','機構代碼','就醫序號','主診斷代碼','主診斷','次診斷','主處置代碼','主處置','次處置','部分負擔(元)','健保支付點數'],
      dv, {'主處置':44,'次處置':50,'次診斷':36,'主診斷':24})
sheet('牙科醫令明細', ['就醫編號','日期','醫事機構','醫令代碼','醫令名稱','數量','牙位代碼','牙位','給藥日份'], dd, {'醫令名稱':36})
# 疫苗
sheet('疫苗接種', ['接種日期','疫苗名稱','接種地點'], [[d(g(r,'r6.1')), g(r,'r6.3'), g(r,'r6.5')] for r in b.findall('r6')], {'疫苗名稱':45,'接種地點':20})
# 檢驗
lab = []
for r in b.findall('r7'):
    rng = g(r,'r7.12'); m = re.fullmatch(r'\[(.*?)\]\[(.*?)\]', rng)
    lo, hi = (m.group(1), m.group(2)) if m else (rng, '')
    res = g(r,'r7.11')
    lab.append([d(g(r,'r7.5')), d(g(r,'r7.6')), g(r,'r7.7'), g(r,'r7.4'), g(r,'r7.8'), g(r,'r7.9'), g(r,'r7.10'), num(res), lo, hi])
lab.sort(key=lambda x: (str(x[1]), x[2]))
sheet('檢驗檢查結果', ['就醫日期','採檢/報告日期','上傳時間','醫事機構','醫令代碼','醫令名稱','檢驗項目','結果','參考值下限/說明','參考值上限'], lab,
      {'醫令名稱':30,'檢驗項目':20,'上傳時間':16})
# 影像
ws = sheet('影像報告', ['就醫日期','報告日期','醫事機構','醫令代碼','檢查名稱','報告內容'],
      [[d(g(r,'r8.5')), d(g(r,'r8.6')), g(r,'r8.4'), g(r,'r8.8'), g(r,'r8.9'), re.sub(r'\s{2,}', '\n', g(r,'r8.10'))] for r in b.findall('r8')],
      {'檢查名稱':40,'報告內容':80})
for c in ws['F'][1:]: c.alignment = Alignment(wrap_text=True, vertical='top')

# ===== 統計 =====
from collections import Counter
nv, nd, nl = len(vis)+1, len(dv)+1, len(lab)+1
# 年度輔助欄
for wsn, col, n, src_expr in [('門診就醫','M',nv,'IF(B{r}="",C{r},B{r})'), ('牙科就醫','N',nd,'B{r}')]:
    w = wb[wsn]; w[f'{col}1'] = '年度'
    w[f'{col}1'].font = Font(name=F, bold=True, color='FFFFFF'); w[f'{col}1'].fill = hdrfill
    for r in range(2, n+1):
        w[f'{col}{r}'] = f'=YEAR({src_expr.format(r=r)})'; w[f'{col}{r}'].font = Font(name=F)
st = wb.create_sheet('統計', 1)
H = lambda c: (setattr(c,'font',Font(name=F,bold=True,color='FFFFFF')), setattr(c,'fill',hdrfill))
T = lambda c: setattr(c,'font',Font(name=F,bold=True,size=13,color='1F4E78'))
row = 1
st.cell(row,1,'健康存摺統計（公式自動計算，修改資料分頁後會更新）'); T(st.cell(row,1)); row += 2
# 1. 年度
st.cell(row,1,'一、年度就醫統計'); T(st.cell(row,1)); row += 1
hd = ['年度','門診紀錄筆數','門診部分負擔(元)','門診健保點數','牙科紀錄筆數','牙科部分負擔(元)','牙科健保點數','合計部分負擔(元)','合計健保點數']
for j,h in enumerate(hd,1): H(st.cell(row,j,h))
row += 1; y0 = row
years = sorted({v[1].year if hasattr(v[1],'year') else v[2].year for v in vis} | {v[1].year for v in dv if hasattr(v[1],'year')})
M, D = "'門診就醫'", "'牙科就醫'"
for y in years:
    r = row
    vals = [y, f'=COUNTIF({M}!$M$2:$M${nv},A{r})', f'=SUMIF({M}!$M$2:$M${nv},A{r},{M}!$K$2:$K${nv})',
            f'=SUMIF({M}!$M$2:$M${nv},A{r},{M}!$L$2:$L${nv})', f'=COUNTIF({D}!$N$2:$N${nd},A{r})',
            f'=SUMIF({D}!$N$2:$N${nd},A{r},{D}!$L$2:$L${nd})', f'=SUMIF({D}!$N$2:$N${nd},A{r},{D}!$M$2:$M${nd})',
            f'=C{r}+F{r}', f'=D{r}+G{r}']
    for j,v in enumerate(vals,1): st.cell(r,j,v)
    row += 1
st.cell(row,1,'合計').font = Font(name=F,bold=True)
for j in range(2,10):
    L = get_column_letter(j); st.cell(row,j,f'=SUM({L}{y0}:{L}{row-1})').font = Font(name=F,bold=True)
row += 1
st.cell(row,1,'註：同一次就醫可能拆成多筆（看診 / 檢驗 / 藥局調劑），故以「紀錄筆數」計。年度依就醫日期，無則用調劑日期。').font = Font(name=F,italic=True,color='666666')
row += 2
# 2. 醫事機構
st.cell(row,1,'二、各醫事機構統計'); T(st.cell(row,1)); row += 1
for j,h in enumerate(['類別','醫事機構','紀錄筆數','部分負擔(元)','健保點數','首次就醫','最近就醫'],1): H(st.cell(row,j,h))
row += 1
for kind, rows_, sh, n, pc, ptc in [('門診',vis,M,nv,'K','L'),('牙科',dv,D,nd,'L','M')]:
    col = 'D' if kind=='門診' else 'C'
    for inst,_ in Counter(v[3] if kind=='門診' else v[2] for v in rows_).most_common():
        r = row
        st.cell(r,1,kind); st.cell(r,2,inst)
        st.cell(r,3,f'=COUNTIF({sh}!${col}$2:${col}${n},B{r})')
        st.cell(r,4,f'=SUMIF({sh}!${col}$2:${col}${n},B{r},{sh}!${pc}$2:${pc}${n})')
        st.cell(r,5,f'=SUMIF({sh}!${col}$2:${col}${n},B{r},{sh}!${ptc}$2:${ptc}${n})')
        st.cell(r,6,f'=IF(_xlfn.MAXIFS({sh}!$B$2:$B${n},{sh}!${col}$2:${col}${n},B{r})=0,"-",_xlfn.MINIFS({sh}!$B$2:$B${n},{sh}!${col}$2:${col}${n},B{r},{sh}!$B$2:$B${n},">0"))')
        st.cell(r,7,f'=IF(_xlfn.MAXIFS({sh}!$B$2:$B${n},{sh}!${col}$2:${col}${n},B{r})=0,"-",_xlfn.MAXIFS({sh}!$B$2:$B${n},{sh}!${col}$2:${col}${n},B{r}))')
        for c in (6,7): st.cell(r,c).number_format = 'yyyy-mm-dd'
        row += 1
row += 1
# 3. 主診斷
st.cell(row,1,'三、門診主診斷次數'); T(st.cell(row,1)); row += 1
for j,h in enumerate(['診斷代碼','診斷名稱','紀錄筆數','占比'],1): H(st.cell(row,j,h))
row += 1; dx0 = row
for (code,name),_ in Counter((v[7],v[8]) for v in vis).most_common():
    st.cell(row,1,code); st.cell(row,2,name)
    st.cell(row,3,f'=COUNTIF({M}!$H$2:$H${nv},A{row})'); row += 1
for r in range(dx0,row):
    st.cell(r,4,f'=C{r}/SUM($C${dx0}:$C${row-1})'); st.cell(r,4).number_format='0.0%'
row += 1
# 4. 常用藥品 / 醫令
st.cell(row,1,'四、最常開立醫令 / 藥品（前 20）'); T(st.cell(row,1)); row += 1
for j,h in enumerate(['醫令代碼','名稱','開立次數','總量合計'],1): H(st.cell(row,j,h))
row += 1; ND = len(det)+1; MD="'門診醫令明細'"
names = {x[3]:x[4] for x in det}
for code,_ in Counter(x[3] for x in det).most_common(20):
    st.cell(row,1,code); st.cell(row,2,names[code])
    st.cell(row,3,f'=COUNTIF({MD}!$D$2:$D${ND},A{row})')
    st.cell(row,4,f'=SUMIF({MD}!$D$2:$D${ND},A{row},{MD}!$F$2:$F${ND})'); row += 1
row += 1
# 5. 檢驗
st.cell(row,1,'五、重要檢驗項目追蹤（依醫令代碼）'); T(st.cell(row,1)); row += 1
for j,h in enumerate(['醫令代碼','項目','檢驗次數','平均值','最低值','最高值','最近一次結果','最近採檢日'],1): H(st.cell(row,j,h))
row += 1; NL = len(lab)+1; LB="'檢驗檢查結果'"
items = [('09006C','醣化血紅素 HbA1c'),('09005C','空腹血糖'),('09001C','總膽固醇'),('09004C','三酸甘油脂'),
         ('09043C','高密度膽固醇 HDL'),('09044C','低密度膽固醇 LDL'),('09106C','游離甲狀腺素 Free T4'),
         ('09112C','甲狀腺刺激素 TSH'),('09117C','T3'),('09013C','尿酸'),('09015C','肌酸酐'),
         ('09026C','GPT (ALT)'),('09025C','GOT (AST)'),('09022C','鉀'),('Y00001','eGFR'),('Y00005','體重')]
rng = lambda L: f'{LB}!${L}$2:${L}${NL}'
for code,name in items:
    r=row; st.cell(r,1,code); st.cell(r,2,name)
    st.cell(r,3,f'=COUNTIF({rng("E")},A{r})')
    st.cell(r,4,f'=IFERROR(AVERAGEIF({rng("E")},A{r},{rng("H")}),"-")'); st.cell(r,4).number_format='0.00'
    st.cell(r,5,f'=_xlfn.MINIFS({rng("H")},{rng("E")},A{r})')
    st.cell(r,6,f'=_xlfn.MAXIFS({rng("H")},{rng("E")},A{r})')
    st.cell(r,7,f'=LOOKUP(2,1/({rng("E")}=A{r}),{rng("H")})')
    st.cell(r,8,f'=LOOKUP(2,1/({rng("E")}=A{r}),{rng("B")})'); st.cell(r,8).number_format='yyyy-mm-dd'
    row += 1
st.cell(row,1,'註：平均/最低/最高僅計數值結果；「＜0.005」等文字結果不納入，但會顯示在「最近一次結果」。').font = Font(name=F,italic=True,color='666666')
for rr in st.iter_rows():
    for c in rr:
        if c.font is None or c.font.name != F: c.font = Font(name=F, bold=c.font.bold if c.font else False)
for L,w in zip('ABCDEFGHI',[12,34,14,16,14,16,16,16,14]): st.column_dimensions[L].width = w
st.freeze_panes = 'A3'
wb.save(dst)
print(len(vis), len(det), len(dv), len(dd), len(lab))
