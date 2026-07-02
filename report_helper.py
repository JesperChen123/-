import os, sys, traceback, statistics
from datetime import datetime
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))

#小数点处理部分以及N/A处理
#四舍五入到两位
def fmt2(v):
    if v is None: return ''
    return f'{v:.2f}' if isinstance(v, (int, float)) else str(v)
#取整，比如行业排名
def fmt0(v):
    if v is None: return ''
    return str(round(v)) if isinstance(v, (int, float)) else str(v)
#百分比
def fmt_pct(v):
    if v is None: return ''
    return f'{v*100:.2f}%' if isinstance(v, (int, float)) else str(v)
#字符串，比如股票编号
def fmt_str(v):
    return '' if v is None else str(v)
#日期
def fmt_date(v):
    return v.strftime('%Y-%m-%d') if isinstance(v, datetime) else (str(v) if v else '')

def xl(ws, r, c):
    try: return ws.cell(r, c).value
    except: return None

# Word tools
#单元格里写内容
def set_cell(cell, text, font_size=9, bold=False, align='center'):
    if cell is None: return
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == 'center' else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(str(text) if text is not None else '')
    run.font.name = '宋体'; run.font.size = Pt(font_size); run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

#给表格加边框
def add_borders(table):
    xml = (f'<w:tblBorders {nsdecls("w")}>'
        + ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  for s in ['top','left','bottom','right','insideH','insideV'])
        + '</w:tblBorders>')
    pr = table._tbl.tblPr
    if pr is None: pr = parse_xml(f'<w:tblPr {nsdecls("w")}/>'); table._tbl.insert(0, pr)
    for old in pr.findall(qn('w:tblBorders')): pr.remove(old)
    pr.append(parse_xml(xml))

#指定位置创建表格（先创再移）
def make_table(doc, rows, cols, after_el, col_widths=None, compact=False):
    tbl = doc.add_table(rows=rows, cols=cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
    add_borders(tbl)
    if col_widths:
        for row in tbl.rows:
            for i, w in enumerate(col_widths):
                if i < len(row.cells): row.cells[i].width = Cm(w)
    if compact:
        m = parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:top w:w="0" w:type="dxa"/>'
            '<w:left w:w="30" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>'
            '<w:right w:w="30" w:type="dxa"/></w:tblCellMar>')
        for old in tbl._tbl.tblPr.findall(qn('w:tblCellMar')): tbl._tbl.tblPr.remove(old)
        tbl._tbl.tblPr.append(m)
    doc.element.body.remove(tbl._element)
    after_el.addnext(tbl._element)
    return tbl

#特定位置插段落
def make_para(doc, text, after_el, font_size=9, bold=False):
    p = doc.add_paragraph(); run = p.add_run(text)
    run.font.name = '宋体'; run.font.size = Pt(font_size); run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    doc.element.body.remove(p._element); after_el.addnext(p._element)
    return p

def find_para(doc, kw):
    for p in doc.paragraphs:
        if kw in p.text: return p._element
    return None

#清空内容，有空间才能插入表格
def clear_between(doc, start_kw, end_kw):
    body = doc.element.body
    s = find_para(doc, start_kw); e = find_para(doc, end_kw) if end_kw else None
    if s is None: return None
    rm, inside = [], False
    for ch in body.iterchildren():
        if ch is s: inside = True; continue
        if e is not None and ch is e: break
        if inside: rm.append(ch)
    for el in rm: body.remove(el)
    return s

#找到文件夹里第一个word当作模板，和第一个excel当作数据库
def get_files():
    ex = wd = None
    for f in sorted(os.listdir(BASE_DIR)):
        if f.startswith('~$') or '合并完成' in f: continue
        if f.endswith('.xlsx') and not ex: ex = os.path.join(BASE_DIR, f)
        elif f.endswith('.docx') and not wd: wd = os.path.join(BASE_DIR, f)
    return ex, wd

#  公司对比的基础表格
def fill_comparison(doc, wb):
    if '对标' not in wb.sheetnames: return print('没有"对标"sheet')
    ws = wb['对标']
    anchor = clear_between(doc, '1、对标同行业上市公司的关键指标比较', '2、标的证券核心竞争力')
    if anchor is None: return print('未找到对标段落')

    last = anchor
    W6, W10, W4 = [1.5,2.0,2.5,2.5,2.5,2.5], [1.2,1.3,1.5,1.5,1.5,1.5,1.5,1.5,1.5,1.5], [1.5,1.5,5.0,5.0]

    def add_tbl(title, headers, data, fsz=9, w=None, compact=False):
        nonlocal last
        p = make_para(doc, title, last, font_size=10, bold=True); last = p._element
        nc, nr = len(headers[-1]), len(headers)+len(data)
        t = make_table(doc, nr, nc, last, col_widths=w, compact=compact)
        for hi, hr in enumerate(headers):
            for ci, h in enumerate(hr): set_cell(t.cell(hi,ci), h, bold=True, font_size=fsz)
        for ri, dr in enumerate(data):
            for ci, v in enumerate(dr):
                set_cell(t.cell(len(headers)+ri,ci), v, font_size=fsz, align='left' if ci<=1 else 'center')
        last = t._element

#生成简单的表格
    def simple6(title, rows, ylabels):
        add_tbl(title, [['代码','证券简称']+ylabels],
            [[fmt_str(xl(ws,r,3)),fmt_str(xl(ws,r,4))]+[fmt2(xl(ws,r,c)) for c in range(5,9)] for r in rows], w=W6)
        print(f'{title}')

#稍微复杂点的表格（两个数据区）
    def dual10(title, hdr, comps):
        yv = [fmt_str(xl(ws,hdr,c)) for c in [5,6,7,8]]
        yg = [fmt_str(xl(ws,hdr,c)) for c in [9,10,11,12]]
        add_tbl(title,
            [['','',f'{title}（亿元）','','','','同比增长率（%）','','',''], ['代码','证券简称']+yv+yg],
            [[fmt_str(xl(ws,r,3)),fmt_str(xl(ws,r,4))]+[fmt2(xl(ws,r,c)) for c in range(5,9)]+[fmt2(xl(ws,r,c)) for c in range(9,13)] for r in comps],
            fsz=8, w=W10, compact=True)
        print(f'{title}')

    #调用顺序和locate位置
    add_tbl('主营构成', [['代码','证券简称','主营构成(产品）','主要产品名称']],
        [[fmt_str(xl(ws,r,3)),fmt_str(xl(ws,r,4)),fmt_str(xl(ws,r,5)),fmt_str(xl(ws,r,9))] for r in [25,26,27]], fsz=8, w=W4)
    print('主营构成')

    simple6('净资产收益率（ROE）(%) -加权', [35,36,37], ['2023年','2024年','2025年','2026年Q1(年化)'])
    simple6('销售净利率(%)',                [41,42,43], ['2023年','2024年','2025年','2026年Q1'])
    simple6('总资产周转率（次/年）',          [47,48,49], ['2023年','2024年','2025年','2026年Q1(年化)'])

    for t,h,c in [('营业总收入',61,[62,63,64]),('净利润',70,[71,72,73]),('归母净利润',81,[82,83,84]),('扣非后归母净利润',90,[91,92,93])]:
        dual10(t,h,c)

    add_tbl('毛利率',
        [['','','销售毛利率(%)','','','','扣除销售费用后的毛利率(%)','','',''],
         ['代码','证券简称','2023年','2024年','2025年','2026年Q1','2023年','2024年','2025年','2026年Q1']],
        [[fmt_str(xl(ws,r,3)),fmt_str(xl(ws,r,4))]+[fmt2(xl(ws,r,c)) for c in range(5,13)] for r in [103,104,105]],
        fsz=8, w=W10, compact=True)
    print('毛利率')


#找关键词删表格，但是要防止误删段落内容
def fill_industry(doc, wb):
    if '标的券概要' not in wb.sheetnames: return print('没有"标的券概要"sheet')
    ws = wb['标的券概要']

    anchor = None
    for p in doc.paragraphs:
        if '与行业上市公司比较' in p.text or '与同行业上市公司比较' in p.text:
            anchor = p._element; break
    if anchor is None: anchor = find_para(doc, '标的券估值风险分析')
    if anchor is None: return print('未找到行业比较段落')

    el = anchor.getnext()
    while el is not None:
        if el.tag.split('}')[-1] == 'tbl': doc.element.body.remove(el); break
        txt = ''.join(n.text or '' for n in el.iter() if n.tag.endswith('}t')).strip()
        if txt and not any(k in txt for k in ['标的券','下表','与']): break
        el = el.getnext()

    #行业比较的数据处理（11个公司）
    comps = []
    for r in range(61, 72):
        if xl(ws,r,3): comps.append([xl(ws,r,c) for c in [2,3,4,6,7,8,9,10]])

    #中位值和平均值：PE_TTM/PB/PB扣商誉直接从Excel读（45-47），PE 26E/27E自己算
    tm  = xl(ws,45,5) or 0;  ta  = xl(ws,45,4) or 0   #PE TTM
    pm  = xl(ws,46,5) or 0;  pa  = xl(ws,46,4) or 0   #PB
    xm  = xl(ws,47,5) or 0;  xa  = xl(ws,47,4) or 0   #PB扣商誉

    valid = [c for c in comps if isinstance(c[3],(int,float)) and 0 < c[3] <= 100]
    def stats(vals):
        f = [x for x in vals if isinstance(x,(int,float))]
        return (statistics.median(f),statistics.mean(f)) if f else (0,0)
    e6m,e6a = stats([c[4] for c in valid if c[4] and c[4]>0])
    e7m,e7a = stats([c[5] for c in valid if c[5] and c[5]>0])

    #格式
    tbl = make_table(doc, 2+len(comps)+2, 8, anchor)
    for ci,h in enumerate(['排名','代码','简称','市盈率PE','市盈率PE','市盈率PE','市净率PB（LF）','市净率（扣除商誉）']):
        set_cell(tbl.cell(0,ci),h,bold=True)
    for ci,h in enumerate(['排名','代码','简称','TTM','26E','27E','市净率PB（LF）','市净率（扣除商誉）']):
        set_cell(tbl.cell(1,ci),h,bold=True)
    for ri,c in enumerate(comps):
        for ci,v in enumerate([fmt0(c[0]),fmt_str(c[1]),fmt_str(c[2]),fmt2(c[3]),fmt2(c[4]),fmt2(c[5]),fmt2(c[6]),fmt2(c[7])]):
            set_cell(tbl.cell(ri+2,ci),v,align='left' if ci in [1,2] else 'center')
    for label,row_i,vals in [('行业中值',2+len(comps),[tm,e6m,e7m,pm,xm]),('行业均值',3+len(comps),[ta,e6a,e7a,pa,xa])]:
        set_cell(tbl.cell(row_i,0),label,bold=True)
        for ci,v in enumerate(vals): set_cell(tbl.cell(row_i,3+ci),fmt2(v))
    print('行业估值比较 (标的券概要)')

def fill_financial(doc, wb):
    if '标的券概要' not in wb.sheetnames: return print('没有"标的券概要"sheet')
    ws = wb['标的券概要']

    # 跳过目录，定位到最后一个匹配的"三...财务分析"
    anchor = None
    for p in doc.paragraphs:
        if 'toc' in (p.style.name or '').lower(): continue
        if '财务分析' in p.text and p.text.strip().startswith('三'):
            anchor = p._element
    if anchor is None: return print('未找到"三、财务分析"段落')

    el = anchor.getnext()
    while el is not None:
        if el.tag.split('}')[-1] == 'tbl': doc.element.body.remove(el); break
        el = el.getnext()

    # Excel行, 标签, 格式化
    R = [(81,'营业收入(亿元)',fmt2),(82,'同比增长率(%)',fmt_pct),(83,'归母净利润(亿元)',fmt2),
         (84,'同比增长率(%)',fmt_pct),(85,'扣非归母净利润(亿元)',fmt2),(86,'同比增长率(%)',fmt_pct),
         (87,'毛利率(%)',fmt_pct),(88,'净资产收益率ROE(加权,%)',fmt_pct),
         (89,'商誉(亿元)',fmt2),(90,'商誉占总资产比例(%)',fmt_pct),
         (91,'总资产(亿元)',fmt2),(92,'归母净资产(亿元)',fmt2),
         (93,'资产负债率(%)',fmt_pct),(94,'速动比率',fmt2),
         (95,'经营活动现金流量净额(亿元)',fmt2),
         (96,'应收账周转率（年化）',fmt2),(97,'存货周转率（年化）',fmt2),
         (99,'净利润(亿元)',fmt2),(100,'同比增长率(%)',fmt_pct),(101,'销售净利润率(%)',fmt_pct)]

    hdrs = [fmt_date(xl(ws,80,c)) if isinstance(xl(ws,80,c),datetime) else str(xl(ws,80,c)) for c in [4,5,6,7]]
    tbl = make_table(doc, 1+len(R), 5, anchor) #建表，填表标题
    set_cell(tbl.cell(0,0),'主要财务科目及比率',bold=True,align='left')
    for ci,h in enumerate(hdrs): set_cell(tbl.cell(0,ci+1),h,bold=True)
    for ri,(xr,label,f) in enumerate(R): #然后填数据
        set_cell(tbl.cell(ri+1,0),label,align='left')
        for ci,xc in enumerate([4,5,6,7]): set_cell(tbl.cell(ri+1,ci+1),f(xl(ws,xr,xc)))
    print('财务分析 (标的券概要)')


#往word模板自带的表里填数据（看表标题关键词判断）
def fill_legacy(doc, wb_val, wb_fmt):
    def find_tbl(kw):
        for t in doc.tables:
            for row in t.rows[:2]:
                for cell in row.cells:
                    if kw in cell.text: return t
        return None
    #读值，再根据根式改成相对应的数字（一次拿数字，一次拿格式）
    def sv(sh, addr):
        try:
            v = wb_val[sh][addr].value; nf = wb_fmt[sh][addr].number_format
            if isinstance(v,datetime): return v.strftime('%Y-%m-%d')
            if isinstance(v,(int,float)): return f'{v:.2%}' if '%' in nf else (f'{v:.2f}' if isinstance(v,float) else str(v))
            return str(v) if v else ''
        except: return None
    #防止崩溃
    def sc(t,r,c):
        try: return t.cell(r,c) if 0<=r<len(t.rows) and 0<=c<len(t.columns) else None
        except: return None

    if '存续' in wb_val.sheetnames:
        t = find_tbl('序号')
        if t:
            while len(t.rows)<9: t.add_row() #防止行数不够
            for cl,wc in {'C':1,'E':3,'F':4,'G':5,'H':6}.items():
                for i in range(1,11): set_cell(sc(t,i,wc), sv('存续',f'{cl}{i+9}'))
            print('存续交易')
    #还没实装下面的东西，但是逻辑有一点了
    if '本次要素' in wb_val.sheetnames:
        t = find_tbl('持股')
        if t:
            d = {}
            for r in range(63,79):
                k = wb_val['本次要素'][f'B{r}'].value
                if k:
                    key = str(k).strip().replace(' ','')
                    d[key] = {'C':sv('本次要素',f'C{r}'), 'E':sv('本次要素',f'E{r}')}
            for row in t.rows:
                lbl = row.cells[0].text.strip().replace(' ','')
                if lbl in d:
                    if len(row.cells)>1: set_cell(row.cells[1], d[lbl].get('C'))
                    if len(row.cells)>4: set_cell(row.cells[4], d[lbl].get('E'))
            print('持股概要')


#  主函数，标题，打开文件拿数据格式，按顺序加载每个模块，保存，以及报错处理
def main():
    print('='*50+'\n  尽调报告自动填表工具\n'+'='*50)
    print(f'工作目录: {BASE_DIR}\n')
    excel_path, word_path = get_files()
    if not excel_path: print('未找到 .xlsx'); input('回车退出...'); return
    if not word_path:   print('未找到 .docx'); input('回车退出...'); return
    print(f'Excel: {os.path.basename(excel_path)}\nWord:  {os.path.basename(word_path)}\n')
    try:
        wb_val = load_workbook(excel_path, data_only=True)
        wb_fmt = load_workbook(excel_path, data_only=False)
        doc = Document(word_path)
        print('【对标模块】');   fill_comparison(doc, wb_val)
        print('\n【行业比较】'); fill_industry(doc, wb_val)
        print('\n【财务分析】'); fill_financial(doc, wb_val)
        print('\n【基础模块】'); fill_legacy(doc, wb_val, wb_fmt)
        ts = datetime.now().strftime('%Y%m%d_%H%M')
        out = os.path.join(BASE_DIR, f'合并_尽调报告{ts}.docx')
        doc.save(out)
        print(f'\n保存为: {os.path.basename(out)}')
    except Exception as e:
        print(f'\n出错: {e}'); print(traceback.format_exc())
    input('\n回车关闭...')

if __name__ == '__main__':
    main()