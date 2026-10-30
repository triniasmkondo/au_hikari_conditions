"""Phase 1: source-linked business rows; no calculation or approval inference."""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
import hashlib
import json
from pathlib import Path
import re
import unicodedata

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph



def normal(value) -> str:
    return ' '.join(unicodedata.normalize('NFKC', str(value or '')).split())


def agency_candidate(path: Path) -> str:
    labels = re.findall(r'【([^】]+)】', path.stem)
    return re.sub(r'\s*(御中|様)$', '', labels[1]).strip() if len(labels) >= 2 else '代理店不明'


def candidate_key(item: str, dimensions: dict, terms: str) -> str:
    payload = json.dumps([normal(item), dimensions, normal(terms)], sort_keys=True, ensure_ascii=False)
    return 'C-' + hashlib.sha256(payload.encode()).hexdigest()[:24]


@dataclass
class Condition:
    agency: str
    kind: str
    product: str
    section: str
    item: str
    isp: str = ''
    type: str = ''
    target: str = ''
    code: str = ''
    threshold: str = ''
    amount: int | float | None = None
    unit: str = ''
    tax: str = ''
    period: str = ''
    terms: str = ''
    legal: str = ''
    source: str = ''
    locations: list[str] = field(default_factory=list)
    digest: str = ''
    hold: str = ''

    @property
    def id(self):
        return candidate_key(self.item, dict(agency=normal(self.agency), kind=self.kind,
            product=self.product, section=normal(self.section), isp=self.isp, type=self.type,
            target=normal(self.target), code=normal(self.code), threshold=self.threshold,
            unit=self.unit, tax=self.tax, period=self.period), self.terms)


HEADERS = ['代理店', '条件ID', '区分', '商材', '表・施策', '成果項目', 'ISP', '対象タイプ',
           '対象・住居・速度', '特典コード', '閾値・集計条件', '金額', '単位', '税区分',
           '適用期間', '適用条件・備考', '支払・除外・戻入等の条項（根拠位置付き）',
           '根拠資料', '根拠位置', 'SHA256', '取込状態', '保留理由']


def row_values(c):
    return [c.agency, c.id, c.kind, c.product, c.section, c.item, c.isp, c.type, c.target,
            c.code, c.threshold, c.amount, c.unit, c.tax, c.period, c.terms, c.legal,
            c.source, '\n'.join(c.locations), c.digest, '要確認' if c.hold else '原本対応済・未承認', c.hold]


def merged_value(sheet, row, column):
    cell = sheet.cell(row, column)
    for area in sheet.merged_cells.ranges:
        if cell.coordinate in area:
            anchor = sheet.cell(area.min_row, area.min_col)
            return normal(anchor.value), anchor.coordinate
    return normal(cell.value), cell.coordinate


def text_rows(sheet):
    return {r: ' | '.join(f'{c.coordinate}={c.value}' for c in sheet[r] if c.value is not None)
            for r in range(1, sheet.max_row + 1) if any(c.value is not None for c in sheet[r])}


def period_from(text):
    match = re.search(r'(20\d{2})\s*年\s*(\d+)\s*月\s*(\d+)\s*日\s*[~〜～]\s*(20\d{2})\s*年\s*(\d+)\s*月\s*(\d+)\s*日', normal(text))
    if not match:
        return ''
    y,m,d,ey,em,ed = map(int, match.groups())
    return f'{y:04}-{m:02}-{d:02}/{ey:04}-{em:02}-{ed:02}'


def extract_excel(path, root):
    book = load_workbook(path, data_only=False)
    conditions, coverage, legal_rows, raw_rows = [], [], [], []
    agency, source = agency_candidate(path), path.relative_to(root).as_posix()
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    try:
        for sheet in book:
            texts = text_rows(sheet)
            if not texts:
                continue
            plus = 'auひかりプラス' in sheet.title
            product = 'auひかりプラス' if plus else 'auひかり'
            period = next((period_from(t) for t in texts.values() if period_from(t)), '')
            # Every text row is retained in the linked source clauses, including split sentences.
            legal = '\n'.join(f'sheet:{sheet.title};row:{r} {t}' for r,t in texts.items()
                              if not any(isinstance(c.value,(int,float)) or c.data_type=='f' for c in sheet[r]))
            legal_rows.append([agency,source,sheet.title,legal])
            raw_rows.extend([agency,source,f'sheet:{sheet.title};row:{r}',t] for r,t in texts.items())
            sections = []
            for r,t in texts.items():
                values = [normal(c.value) for c in sheet[r] if c.value is not None]
                if ('支払レート' in values or ('項目' in values and ('5G開通手数料' in values or 'タイプ1' in values))):
                    title = ' / '.join(texts[j] for j in range(max(1,r-3),r) if j in texts and
                                       ('■' in texts[j] or re.search(r'表\s*[0-9１-９]|その他諸経費',texts[j])))
                    if not title:
                        title = 'その他諸経費' if '対象品目' in values else ''
                    # Remove cell coordinates from semantic section names.
                    title = re.sub(r'\b[A-Z]+\d+=', '', title)
                    sections.append((r,title,values))
            for ri,row in enumerate(sheet,1):
                for cell in row:
                    if not (isinstance(cell.value,(int,float)) and not isinstance(cell.value,bool) or cell.data_type=='f'):
                        continue
                    loc = f'sheet:{sheet.title};cell:{cell.coordinate}'
                    if ri == 1 and cell.column > 9:
                        coverage.append([agency,source,loc,str(cell.value),'対象外','手数料表外の数値（用途未確認。原値を保持）',''])
                        continue
                    eligible = [s for s in sections if s[0]<ri]
                    if not eligible:
                        raise ValueError(f'金額の見出し未解決: {source} {loc}')
                    header,title,headers = eligible[-1]
                    next_header = next((s[0] for s in sections if s[0]>header),sheet.max_row+1)
                    hvals = {col:merged_value(sheet,header,col)[0] for col in range(1,sheet.max_column+1)}
                    matrix = plus and (hvals.get(cell.column) in ('5G開通手数料','ネット開通手数料','電話開通手数料','タイプ1','タイプ2','タイプ3'))
                    v = lambda col: merged_value(sheet,ri,col)[0]
                    anchors = [merged_value(sheet,ri,col)[1] for col in range(2,min(sheet.max_column,9)+1)]
                    if matrix:
                        if '5G開通手数料' in headers:
                            item,typ,target = hvals[cell.column],v(2),v(3)
                        else:
                            item,typ,target = v(3),hvals[cell.column], 'ホームのみ' if 'ホームのみ' in v(3) else ''
                        isp,code,unit = '', '', '円/件'
                        remark = ''
                    else:
                        item = v(3)
                        typ = v(2) if plus and 'タイプ' in v(2) else (v(4) if plus and '対象タイプ' in headers else '')
                        target = v(3) if plus and '対象特典コード' in headers else (v(4) if not plus and cell.column==6 and 'ホーム' in v(4) or not plus and cell.column==6 and 'マンション' in v(4) else '')
                        if plus and '対象特典コード' in headers:
                            item='回線サービス'
                        code = v(4) if plus and '対象特典コード' in headers else (v(8) if not plus and cell.column==6 and '特典コード' in ' '.join(headers) else '')
                        isp = 'BIGLOBE' if 'BIGLOBE' in title or ('BIGLOBE' in v(2) and 'でんき' not in title) else ('So-net' if 'So-net' in v(2) else '')
                        unit = '円/1ID' if '/1ID' in normal(' '.join(v(c) for c in range(cell.column+1,min(sheet.max_column,9)+1))) else '原本参照'
                        if plus and '対象タイプ' in headers:
                            unit='円/件'
                        if not plus and not isp and cell.column==6:
                            prior_labels=[normal(sheet.cell(j,2).value) for j in range(1,ri+1) if sheet.cell(j,2).value]
                            markers=[t for t in prior_labels if 'BIGLOBE' in t or 'So-net' in t or 'auでんき' in t]
                            if markers:
                                isp='BIGLOBE' if 'BIGLOBE' in markers[-1] else ('So-net' if 'So-net' in markers[-1] else '')
                        remark = v(6) if plus and cell.column==5 and '対象タイプ' in headers else (v(6) if plus and cell.column==4 else (v(9) if not plus and cell.column==6 else (v(8) if not plus and cell.column==5 else '')))
                    # Notes belonging to this table, not the subsequent table; all other clauses remain linked in legal.
                    notes = '\n'.join(f'sheet:{sheet.title};row:{j} {texts[j]}' for j in range(header+1,next_header)
                                      if j in texts and not any(isinstance(c.value,(int,float)) or c.data_type=='f' for c in sheet[j])
                                      and (texts[j].lstrip().startswith('B') and j>ri or '※' in texts[j]))
                    terms = '\n'.join(x for x in [remark, notes] if x)
                    # Row coordinates are evidence, never part of business identity.
                    identity_terms = re.sub(r'sheet:[^;\n]+;row:\d+\s*|\b[A-Z]+\d+=','',terms)
                    amount = cell.value if cell.data_type!='f' else None
                    kind = '控除' if '諸経費' in title or amount is not None and amount<0 else '通常報酬'
                    hold = '数式金額は再計算値の確認が必要' if amount is None else ''
                    if not plus and cell.column==6 and isp=='BIGLOBE' and not code and '特典コード' in ' '.join(headers):
                        hold += ' 原本の特典コード欄が空欄（結合元を含む）。他行から補完せず適用範囲保留'
                    if not item or not title or not period:
                        hold += ' 成果項目・表・期間の不足を確認'
                    # A 40% boundary is explicitly contradictory in these source documents.
                    if '40%以下' in normal(legal) and '40%以上' in normal(legal):
                        hold += ' 原本の通過率40%境界が以下/以上で重複。適用判断保留'
                    c = Condition(agency,kind,product,title,item,isp,typ,target,code,'',amount,unit,
                                  '税別' if '税別' in normal(title) else '原本に税区分明示なし',
                                  period,identity_terms,legal,source,[loc,'見出し:'+f'sheet:{sheet.title};row:{header}',
                                  '結合セル根拠:'+','.join(dict.fromkeys(anchors))],digest,hold.strip())
                    conditions.append(c)
                    coverage.append([agency,source,loc,str(cell.value),'反映' if amount is not None else '保留',c.hold,c.id])
            # Monetary liabilities embedded in prose must not become reward rates.
            for r,t in texts.items():
                match=re.search(r'(\d+)\s*万円',normal(t))
                if match:
                    loc=f'sheet:{sheet.title};row:{r}'
                    c=Condition(agency,'違約金・相殺',product,'支払条件・相殺','消費者センター案件の相殺・請求',
                        amount=int(match[1])*10000,unit='円/案件',tax='原本に税区分明示なし',period=period,
                        terms='\n'.join(texts[j] for j in range(max(1,r-1),min(sheet.max_row,r+5)+1) if j in texts),
                        legal=legal,source=source,locations=[loc],digest=digest)
                    conditions.append(c)
                    coverage.append([agency,source,loc,match[0],'反映','',c.id])
    finally:
        book.close()
    return conditions,coverage,legal_rows,raw_rows


def extract_word(path,root):
    document=Document(path)
    agency,source=agency_candidate(path),path.relative_to(root).as_posix()
    digest=hashlib.sha256(path.read_bytes()).hexdigest()
    paragraphs=[(i,p.text) for i,p in enumerate(document.paragraphs,1) if p.text.strip()]
    legal='\n'.join(f'paragraph:{i} {t}' for i,t in paragraphs)
    period=next((period_from(t) for _,t in paragraphs if period_from(t)), '')
    result,coverage,raw=[],[],[]
    raw.extend([agency,source,f'paragraph:{i}',t] for i,t in paragraphs)
    heading=''
    heading_location=''
    scope=[]
    pi=ti=0
    for element in document.element.body:
        if element.tag==qn('w:p'):
            pi+=1
            text=Paragraph(element,document).text
            if any(x in text for x in ['ボリュームインセンティブ（','販促支援金（','特別調整金（']):
                heading=normal(text).lstrip('■');heading_location=f'paragraph:{pi}';scope=[]
            elif heading and text.strip():
                scope.append(f'paragraph:{pi} {text}')
        elif element.tag==qn('w:tbl'):
            ti+=1
            table=Table(element,document)
            for ri,row in enumerate(table.rows,1):
                for ci,cell in enumerate(row.cells,1):
                    if cell.text.strip():raw.append([agency,source,f'table:{ti};row:{ri};cell:{ci}',cell.text])
            if not any('金額' in cell.text for cell in table.rows[0].cells):
                continue
            typ=heading.split('（',1)[-1].rstrip('）') if '（' in heading else heading.split('(',1)[-1].rstrip(')')
            # Following paragraph carries the aggregation footnote; retain it on every tier.
            following=element.getnext()
            after=[]
            while following is not None and following.tag!=qn('w:tbl'):
                if following.tag==qn('w:p'):
                    t=Paragraph(following,document).text
                    if '※閾値' in t:after.append(t)
                    if any(x in t for x in ['ボリュームインセンティブ（','販促支援金（','特別調整金（']):break
                following=following.getnext()
            for ri,row in enumerate(table.rows[1:],2):
                threshold,amount_text=[normal(c.text) for c in row.cells[:2]]
                match=re.fullmatch(r'([\d,]+)円/件',amount_text)
                amount=int(match[1].replace(',','')) if match else None
                loc=f'table:{ti};row:{ri};cell:2'
                terms='\n'.join(scope+after)
                terms=re.sub(r'paragraph:\d+\s*','',terms)
                c=Condition(agency,'キャンペーン','auでんき' if 'でんき' in heading else 'auひかり/auひかりプラス',
                    heading,heading.split('(')[0],type=typ,threshold=threshold,amount=amount,unit='円/件',tax='税別',
                    period=period,terms=terms,legal=legal,source=source,
                    locations=[loc,f'table:{ti};row:{ri};cell:1',heading_location],digest=digest,
                    hold='' if amount is not None and heading and period else '施策・期間・金額の読取確認')
                result.append(c);coverage.append([agency,source,loc,amount_text,'反映' if amount is not None else '保留',c.hold,c.id])
    # Separate explicit monetary liabilities from reward tiers; 万円 is not a commission.
    for i,text in paragraphs:
        match=re.search(r'(\d+)\s*万円\s*/\s*1\s*事案',normal(text))
        if match:
            loc=f'paragraph:{i}'
            c=Condition(agency,'違約金','共通','別紙・違約金',normal(text).split(':')[0],amount=int(match[1])*10000,
                        unit='円/事案',tax='原本に税区分明示なし',period=period,terms=normal(text),legal=legal,
                        source=source,locations=[loc],digest=digest)
            result.append(c);coverage.append([agency,source,loc,match[0],'反映','',c.id])
    return result,coverage,[[agency,source,'本文・別紙',legal]],raw


def extract_source(path,root):
    if path.suffix=='.xlsx':return extract_excel(path,root)
    if path.suffix=='.docx':return extract_word(path,root)
    raise ValueError('未対応形式: '+path.name)


def build_ledger(root: Path, output: Path, period: str='2026-10', agency_filter: str | None=None) -> dict:
    paths=sorted(p for p in (root/'03_条件表'/period).rglob('*')
                 if p.suffix.lower() in {'.xlsx','.docx','.pdf'} and not p.name.startswith('~$')
                 and (agency_filter is None or agency_candidate(p)==agency_filter))
    if not paths:raise ValueError('解析対象資料がありません')
    if output.exists():raise FileExistsError(output)
    conditions=[];coverage=[];legal=[];raw=[];sources=[]
    for path in paths:
        before=hashlib.sha256(path.read_bytes()).hexdigest()
        cs,cv,ls,rs=extract_source(path,root)
        if hashlib.sha256(path.read_bytes()).hexdigest()!=before:raise ValueError('解析中に原本変更')
        conditions.extend(cs);coverage.extend(cv);legal.extend(ls);raw.extend(rs)
        sources.append([agency_candidate(path),path.relative_to(root).as_posix(),before,len(cs),len(cv)])
    # Same semantic condition: retain all original monetary cells and flag indistinguishable source duplicates.
    groups={}
    for c in conditions:
        if c.id in groups:
            old=groups[c.id]
            if old.amount!=c.amount:
                raise ValueError('同一条件で金額矛盾: '+c.source+' '+str(c.locations))
            old.locations.extend(c.locations)
            old.hold=(old.hold+' / '+c.hold+' / ' if old.hold or c.hold else '')+'同一成果項目・条件の原本重複。全根拠位置を保持。区分の相違は原本だけで判定できず保留'
            for cv in coverage:
                if cv[-1]==c.id:cv[4:6]=['保留',old.hold]
        else:groups[c.id]=c
    conditions=list(groups.values())
    wb=Workbook();wb.remove(wb.active)
    sheets={}
    def add(name,headers,rows):
        s=wb.create_sheet(name);s.append(headers)
        for row in rows:s.append(row)
        sheets[name]=s
        return s
    summary=dict(source_count=len(paths),agency_count=len(set(c.agency for c in conditions)),
                 condition_count=len(conditions), monetary_coverage_count=sum(r[4]!='対象外' for r in coverage),
                 held_condition_count=sum(bool(c.hold) for c in conditions),read_errors=0,
                 phase1_status='COMPLETED_WITH_EXPLICIT_HOLDS',next_phase_allowed=False)
    add('利用説明',['項目','内容'],[
        ['状態','Phase 1取込・照合用。業務承認・Phase 2以降の許可ではありません。'],
        ['条件ID','成果項目・条件に基づくローカル台帳ID。社内承認済み成果項目IDではありません。金額・行番号を含みません。'],
        ['条項','各行の条項列は元資料の支払・除外・戻入・別紙等を全文保持。表別の適用範囲は原本の表番号・条番号に従います。'],
        ['保留','原本の重複・矛盾・税区分不明を推測で確定しません。金額0も保存。控除・違約金は報酬と別区分。'],
        *[[k,str(v)] for k,v in summary.items()]])
    add('資料一覧',['代理店','根拠資料','SHA256','条件行数（統合前）','金額検証件数（対象外含む）'],sources)
    add('共通条件台帳',HEADERS,[row_values(c) for c in conditions])
    for i,agency in enumerate(sorted(set(c.agency for c in conditions)),1):
        add(f'代理店{i:02}',HEADERS,[row_values(c) for c in conditions if c.agency==agency])
    add('金額全件検証',['代理店','根拠資料','根拠位置','原本金額','反映状態','保留・対象外理由','条件ID'],coverage)
    add('支払除外戻入条項',['代理店','根拠資料','範囲','条項全文（位置付き）'],legal)
    add('控除違約金返還',['代理店','区分','成果項目・範囲','金額','単位','条件・条項','根拠資料','根拠位置・条件ID'],
        [[c.agency,c.kind,c.item,c.amount,c.unit,c.terms+'\n'+c.legal,c.source,'\n'.join(c.locations)+'\n'+c.id]
         for c in conditions if c.kind not in ('通常報酬','キャンペーン')]+
        [[agency,'返還・戻入条件（原本条項）',scope,None,'原本に従う',text,source,'支払除外戻入条項参照']
         for agency,source,scope,text in legal if re.search('返還|戻入|取戻|返金',text)])
    add('原文索引',['代理店','根拠資料','根拠位置','原文'],raw)
    add('未解決項目',['代理店','条件ID','根拠資料','根拠位置','理由'],
        [[c.agency,c.id,c.source,'\n'.join(c.locations),c.hold] for c in conditions if c.hold])
    # Validate all monetary evidence resolves to exactly one business row and value.
    for cv in coverage:
        if cv[4]=='対象外':continue
        c=groups.get(cv[-1])
        if c is None or cv[2] not in c.locations:raise ValueError('金額根拠が欠落')
        if cv[4]=='反映' and c.amount is None:raise ValueError('金額未取込')
    for s in wb:
        s.freeze_panes='F2' if s.max_column==len(HEADERS) else 'A2'
        s.auto_filter.ref=s.dimensions
        for cell in s[1]:cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='24476A')
        for row in s:
            for cell in row:
                cell.alignment=Alignment(vertical='top',wrap_text=True)
                if isinstance(cell.value,str):
                    if len(cell.value)>32767:raise ValueError('Excel上限。原文を切り捨てず停止')
                    cell.data_type='s'
        for col in s.columns:s.column_dimensions[col[0].column_letter].width=22
        if s.max_column==len(HEADERS):
            for col in ['P','Q']:s.column_dimensions[col].width=65
            s.column_dimensions['F'].width=34;s.column_dimensions['S'].width=45
            for row in s.iter_rows(min_row=2):
                row[11].number_format='#,##0;[Red]-#,##0;0'
            for r in range(2,s.max_row+1):s.row_dimensions[r].height=90
    output.parent.mkdir(parents=True,exist_ok=True)
    with output.open('xb') as stream:wb.save(stream)
    check=load_workbook(output,read_only=True,data_only=False)
    try:
        saved=list(check['共通条件台帳'].iter_rows(min_row=2,values_only=True))
        if saved != [tuple(row_values(c)) for c in conditions]:
            # openpyxl reads empty strings as None.
            norm=lambda rows:[[None if v=='' else v for v in row] for row in rows]
            if norm(saved)!=norm([row_values(c) for c in conditions]):raise ValueError('保存後内容不一致')
        if check['金額全件検証'].max_row!=len(coverage)+1:raise ValueError('検証件数不一致')
    finally:check.close()
    return summary
