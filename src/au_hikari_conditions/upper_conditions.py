"""Read the supplied upstream PDFs into evidence-backed Phase 2 review records.

Layout families are detected by annex headings, never file order or monetary equality.
Unresolved columns remain explicit review rows; dash and zero are distinct states.
"""
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import hashlib
import re
import unicodedata
from pypdf import PdfReader
from .condition_ledger import normal, period_from, candidate_key


def compact(text):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(text or ''))).replace('／','/')


def types(text):
    value=compact(text)
    if '全タイプ' in value:return 'ALL'
    if re.fullmatch(r'[123](?:,[123])*',value):return ','.join(sorted(set(value.split(','))))
    groups=re.findall(r'タイプ([123](?:[・,、][123])*)',value)
    return ','.join(sorted(set(re.findall(r'[123]',''.join(groups)))))


def dwelling(text):
    t=compact(text)
    return '/'.join(x for x in ('ホーム','マンション') if x in t)


def speed(text):
    t=compact(text).replace('ギガ','G')
    return ','.join(sorted(set(re.findall(r'(?<!\d)(10G|1G)',t))))


def codes(text):
    return ','.join(sorted(set(re.findall(r'K(?:CM|OF|WE)|BDC|NBT|CCB|ACB|YY\d+|K0\d+', compact(text)))))


def option_name(text):
    t=compact(text).replace('Wi‐Fi','Wi-Fi').replace('Wi－Fi','Wi-Fi').replace('WiFi','Wi-Fi')
    aliases={
        'So-net安心サポートプラス':'安心サポートプラス',
        'BIGLOBEお助けサポート':'BIGLOBEお助けサポート+',
        'インターネット端末保証サービス':'BIGLOBEインターネット端末保証サービス',
        'TVサービス':'TVサービス',
    }
    if t.startswith('So-netID'):return 'So-netID'
    if t.startswith('10Gインセンティブ'):return '10Gインセンティブ'
    return aliases.get(t,t).replace('（','(').replace('）',')')


@dataclass
class UpperCondition:
    product: str
    family: str
    item: str
    isp: str=''
    type: str=''
    dwelling: str=''
    speed: str=''
    access: str=''
    code: str=''
    threshold: str=''
    amount: int | None=None
    state: str='VALUE'
    unit: str='円/件'
    tax: str='税別'
    period: str=''
    terms: str=''
    source: str=''
    location: str=''
    digest: str=''
    hold: str=''
    components: tuple[str,...]=()

    @property
    def axes(self):
        return {k:getattr(self,k) for k in ('product','family','item','isp','type','dwelling','speed','access','code','threshold')}

    @property
    def id(self):
        return candidate_key('UPPER',self.axes,self.period)


def _money(text):
    t=compact(text).replace('戻入対象','')
    if t in ('-','―','−','－','ー',''):return None,'DASH'
    return int(t.replace(',','').replace('円/件','').replace('円','').replace('−','-')),'VALUE'


def _tokens(text):
    return re.findall(r'(?<![\dA-Za-z])(?:\d{1,3}(?:,\d{3})+|0|[-―−])(?![\dA-Za-z])',normal(text))


def extract_upper_pdf(path: Path, root: Path):
    source=path.relative_to(root).as_posix()
    digest=hashlib.sha256(path.read_bytes()).hexdigest()
    pages=[p.extract_text(extraction_mode='layout') or '' for p in PdfReader(path).pages]
    if any(not p.strip() for p in pages):raise ValueError('テキストを読み取れない上位PDFページがあります')
    raw=[[source,i,text] for i,text in enumerate(pages,1)]
    period=next((period_from(p) for p in pages if period_from(p)), '')
    result=[]
    def add(page, line_no, line, product, family, item, token, **kwargs):
        amount,state=_money(token)
        if not token:state='NOT_PUBLISHED'
        c=UpperCondition(product,family,item,amount=amount,state=state,period=period,
            source=source,location=f'page:{page};line:{line_no}',digest=digest,**kwargs)
        result.append(c)
        return c
    if 'キャンペーン通知書' in path.name:
        active=None;section=[]
        for pi,text in enumerate(pages,1):
            for li,line in enumerate(text.splitlines(),1):
                t=normal(line);k=compact(line)
                if k.startswith('【別紙】'):
                    active=None
                if t.startswith('■') and not any(x in k for x in ('場合','事案')):
                    active=t.lstrip('■').strip();section=[]
                if active:section.append(line)
                match=re.search(r'([\d,]+)件[~〜](.*?)([−-]?[\d,]+)円/件',k)
                if active and match:
                    threshold=str(int(match[1].replace(',','')))+'件以上'
                    kind='deduction' if active in ('後確認手数料','消費者センター案件') else 'campaign'
                    c=add(pi,li,line,'auでんき' if 'でんき' in active else 'auひかり/auひかりプラス',kind,
                          compact(active).split('(')[0],match[3],type=types(active),threshold=threshold,
                          terms='\n'.join(section),hold='集計主体・取消除外・開通条件・階層方式のレビューが必要')
                    if match[2]:c.hold+=' / 閾値と単価の間に追加条件あり';c.terms+='\n'+match[2]
                penalty=re.search(r'(\d+)\s*万円\s*/\s*1\s*事案',t)
                if penalty:
                    c=add(pi,li,line,'共通','penalty',compact(t).split(':')[0].lstrip('■'),str(int(penalty[1])*10000),
                          unit='円/事案',tax='原本に税区分明示なし',terms=t)
        # Include full campaign pages so post-table footnotes and legal conditions cannot disappear.
        for c in result:
            page=int(c.location.split(';')[0].split(':')[1]);c.terms=pages[page-1]
    else:
        for pi,text in enumerate(pages,1):
            if '【別紙' not in compact(text) and '乙商品の名称' not in compact(text):continue
            kpage=compact(text)
            lines=text.splitlines()
            annex_match=re.search(r'【別紙([1-6])】',kpage)
            annex=annex_match[1] if annex_match else ''
            # Continuation page: BIGLOBE plus options has no annex heading.
            if not annex and '乙商品の名称' in kpage:annex='5'
            if annex in ('3','6'):
                # These tables contain merged CNO/DNO, consumer cashback, and two volume columns.
                # Retain each rate row but do not guess the missing merged axes.
                for li,line in enumerate(lines,1):
                    width=max(len(x.rstrip()) for x in lines)
                    vals=[m[1] for m in re.finditer(r'[\\¥]\s*([\d,]+)',line) if m.start()>=width*0.60]
                    if not vals:continue
                    for vi,value in enumerate(vals):
                        if len(vals)>2:continue
                        prefix=re.sub(r'[\\¥]\s*[\d,]+','',normal(line)).strip()
                        c=add(pi,li,line,'auひかりプラス' if annex=='6' else 'auひかり','nifty_unresolved',
                            prefix or '回線・オプション（結合見出し未確定）',value,isp='@nifty',
                            code=codes(text),threshold=('0～99件' if vi==0 else '100件以上') if len(vals)==2 else '旧単一レート',
                            terms=text,hold='CNO/DNO・対象タイプ・速度・獲得件数階層の結合を個別確認。自動対応不可')
                        # Same text can identify several monetary rows: don't invent a business identity.
                        c.access='原文の列対応未確定'
                        c.state='UNRESOLVED';c.amount=None
                        c.hold+=' / 価格欄の原文値 '+value+' 円（受領条件未確定）'
                continue
            if annex=='5' and '乙商品の名称' not in kpage:
                # Position map for the two supplied PDF table layouts. Values always come from
                # the matched price line; coordinates/order are evidence, not condition IDs.
                rate_lines=[];table='line'
                for li,line in enumerate(lines,1):
                    k=compact(line)
                    if k.startswith('※'):table='notes'
                    if '◎回線サービスボリュームインセンティブ' in k:table='volume'
                    match=re.search(r'([\d,]+)\s*円\s*/\s*件',normal(line))
                    if match and table in ('line','volume'):rate_lines.append((li,line,match[1],table))
                modern='メニュー' in kpage
                specs=(
                    [('1','ホーム','10G','KCM',''),('1','ホーム','1G','KCM',''),
                     ('1','ホーム','10G','KOF','回線新規'),('1','ホーム','1G','KOF','回線新規'),
                     ('2','ホームルーター','','KWE',''),
                     ('3','ホーム/マンション','10G','KWE','回線新規'),('3','ホーム/マンション','10G','KWE','事業者変更'),
                     ('3','ホーム/マンション','10G','KCM','回線新規'),('3','ホーム/マンション','10G','KOF','回線新規'),
                     ('3','ホーム/マンション','1G','KWE','回線新規'),('3','ホーム/マンション','1G','KWE','事業者変更'),
                     ('3','ホーム','1G','KCM','回線新規'),('3','マンション','1G','KCM','回線新規'),
                     ('3','ホーム/マンション','1G','KOF','回線新規')]
                    if modern else
                    [('1','ホーム','1G','KOF',''),('1','ホーム','1G','KCM',''),
                     ('1','ホーム','10G','KOF',''),('1','ホーム','10G','KCM',''),
                     ('2','ホームルーター','','',''),('3','ホーム/マンション','1G','KWE',''),
                     ('3','ホーム','1G','KCM',''),('3','マンション','1G','KCM',''),
                     ('3','ホーム/マンション','10G','KWE',''),('3','ホーム/マンション','10G','KCM',''),
                     ('3','ホーム/マンション','10G','KOF','')])
                main=[r for r in rate_lines if r[3]=='line']
                if len(main)!=len(specs):raise ValueError('BIGLOBEプラスの金額行数が対応レイアウトと不一致')
                for (li,line,value,_),(typ,dw,sp,code,access) in zip(main,specs):
                    explicit_codes=codes(line)
                    if explicit_codes and explicit_codes!=code:raise ValueError('BIGLOBEプラスのコード位置が不一致')
                    add(pi,li,line,'auひかりプラス','isp_line','回線サービス',value,isp='BIGLOBE',type=typ,
                        dwelling=dw,speed=sp,code=code,access=access,terms=text)
                volumes=[r for r in rate_lines if r[3]=='volume']
                if modern and len(volumes)!=2:raise ValueError('BIGLOBEプラスの新ボリューム表が不完全')
                for (li,line,value,_),sp in zip(volumes,('10G','1G')):
                    add(pi,li,line,'auひかりプラス','isp_volume','ボリュームインセンティブ',value,isp='BIGLOBE',
                        type='1',dwelling='ホーム',speed=sp,threshold='1件以上',terms=text)
                continue
            if annex=='2' or annex=='5' and '乙商品の名称' in kpage:
                start=next((i for i,line in enumerate(lines) if '◎オプションサービス' in compact(line)),0)
                rate_lines=[(i+1,line,re.search(r'([\d,]+)\s*円\s*/\s*件',normal(line)))
                            for i,line in enumerate(lines) if i>=start and re.search(r'([\d,]+)\s*円\s*/\s*件',normal(line))]
                labels=(['トータル・ネットセキュリティ','BIGLOBEお助けサポート+',
                         'BIGLOBEあんしんバックアップ(容量無制限AOSBOX)','生活かけつけレスキュー',
                         'BIGLOBEインターネット端末保証サービス','暮らしのあんしん相談','エンジョイプラス']
                        if annex=='2' else ['コラボ光電話','コラボ光電話A','コラボ光テレビ',
                         'BIGLOBEあんしんパック','トータル・ネットセキュリティ','BIGLOBEインターネット端末保証サービス',
                         'BIGLOBEお助けサポート+','U-NEXTforBIGLOBE','エンジョイプラス',
                         'BIGLOBEあんしんバックアップ','暮らしのあんしん相談','生活かけつけレスキュー'])
                if len(rate_lines)!=len(labels):raise ValueError('BIGLOBEオプションの金額行数がレイアウトと不一致')
                for (li,line,match),label in zip(rate_lines,labels):
                    dw=('ホーム' if label=='生活かけつけレスキュー' else 'マンション' if label=='BIGLOBEインターネット端末保証サービス' else 'ホーム/マンション') if annex=='2' else ''
                    add(pi,li,line,'auひかり' if annex=='2' else 'auひかりプラス','isp_option',label,match[1],isp='BIGLOBE',
                        type=('3' if label in ('コラボ光電話','コラボ光電話A','コラボ光テレビ') else 'ALL') if annex=='5' else '',
                        dwelling=dw,code=codes(line) if annex=='2' else '',terms=text)
                if annex=='5':continue
            section=''
            for li,line in enumerate(lines,1):
                t=normal(line);k=compact(line)
                if k.startswith('表'):section=k.split('■')[0]
                if '◎回線サービス' in k:section='isp_volume' if 'ボリューム' in k else 'isp_line'
                if '◎オプションサービス' in k:section='isp_option'
                if annex=='1':
                    if k.startswith(('auひかりホーム','auひかりマンション')):
                        vals=_tokens(re.sub(r'^au\s*ひかり(?:ホーム|マンション)','',t))
                        if len(vals) not in (4,5):raise ValueError(f'通常回線の列を解釈できません: {source} page:{pi} {t}')
                        dw=dwelling(t)
                        for item,val in zip(('application','net_open_component','opening_incentive','phone_application','phone_open') if len(vals)==5 else ('application','net_open','phone_application','phone_open'),vals):
                            add(pi,li,line,'auひかり','base',item,val,dwelling=dw,terms=text)
                        # Explicit components; old agency single net commission compares against their sum only as a candidate.
                        nums=[_money(vals[i])[0] for i in (1,2)] if len(vals)==5 else []
                        if nums and all(v is not None for v in nums):
                            c=add(pi,li,line,'auひかり','base','net_open',str(sum(nums)),dwelling=dw,terms=text,
                                  hold='旧回線単価と上位の開通手数料＋開通インセンティブの合算対応を要承認')
                            c.components=tuple(x.id for x in result[-6:-1] if x.item in ('net_open_component','opening_incentive'))
                    elif section in ('表2','表3') and ('ISP' in k or 'So-net' in k) and _tokens(t):
                        vals=_tokens(t);name=re.split(r'\s{2,}',line.strip())
                        label='10Gインセンティブ' if 'インセンティブ' in k else ('おうちどこでもWi-Fi' if 'Wi-Fi' in k else ('So-netID' if 'So-netID' in k else ('S-SAFE' if 'S-SAFE' in k else '安心サポートプラス')))
                        add(pi,li,line,'auひかり','option',label,vals[-1],isp='So-net' if section=='表3' else '',dwelling=dwelling(t),terms=text)
                    elif section=='表4' and ('Mプラン' in k or 'Lプラン' in k):
                        add(pi,li,line,'auでんき','electric','Mプラン' if 'Mプラン' in k else 'Lプラン',_tokens(t)[-1],terms=text)
                elif annex=='4':
                    if k.startswith('auひかりプラスタイプ'):
                        typ=types(t);tail=re.split(r'タイプ\s*[123]',t,maxsplit=1)[-1];vals=_tokens(tail)
                        if len(vals)==4:vals.append('')
                        if len(vals)!=5:raise ValueError(f'プラス回線の列を解釈できません: {source} page:{pi} {t}')
                        for item,val in zip(('application','5g_open','ftth_open','net_open','phone_open'),vals):
                            add(pi,li,line,'auひかりプラス','base',item,val,type=typ,
                                dwelling='ホーム' if typ=='1' else 'ホーム/マンション',terms=text,
                                hold='FTTH受領分の代理店への支払可否は別途承認' if item=='ftth_open' else '')
                    elif section=='表2' and k.startswith('ISPを問わない。'):
                        vals=_tokens(t)
                        if len(vals)!=3:raise ValueError('プラスオプションの3タイプ列を解釈できません')
                        label='Wi-Fiパック' if 'Wi-Fiパック' in k else ('メッシュWi-Fi子機レンタル' if 'メッシュ' in k else ('TVサービス' if 'TVサービス' in k else '10Gインセンティブ'))
                        for typ,val in zip(('1','2','3'),vals):
                            add(pi,li,line,'auひかりプラス','option',label,val,type=typ,
                                dwelling='ホーム' if label=='10Gインセンティブ' else '',terms=text)
                    elif section=='表3' and k.startswith('So-net(ISP)'):
                        vals=_tokens(t);label='So-netID' if 'So-netID' in k else ('ボリュームインセンティブ' if 'ボリューム' in k else ('S-SAFE' if 'S-SAFE' in k else '安心サポートプラス'))
                        add(pi,li,line,'auひかりプラス','option',label,vals[-1],isp='So-net',type='1,3',speed=speed(t),terms=text)
                elif annex=='2' and section=='isp_line':
                    if re.search(r'[\d,]+\s*円\s*/\s*件',t):
                        label='ホーム' if k.startswith('ホーム') else 'マンション'
                        add(pi,li,line,'auひかり','isp_line','回線サービス',re.search(r'([\d,]+)\s*円',t)[1],isp='BIGLOBE',dwelling=label,code='BDC',terms=text)
                elif annex=='2' and section=='isp_volume' and re.search(r'[\d,]+\s*円',t):
                    add(pi,li,line,'auひかり','isp_volume','ボリュームインセンティブ',re.search(r'([\d,]+)\s*円',t)[1],isp='BIGLOBE',dwelling='ホーム/マンション',code='BDC',threshold='250件以上',terms=text)
    # Every page is retained. Stable business-key collisions become an explicit unresolved record,
    # not an overwritten value or a fabricated unique condition ID.
    seen={}
    unique=[]
    for c in result:
        if c.id in seen:
            old=seen[c.id]
            old.hold+=' / 同一キーの上位原文が複数。個別列の対応未確定'
            old.state='UNRESOLVED';old.amount=None
            old.location+='; '+c.location
        else:seen[c.id]=c;unique.append(c)
    if not unique:raise ValueError('上位条件を抽出できませんでした: '+source)
    if hashlib.sha256(path.read_bytes()).hexdigest()!=digest:raise ValueError('解析中に上位原本が変更されました')
    return unique,raw
