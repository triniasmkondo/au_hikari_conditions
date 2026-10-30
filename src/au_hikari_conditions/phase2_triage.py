"""Unapproved arithmetic references and a traceable review work queue for Phase 2.

A price spread is not an approved gross margin. Grouping questions never approves,
resolves or drops any of the underlying condition rows.
"""
from __future__ import annotations
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from datetime import date
from decimal import Decimal, InvalidOperation
import hashlib
import json

from .upper_conditions import compact
from .p2 import _parse_target_period


VERSION='phase2-triage-v1'


def stable_id(prefix,values):
    return prefix+'-'+hashlib.sha256(json.dumps(values,sort_keys=True,ensure_ascii=False,
        default=str,separators=(',',':')).encode()).hexdigest()[:24]


def number(value):
    if value is None or value=='' or isinstance(value,bool):return None
    try:
        n=Decimal(str(value))
        return n if n.is_finite() else None
    except InvalidOperation:return None


def active(period,month):
    try:
        first,last=map(date.fromisoformat,period.split('/'));start,end=_parse_target_period(month)
        return first<=start and last>=end
    except (ValueError,AttributeError):return False


@dataclass(frozen=True)
class InitialCandidate:
    match_state: str
    spread: Decimal | None
    margin_state: str
    candidate: Decimal | None
    candidate_state: str
    formula: str
    assumptions: str
    missing: str
    exclusion: str
    evidence: str


def initial_candidate(p,chosen,previous_period,target_period):
    old,up=p.old,p.previous
    failures=[]
    if getattr(p,'initial_identity_collision',False):failures.append('同じ条件軸に代理店旧条件が複数あり、金額の一対一対応が未確定')
    if not old:failures.append('新規採用候補には10月代理店単価がない')
    if not up:failures.append('10月上位が同一条件軸で一意に対応しない')
    if old and old.kind not in ('通常報酬','キャンペーン'):failures.append('控除・違約金・相殺は粗利候補の対象外')
    if old and old.hold:failures.append('Phase 1保留が未解消')
    if old and old.tax!='税別':failures.append('10月代理店の税区分未確定')
    if up and (up.axes!=p.axes or up.state!='VALUE' or up.tax!='税別' or up.components or 'unresolved' in up.family):
        failures.append('10月上位の条件・金額・税・コンポーネントが未確定')
    benign_review='集計主体・取消除外・開通条件・階層方式のレビューが必要'
    if up and up.hold and up.hold!=benign_review:failures.append('10月上位に個別の構造・適用条件保留あり')
    if old and up and old.unit!=up.unit:failures.append('単位表記の同等性未確認（'+old.unit+' / '+up.unit+'）')
    if old and not active(old.period,previous_period):failures.append('10月代理店の期間が対象月を含まない')
    if up and not active(up.period,previous_period):failures.append('10月上位の期間が対象月を含まない')
    old_amount=number(old.amount) if old else None
    old_up=number(up.amount) if up else None
    if old_amount is None or old_up is None or old_amount<0 or old_up<0:failures.append('非負の10月両単価が揃わない')
    evidence=''
    if old and up:
        evidence=(old.source+'#'+'; '.join(old.locations)+' SHA256='+old.digest+'\n'+
                  up.source+'#'+up.location+' SHA256='+up.digest)
    if failures:
        return InitialCandidate('算出対象外',None,'未算出',None,'提示不可','','',
            ' / '.join(dict.fromkeys(failures)),' / '.join(dict.fromkeys(failures)),evidence)
    spread=old_up-old_amount
    assumptions=('実粗利10＝10月差額候補＋追加受領10－追加支払10－直接原価10。'
        '各追加報酬・原価・返還/回収条件は未確認。差額候補を承認済み粗利とは扱わない。')
    if spread==0:assumptions+=' 差額0円でも実粗利0円とは限らず、費用等があれば実粗利が負になる可能性がある。'
    missing=('直接原価10/11、追加受領10/11、追加支払10/11、支払・除外・戻入・回収条件の同等性、'
             '階層・併給排他の該当性、粗利基準と最終条件の正式承認')
    quote_fail=[]
    if not chosen or chosen.axes!=p.axes:quote_fail.append('11月は同一条件軸で一意に対応しない')
    if chosen and (chosen.state!='VALUE' or chosen.tax!='税別' or chosen.components or 'unresolved' in chosen.family):
        quote_fail.append('11月上位の金額・税・コンポーネントが未確定')
    if chosen and chosen.hold and chosen.hold!=benign_review:quote_fail.append('11月上位に個別の構造・適用条件保留あり')
    if chosen and chosen.unit!=up.unit:quote_fail.append('11月上位の単位が一致しない')
    if chosen and not active(chosen.period,target_period):quote_fail.append('11月上位の期間が対象月を含まない')
    if spread<0:quote_fail.append('10月単価差が負。追加報酬等の確認を先行')
    new_up=number(chosen.amount) if chosen else None
    if new_up is None or new_up<0:quote_fail.append('11月上位の非負単価未確定')
    candidate=None
    formula='10月差額候補＝'+str(old_up)+'－'+str(old_amount)+'＝'+str(spread)
    if not quote_fail:
        candidate=old_amount+new_up-old_up
        if candidate<0:quote_fail.append('差額維持による候補単価が負');candidate=None
        else:formula+=' / 11月候補＝'+str(old_amount)+'＋('+str(new_up)+'－'+str(old_up)+')＝'+str(candidate)
    if chosen:
        evidence+='\n'+chosen.source+'#'+chosen.location+' SHA256='+chosen.digest
    if candidate is not None:
        assumptions+=(' 11月候補は「(直接原価11－直接原価10)＋(追加支払11－追加支払10)'
            '－(追加受領11－追加受領10)＝0」かつ支払・除外・戻入等が同等と確認された場合のみ差額を維持する条件付き案。'
            '月差0を事実として入力・確定していない。')
    return InitialCandidate('10月の成果項目・条件軸・両単価・単位・税・対象月が一意に対応',spread,
        '暫定：費用等控除前の差額（実粗利未算出）',candidate,
        '条件付き・未確定（原価/追加報酬の月差等未反映）' if candidate is not None else '提示不可',
        formula,assumptions,missing,' / '.join(quote_fail),evidence)


@dataclass(frozen=True)
class Issue:
    id: str
    scope: str
    code: str
    label: str
    question: str
    key: str


def _issue(scope,code,label,key,question):
    serialized=json.dumps(key,sort_keys=True,ensure_ascii=False,default=str)
    return Issue(stable_id('Q',[scope,code,key]),scope,code,label,question,serialized)


def proposal_issues(result):
    p=result['proposal'];old=p.old;u=result['chosen'] or p.upper
    issues=[]
    def common(code,label,key,question):issues.append(_issue('共通',code,label,key,question))
    def individual(code,label,key,question):issues.append(_issue('代理店個別',code,label,[p.agency,key],question))
    if u and 'unresolved' in u.family:
        common('UPPER_LAYOUT','上位の結合見出し・CNO/DNO・階層が未確定',
            [u.source,u.family,u.product],
            '原本発行元に対象タイプ・コード・獲得件数階層の対応を確認し、上位条件の構造を確定する。顧客還元額を報酬と混同しない。')
    if u and u.state in ('DASH','NOT_PUBLISHED'):
        common('UPPER_AMOUNT_STATE','上位のダッシュ・未掲載の支払可否',
            [u.source,u.product,u.family,u.item,u.state],
            '該当するタイプ別のダッシュ/未掲載が支払対象外・未公表等のどれに当たるか確認する。0円・終了とは補完しない。')
    if old and len(p.candidates)>1:
        axes=dict(p.axes);axes.pop('access',None)
        common('SCOPE_SPLIT','複数候補への条件分岐・住居区分',axes,
            '新規/事業者変更、住居等の分岐を確認し、旧条件のどの範囲に各上位行が対応するか確定する。')
    if old and old.hold:
        if '40%' in compact(old.hold):
            common('BOUNDARY_40','原本の通過率40%境界が重複',
                ['40%以下','40%以上','前後確通過率'],
                '原本発行元に40%ちょうどの適用先と対象代理店・契約範囲を確認する。')
        if '特典コード欄が空欄' in old.hold:
            individual('AGENCY_CODE_BLANK','代理店原本の特典コード空欄',
                [p.axes,old.product],
                '代理店契約で空欄がどの特典コードに適用されるか確認し、根拠を記録する。他社・他行から補完しない。')
        if '原本重複' in old.hold:
            individual('AGENCY_DUPLICATE','代理店原本の同一条件重複',
                [p.axes,old.source,old.amount],
                '複数の元金額欄が同一報酬か別区分かを確認する。重複して支払わない。')
        if not any(s in old.hold for s in ('40%','特典コード欄が空欄','原本重複')):
            individual('PHASE1_OTHER','Phase 1個別保留',
                [p.axes,old.hold],old.hold)
    if old and not p.candidates:
        individual('NO_UPPER_MATCH','代理店独自条件・上位対応なし',
            [p.axes,old.source,old.amount,old.terms],
            '旧条件の継続/変更/終了を代理店契約または社内経費根拠で確認する。上位未掲載のみで終了させない。')
    if not old and not p.exact and u:
        # Find a common question shared by the old row and its split upper counterparts.
        axes=dict(u.axes);axes.pop('access',None)
        if any('access' in x.axes and x.axes['access'] for x in p.candidates):
            common('SCOPE_SPLIT','複数候補への条件分岐・住居区分',axes,
                '新規/事業者変更、住居等の分岐を確認し、旧条件のどの範囲に各上位行が対応するか確定する。')
        else:
            common('UPPER_RESHAPE','上位条件と既存行の分割・対象範囲変更',
                [u.source,u.axes],
                '旧条件との範囲の違いを定義し、既存行への反映か別条件として採用するかの選択肢を整理する。')
        individual('AGENCY_ADOPTION','代理店での新規・分割条件の採用判断',
            [u.product],
            '共通の条件定義を確認後、対象明細ごとに採用/非採用・分割・重複防止を判断する。論点をまとめても条件別承認は必要。')
    if not old and u and u.hold and u.family=='campaign':
        common('CAMPAIGN_RULE','キャンペーンの集計主体・段階方式・併給排他',
            [u.source,u.product,u.item,u.type],
            '計数主体・対象案件・取消除外・到達段階/累進・併給排他を確認する。代理店ごとの適用確認は別に行う。')
    if u and compact(p.agency).casefold()=='keioh株式会社' and u.family=='base' and u.item=='ftth_open':
        individual('PAYMENT_PROHIBITED','keioh FTTHは上位受領のみ・代理店支払禁止',
            [u.product,u.item],
            '既存の支払禁止ルールを確認し、上位受領分を代理店支払へ加算しない。')
    if old and old.tax!='税別' or u and u.tax!='税別':
        common('TAX_BASIS','原本の税区分・課税範囲が未明示',
            [p.axes['product'],p.axes['family'],old.kind if old else '上位のみ'],
            '原本発行元・経理に当該表の税区分と対象範囲を確認する。各代理店への適用・換算根拠は条件別に承認する。')
    if old and p.previous and old.unit!=p.previous.unit:
        common('UNIT_EQUIVALENCE','1ID・件等の単位同等性が未確認',
            [p.axes['product'],p.axes['family'],old.unit,p.previous.unit],
            '原本の1IDと件が同一の成果を数える単位か確認する。確認前に換算・粗利逆算しない。')
    if not old and u and not any(i.code=='AGENCY_ADOPTION' for i in issues):
        individual('AGENCY_ADOPTION','代理店での新規・分割条件の採用判断',
            [u.product],
            '共通の条件定義を確認後、対象明細ごとに採用/非採用・分割・重複防止を判断する。論点をまとめても条件別承認は必要。')
    if any(b.startswith('粗利・原価:') for b in result['blocks']):
        individual('FINANCIAL_EVIDENCE','原価・追加報酬と粗利基準の明細回収',
            ['finance-input'],
            '10月/11月の条件別直接原価・追加受領・追加支払を回収し、初期差額と実粗利を照合する。共通0円で補完せず、根拠付き粗利基準を条件別に承認する。')
    if result['category']=='要確認' and not issues:
        individual('OTHER_REVIEW','その他の個別対応確認',[p.axes,p.reason],p.reason)
    if result['category']=='要確認' and all(i.code in ('FINANCIAL_EVIDENCE','AGENCY_ADOPTION') for i in issues):
        issues.insert(0,_issue('代理店個別','OTHER_REVIEW','その他の個別対応確認',[p.agency,p.axes,p.reason],p.reason))
    return list({i.id:i for i in issues}.values())


PREVIEW_HEADERS=['10月同条件照合','10月差額・初期粗利候補','粗利候補の状態',
    '11月条件付き候補単価','候補単価の状態','候補の算式','候補成立の前提（未確認）',
    '不足情報','候補を提示しない理由','初期候補の原本根拠','共通論点ID','代理店個別論点ID']


def preview_values(result):
    c=result['initial_candidate'];issues=result['review_issues']
    return [c.match_state,c.spread,c.margin_state,c.candidate,c.candidate_state,c.formula,c.assumptions,
            c.missing,c.exclusion,c.evidence,';'.join(i.id for i in issues if i.scope=='共通'),
            ';'.join(i.id for i in issues if i.scope=='代理店個別')]


def build_triage(results,previous_period,target_period):
    groups={};members=defaultdict(list)
    for r in results:
        p=r['proposal']
        r['initial_candidate']=initial_candidate(p,r['chosen'],previous_period,target_period)
        r['review_issues']=proposal_issues(r)
        for issue in r['review_issues']:
            groups[issue.id]=issue;members[issue.id].append(r)
    baseline=[r for r in results if r['category']=='要確認']
    baseline_ids={r['proposal'].id for r in baseline}
    covered={r['proposal'].id for values in members.values() for r in values if r['proposal'].id in baseline_ids}
    baseline_groups={i.id:i for r in baseline for i in r['review_issues']}
    if covered!={r['proposal'].id for r in baseline}:raise ValueError('要確認行の論点整理に欠落')
    primary=Counter(r['review_issues'][0].code for r in baseline)
    stats=dict(original_review_rows=len(baseline),covered_review_rows=len(covered),
        common_questions=sum(i.scope=='共通' for i in baseline_groups.values()),
        individual_questions=sum(i.scope=='代理店個別' for i in baseline_groups.values()),
        review_question_count=len(baseline_groups),question_count_reduction=len(baseline)-len(baseline_groups),
        all_review_question_count=len(groups),
        unresolved_review_rows=len(baseline),resolved_by_grouping=0,
        primary_reason_counts=dict(primary),
        initial_spread_count=sum(r['initial_candidate'].spread is not None for r in results),
        conditional_price_count=sum(r['initial_candidate'].candidate is not None for r in results),
        formal_approved_count=sum(r['status']=='APPROVED' for r in results))
    return groups,members,stats


def write_triage(book,results,groups,members,stats,write_sheet):
    write_sheet(book,'確認ダッシュボード',['項目','内容'],[
        ['最初に見る順序','共通確認論点 → 代理店個別論点 → 粗利初期候補 → 担当者確認一覧 → 黄色の承認入力'],
        ['確認単位','論点シートは全条件案の不足事項を収録。削減指標は元の要確認行を含む論点だけを数える。要確認行を原文・対象範囲に基づく論点へ整理。共通の判断と代理店の採用判断を分ける。対象明細シートから全行へ戻れる。'],
        ['削減の意味','論点の重複確認を減らした件数。条件別の承認や未解決行を削除・解消した件数ではない。共通論点も対象代理店の適用確認が必要。'],
        ['金額の区別','10月差額は費用等控除前の初期粗利候補。11月の条件付き候補は原価・追加報酬の月差等を未反映。正式提案・承認額とは別列。'],
        ['次の入力','直接原価10/11・追加受領/追加支払10/11の明細を回収し、実粗利と条項同等性を確認。承認可能な粗利基準・11月直接原価を黄色の入力へ記録する。'],
        *[[k,json.dumps(v,ensure_ascii=False)] for k,v in stats.items()]])
    baseline=[r for r in results if r['category']=='要確認']
    codes_in_baseline=sorted({i.code for r in baseline for i in r['review_issues']})
    write_sheet(book,'要確認理由別集計',['理由コード','理由','主理由の行数（排他集計）','副理由を含む延べ対象行数',
        '共通論点数','代理店個別論点数'],[
        [code,next(i.label for r in baseline for i in r['review_issues'] if i.code==code),
         sum(r['review_issues'][0].code==code for r in baseline),
         sum(any(i.code==code for i in r['review_issues']) for r in baseline),
         len({i.id for r in baseline for i in r['review_issues'] if i.code==code and i.scope=='共通'}),
         len({i.id for r in baseline for i in r['review_issues'] if i.code==code and i.scope=='代理店個別'})]
         for code in codes_in_baseline])
    headers=['論点ID','理由コード','論点','確認する内容','対象代理店数','対象代理店','対象行数',
             '代表根拠','状態','対象範囲キー','元要確認の対象行数','削減指標の対象論点']
    for scope,name in [('共通','共通確認論点'),('代理店個別','代理店個別論点')]:
        rows=[]
        for issue in sorted(groups.values(),key=lambda i:(i.code,i.key)):
            if issue.scope!=scope:continue
            rs=members[issue.id];agencies=sorted({r['proposal'].agency for r in rs})
            refs=[]
            for r in rs:
                p=r['proposal'];u=r['chosen'] or p.upper
                for ref in ([p.old.source] if p.old else [])+[x.source+'#'+x.location for x in p.candidates]:
                    if ref not in refs:refs.append(ref)
            rows.append([issue.id,issue.code,issue.label,issue.question,len(agencies),' / '.join(agencies),
                         len(rs),'\n'.join(refs[:3]),'未解決・共通判断の自動適用なし',issue.key,
                         sum(r['category']=='要確認' for r in rs),'対象' if any(r['category']=='要確認' for r in rs) else 'その他の確認'])
        write_sheet(book,name,headers,rows)
    detail=[]
    for issue in groups.values():
        for r in members[issue.id]:
            p=r['proposal'];u=r['chosen'] or p.upper
            detail.append([issue.id,issue.scope,issue.code,p.agency,p.id,p.old.id if p.old else '',p.new_id,
                p.old.item if p.old else p.axes['item'],r['category'],
                p.old.source+'#'+'; '.join(p.old.locations) if p.old else '',
                '\n'.join(x.source+'#'+x.location for x in p.candidates),p.old.hold if p.old else ''])
    write_sheet(book,'論点対象明細',['論点ID','範囲','理由コード','代理店','提案ID','10月条件ID','11月条件ID',
        '成果項目','分類案','旧原本根拠','上位原本根拠','Phase1保留'],detail)
    write_sheet(book,'粗利初期候補',['代理店','提案ID','10月条件ID','成果項目','タイプ','10月上位単価',
        '10月代理店単価','11月上位単価']+PREVIEW_HEADERS,
        [[r['proposal'].agency,r['proposal'].id,r['proposal'].old.id,r['proposal'].old.item,
          r['proposal'].axes['type'],r['proposal'].previous.amount,r['proposal'].old.amount,
          r['chosen'].amount if r['chosen'] else None]+preview_values(r)
         for r in results if r['initial_candidate'].spread is not None])
    write_sheet(book,'担当者確認一覧',['優先順','代理店','提案ID','10月条件ID','11月条件ID','成果項目','分類案',
        '担当者が確認する内容','正式提案単価','最終状態']+PREVIEW_HEADERS,
        [[1 if any(i.scope=='共通' for i in r['review_issues']) and r['category']=='要確認' else
          2 if r['category']=='要確認' else 3,
          r['proposal'].agency,r['proposal'].id,r['proposal'].old.id if r['proposal'].old else '',
          r['proposal'].new_id,r['proposal'].old.item if r['proposal'].old else r['proposal'].axes['item'],
          r['category'],' / '.join(dict.fromkeys(i.question for i in r['review_issues'])) or
          '費用・追加報酬と条項同等性を確認し、粗利基準と単価を個別承認する',r['proposed'],r['status']]+preview_values(r)
         for r in sorted(results,key=lambda r:(0 if r['category']=='要確認' and any(i.scope=='共通' for i in r['review_issues']) else 1 if r['category']=='要確認' else 2,r['proposal'].agency,r['proposal'].id))])
    write_sheet(book,'代理店別確認負荷',['代理店','元の要確認行','共通論点への参照数','個別論点数',
        '10月初期差額候補','11月条件付き単価候補','未解決行'],[
        [agency,sum(r['category']=='要確認' and r['proposal'].agency==agency for r in results),
         sum(i.scope=='共通' and any(r['proposal'].agency==agency and r['category']=='要確認' for r in members[i.id]) for i in groups.values()),
         sum(i.scope=='代理店個別' and any(r['proposal'].agency==agency and r['category']=='要確認' for r in members[i.id]) for i in groups.values()),
         sum(r['proposal'].agency==agency and r['initial_candidate'].spread is not None for r in results),
         sum(r['proposal'].agency==agency and r['initial_candidate'].candidate is not None for r in results),
         sum(r['category']=='要確認' and r['proposal'].agency==agency for r in results)]
        for agency in sorted({r['proposal'].agency for r in results})])
