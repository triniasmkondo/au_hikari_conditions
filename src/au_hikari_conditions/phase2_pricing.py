"""Agency pricing policies and evidence-bound draft approvals, without margin claims."""
from collections import Counter
from dataclasses import asdict
import json
from openpyxl import load_workbook
from .phase2_triage import active, number, stable_id
from .upper_conditions import compact


def read_sheet(path,name,headers,key):
    if not path:return {}
    book=load_workbook(path,read_only=True,data_only=False)
    try:
        if name not in book:return {}
        sheet=book[name]
        if list(next(sheet.values))!=headers:raise ValueError(name+'の列が不一致')
        rows={}
        for cells in sheet.iter_rows(min_row=2):
            if any(c.data_type=='f' for c in cells):raise ValueError(name+'に数式は使用できません')
            r=dict(zip(headers,[c.value for c in cells]))
            if not any(v is not None for v in r.values()):continue
            k=tuple(r[x] for x in key)
            if k in rows:raise ValueError(name+'のキー重複')
            rows[k]=r
        return rows
    finally:book.close()


def identity_errors(p,u,previous_period,target_period):
    old=p.old;prev=p.previous;errors=[]
    if not old or not prev:return ['新規・10月対応なし']
    if not p.exact or len(p.candidates)!=1 or p.initial_identity_collision or p.axes!=u.axes or prev.axes!=u.axes:
        errors.append('条件軸が一意に一致しない')
    if old.hold:errors.append('10月条件に保留あり')
    for v,label in ((old,'10月代理店'),(prev,'10月上位'),(u,'11月上位')):
        if v.tax!='税別' or v.unit!=old.unit:errors.append(label+'の税・単位未確定/不一致')
        n=number(v.amount)
        if n is None or n<0:errors.append(label+'の非負単価不明')
    for v,month in ((old,previous_period),(prev,previous_period),(u,target_period)):
        if not active(v.period,month):errors.append('対象月全体の適用期間未確認')
    for v in (prev,u):
        if v.state!='VALUE' or v.components or 'unresolved' in v.family:errors.append('上位構造・金額状態未確定')
        if v.hold and v.hold!='集計主体・取消除外・開通条件・階層方式のレビューが必要':errors.append('上位に適用保留あり')
    if u.family in ('deduction','penalty'):errors.append('控除・違約金は単価案の対象外')
    if compact(p.agency).casefold()=='keioh株式会社' and u.family=='base' and u.item=='ftth_open':errors.append('代理店支払禁止')
    return sorted(set(errors))


def write_pricing(book,results,current,snapshot,review_input,write_sheet,previous_period,target_period):
    from .phase2_review import APPROVAL_FIELDS,input_hash,approval_errors,old_axes
    schemas={
        '単価共通同等性入力':(['upper_id','scope_key','equivalence_status','resolution','snapshot_hash']+APPROVAL_FIELDS,('upper_id',)),
        '単価基本方針入力':(['agency','method','snapshot_hash']+APPROVAL_FIELDS,('agency',)),
        '成果項目例外入力':(['agency','product','family','item','method','snapshot_hash']+APPROVAL_FIELDS,('agency','product','family','item')),
        '単価同等性入力':(['agency','upper_id','old_id','equivalence_status','resolution','snapshot_hash']+APPROVAL_FIELDS,('agency','upper_id')),
        '方式別単価承認入力':(['agency','upper_id','manual_amount','candidate_hash','snapshot_hash']+APPROVAL_FIELDS,('agency','upper_id'))}
    saved={n:read_sheet(review_input,n,*s) for n,s in schemas.items()}
    agencies=sorted({r['proposal'].agency for r in results});allowed={(a,u.id) for a in agencies for u in current}
    for n in ('単価同等性入力','方式別単価承認入力'):
        if set(saved[n])-allowed:raise ValueError(n+'の対象外条件')
    if set(saved['単価基本方針入力'])-{(a,) for a in agencies}:raise ValueError('対象外代理店方針')
    exception_keys={(a,u.product,u.family,u.item) for a in agencies for u in current}
    if set(saved['成果項目例外入力'])-exception_keys:raise ValueError('対象外の成果項目例外')
    def record(name,key,fields):
        r=saved[name].get(key) or dict.fromkeys(schemas[name][0])
        if key in saved[name] and r['snapshot_hash']!=snapshot:raise ValueError(name+'の資料不一致')
        r.update(fields,snapshot_hash=snapshot)
        if name in ('単価基本方針入力','成果項目例外入力') and r.get('method') not in (('A','B','C',None,'') if name=='成果項目例外入力' else ('A','B','C')):
            raise ValueError('方式はA/B/Cで指定してください')
        r['record_hash']=input_hash(name,r);return r
    policies={a:record('単価基本方針入力',(a,),{'agency':a, 'method':saved['単価基本方針入力'].get((a,),{}).get('method') or 'A'}) for a in agencies}
    exceptions={k:record('成果項目例外入力',k,dict(zip(schemas['成果項目例外入力'][0][:4],k))) for k in sorted(exception_keys)}
    common_records={}
    if set(saved['単価共通同等性入力'])-{(u.id,) for u in current}:raise ValueError('対象外の共通単価条件')
    for u in current:
        scope=json.dumps([(a,sorted(r['proposal'].old.id for r in results if r['proposal'].agency==a and r['proposal'].old and old_axes(r['proposal'].old)==u.axes)) for a in agencies],ensure_ascii=False)
        old_common=saved['単価共通同等性入力'].get((u.id,))
        if old_common and old_common.get('scope_key')!=scope:raise ValueError('共通同等性の対象代理店・10月条件が不一致')
        common_records[u.id]=record('単価共通同等性入力',(u.id,),dict(upper_id=u.id,scope_key=scope))
    output=[];identities=[];approvals=[];audit=[];summary=[];tasks=[]
    for agency in agencies:
        counts=Counter();old_results=[r for r in results if r['proposal'].agency==agency and r['proposal'].old]
        for u in current:
            key=(agency,u.id);ek=(agency,u.product,u.family,u.item)
            policy=exceptions[ek] if exceptions[ek].get('method') else policies[agency]
            policy_name='成果項目例外入力' if exceptions[ek].get('method') else '単価基本方針入力'
            requested=policy['method']
            matches=[r for r in old_results if old_axes(r['proposal'].old)==u.axes]
            p=matches[0]['proposal'] if len(matches)==1 else None
            errors=identity_errors(p,u,previous_period,target_period) if p else ['新規・対応不明・複数の10月条件']
            eq=record('単価同等性入力',key,dict(agency=agency,upper_id=u.id))
            expected_old=p.old.id if p else ''
            if eq.get('old_id') not in (None,'',expected_old):raise ValueError('同等性判断の10月条件ID不一致')
            eq['old_id']=expected_old;eq['record_hash']=input_hash('単価同等性入力',eq)
            # Conditions, exclusions and reversal rules cannot be proven from matching price axes.
            common_eq=common_records[u.id]
            common_valid=common_eq.get('equivalence_status')=='CONFIRMED' and bool(common_eq.get('resolution')) and not approval_errors('単価共通同等性入力',common_eq,snapshot)
            equivalent=eq.get('equivalence_status')=='CONFIRMED' and bool(eq.get('resolution')) and not approval_errors('単価同等性入力',eq,snapshot)
            equivalent=equivalent or common_valid
            method=requested if requested in ('A','B') and not errors and equivalent else 'C'
            reason=list(errors)
            if not equivalent:reason.append('適用・支払・除外・戻入・個別約定の同等性を根拠付きで確認')
            price=None;formula='';delta=None
            if p and not errors:
                delta=number(u.amount)-number(p.previous.amount)
            if method=='A':price=number(p.old.amount);formula='10月代理店単価を継承'
            elif method=='B':price=number(p.old.amount)+delta;formula='10月代理店単価＋(11月上位単価－10月上位単価)'
            if price is not None and price<0:reason.append('算出額が負');price=None;method='C'
            final=record('方式別単価承認入力',key,dict(agency=agency,upper_id=u.id))
            manual=number(final.get('manual_amount'))
            if method=='C' and final.get('manual_amount') not in (None,''):
                if manual is None or manual<0:reason.append('個別単価の形式不正')
                elif not errors and equivalent:price=manual;formula='個別設定'
                else:reason.append('個別設定にも対応・適用根拠の確認が必要')
            # For genuinely new C rows, approve applicability explicitly; no October identity exists.
            if method=='C' and manual is not None and manual>=0 and equivalent and not (compact(agency).casefold()=='keioh株式会社' and u.family=='base' and u.item=='ftth_open'):
                if u.state=='VALUE' and number(u.amount) is not None and u.tax=='税別' and not u.hold and 'unresolved' not in u.family and u.family not in ('deduction','penalty') and active(u.period,target_period):
                    price=manual;formula='適用根拠確認済みの個別設定';reason=['個別設定の適用根拠・単価の最終確認']
            fingerprint=stable_id('PRICE',dict(version=1,snapshot=snapshot,agency=agency,
                upper=asdict(u),old=asdict(p.old) if p else None,previous=asdict(p.previous) if p and p.previous else None,
                common_identity={k:v for k,v in common_eq.items() if k!='record_hash'},
                policy={k:v for k,v in policy.items() if k!='record_hash'},identity={k:v for k,v in eq.items() if k!='record_hash'},
                method=method,price=str(price),manual=final.get('manual_amount')))
            old_candidate=final.get('candidate_hash')
            if final.get('approval_status') not in ('APPROVED','REJECTED'):final['candidate_hash']=fingerprint
            final['record_hash']=input_hash('方式別単価承認入力',final)
            policy_valid=not approval_errors(policy_name,policy,snapshot)
            state='未承認案' if price is not None else '保留'
            approval_problems=[]
            if final.get('approval_status')=='REJECTED':state='REJECTED'
            elif final.get('approval_status')=='APPROVED':
                approval_problems=approval_errors('方式別単価承認入力',final,snapshot)
                if old_candidate!=fingerprint:approval_problems.append('単価案改変・再承認が必要')
                if not policy_valid:approval_problems.append('基本方針/例外が未承認')
                if price is None:approval_problems.append('提案不可')
                state='APPROVAL_INVALID' if approval_problems else 'APPROVED'
            counts[method]+=1;counts['possible']+=price is not None;counts['approved']+=state=='APPROVED'
            counts['structural']+=not errors
            # One practical pricing decision per canonical agency/upper row, not per blocker.
            if state!='APPROVED':
                counts['remaining']+=1
                tasks.append([agency,u.id,method,' / '.join(reason) if price is None else '方式・単価案を確認し承認',state])
            identities.append([eq.get(h) for h in schemas['単価同等性入力'][0]])
            approvals.append([final.get(h) for h in schemas['方式別単価承認入力'][0]])
            output.append([agency,u.id,u.product,u.family,u.item,requested,method,expected_old,
                p.old.amount if p else None,p.previous.amount if p and p.previous else None,u.amount,delta,
                price,formula,'粗利保証なし（原価・追加報酬未確認）',state,' / '.join(reason+approval_problems),fingerprint])
            audit.append(dict(agency=agency,upper_id=u.id,method=method,amount=str(price) if price is not None else None,
                              candidate_hash=fingerprint,status=state,policy=policy,identity=eq,common_identity=common_eq,approval=final))
        summary.append([agency,policies[agency]['method'],counts['A'],counts['B'],counts['C'],counts['possible'],counts['remaining'],counts['approved'],counts['structural']])
    for name,records in [('単価共通同等性入力',common_records.values()),('単価基本方針入力',policies.values()),('成果項目例外入力',exceptions.values())]:
        write_sheet(book,name,schemas[name][0],[[r.get(h) for h in schemas[name][0]] for r in records],True)
    write_sheet(book,'単価同等性入力',schemas['単価同等性入力'][0],identities,True)
    write_sheet(book,'方式別単価承認入力',schemas['方式別単価承認入力'][0],approvals,True)
    write_sheet(book,'方式別11月単価案',['代理店','11月上位ID','商材','区分','成果項目','設定方式','適用方式','10月条件ID',
        '10月代理店単価','10月上位単価','11月上位単価','上位増減額','提案単価','算式','粗利表示','承認状態','保留・確認理由','単価案ハッシュ'],output)
    write_sheet(book,'方式別代理店集計',['代理店','基本方式','A件数','B件数','C件数','提案可能件数','残る実務判断件数','承認済み件数','同等性承認待ちを含む構造照合可能件数'],summary)
    write_sheet(book,'方式別残実務判断',['代理店','11月上位ID','方式','判断内容','状態'],tasks)
    write_sheet(book,'単価方式の使い方',['項目','内容'],[
        ['設定','単価基本方針入力で代理店別A/B、成果項目例外入力で商材・区分・成果項目別A/B/C。空欄の例外は基本方針を使用。初期Aは未承認の仮設定。'],
        ['A/B条件','条件軸・期間・単位・税・金額・構造が一意かつ明確であり、単価同等性入力に支払・除外・戻入・個別約定等の同等性を根拠付き承認した場合のみ。Aは旧単価維持、Bは上位増減額反映。'],
        ['共通確認','単価共通同等性入力で対象代理店・旧条件IDを固定した範囲に一括承認可能。個別の構造・税・単位・保留検査は省略しない。個別入力とのいずれか有効な確認を利用。'],
        ['C','新規・適用変更・不明はC。方式別単価承認入力のmanual_amountに個別単価。適用関係・構造・支払禁止が未解消なら金額を出さない。'],
        ['承認','既存と同じ再計算→record_hash確認→approved_hash転記と証跡入力→再計算。方式別単価承認入力はcandidate_hashにも結合。'],
        ['件数','11月上位条件×代理店が母集団。A+B+C=全行、提案可能は単価案あり、残判断はAPPROVED以外の行を1件。旧10月のみ終了判断・共通論点は別シート。'],
        ['粗利','単価案は粗利保証なし。既存の粗利計算・最終承認・Phase 3接続には自動昇格しない。既存承認入力は維持。']])
    stats=dict(rows=len(output),A=sum(x[2] for x in summary),B=sum(x[3] for x in summary),C=sum(x[4] for x in summary),
               structural_matches=sum(x[8] for x in summary),possible=sum(x[5] for x in summary),remaining=sum(x[6] for x in summary),approved=sum(x[7] for x in summary))
    return stats,audit
