> 最新運用：[完成テンプレートによる月次作成・LOCK・PDF・Zoho下書き](docs/monthly-template-workflow.md)。11月は10社分の条件表を承認・LOCKし、PDF2件ずつのZoho API下書き作成まで実施済み。送信は担当者が手動で行います。

# auひかり条件表更新・配信管理システム

P0のデータモデル、P1の原本読み取り・項目候補抽出、P2の承認済み条件計算、P3の帳票生成・読み戻し検証、P4-AのZoho Mail OAuth基盤、P4-Bの下書き準備、およびP5の月次統合フローを備えます。P5では承認前候補と承認後の配布版を別工程・別場所で管理します。完成テンプレートとLOCK済みPDFを使う現行の下書き経路は実APIで検証済みです。旧P5の再生成経路とは区別します。自動送信機能はありません。

## 構成

- `01_上位原本/`: 上位原本の保管
- `02_取得データ/`: 承認済み粗利・原価・条件対応・例外・承認管理の入力場所。抽出候補は実データを含むためGit外で管理
- `03_条件表/`: 条件表の保管（仮置き）
- `04_成果物/`: 生成成果物の保管（仮置き）
- `05_配信管理/`: 承認済み宛先・本文・月次承認・送信履歴の管理スキーマ。実運用データはGit外で管理
- `06_作業記録/`: 作業記録の保管（仮置き）
- `07_監査ログ/`: 監査ログの保管
- `src/au_hikari_conditions/`: Pythonパッケージとデータモデル
- `scripts/`: 開発・運用補助スクリプト
- `tests/`: スモークテスト
- `docs/`: 設計・運用ドキュメント

成果項目ID辞書のP1項目IDは暫定です。正式ID、正規化規則、10月原本の優先関係は設計書承認後に確定してください。

## 環境確認

Python 3.10以降を用意し、リポジトリルートで実行します。

```powershell
py -m pip install -e .
py scripts/check_environment.py
py -m unittest discover -s tests
```

パッケージとしてインストールする場合:

```powershell
py scripts/run_p1_analysis.py
```

P1解析結果はJSONで標準出力されます。P1ゲートは文書読取・候補抽出・根拠参照・上位条件の候補差分だけを判定し、P2以降に必要な承認・代理店条件との照合は`review_holds`に分けて報告します。P1ゲート未達の場合、このコマンドは終了コード1を返します。P1ゲート達成は条件確定・帳票生成・発送の許可を意味しません。

承認入力テンプレート、抽出候補ブック、架空配信サンプルの生成:

```powershell
py scripts/prepare_input_workbooks.py
```

10月代理店資料と11月上位原本からの候補ブックには実データが含まれるため、`.gitignore`対象です。会社承認済みのGit外保管先へ移し、候補の承認前に支払額を確定しないでください。このコマンドは既存ファイルを上書きしません。配信管理ブックの送付先・本文例は架空の`example.invalid`であり、実送信やZoho接続は行いません。

### 月次P1比較からP2入力への引継ぎ

会社が承認したローカル環境で、前月上位原本・前月代理店帳票・当月上位原本の3種類を比較候補ブックにまとめます。出力には根拠抜粋とファイルハッシュが含まれるため、入力資料と同じ会社承認済みGit外領域へ保存してください。既存ファイルは上書きしません。

```powershell
py scripts/prepare_monthly_comparison.py `
  --previous-period 2026-10 `
  --target-period 2026-11 `
  --select-agency-interactively `
  --output "C:\company-approved\2026-11\monthly-comparison.xlsx"
```

ファイル選択ダイアログで同一代理店のファイルを1つ選ぶと、ファイル名の種別ラベルと共通代理店キー候補から`auひかり`、`auひかりプラス`、`キャンペーン条件`を自動でグループ化します。各資料は別々に抽出・比較し、候補ブックに3種類を区別して記録します。自動グループ化は仮候補です。資料一覧と代理店資料種別シートで同一代理店・種別を必ず人が確認してください。欠損・種別重複・判定不能は保留となり、自動で統合しません。

選択したパスはコマンド引数や出力ログに表示されません。キャンペーン条件は対象商材へのリンク候補を保持しますが、通常報酬へ自動加算しません。資料間の条件キー重複・金額等の矛盾は`代理店資料間照合`に未承認候補として記録し、自動統合しません。

ブックの`抽出根拠候補`にP1の候補値・条件軸・根拠位置、`上位差分候補`に前月/当月上位原本の差分候補、`P2項目対応`にv4入力列との対応案と不足する社内承認データ、`不足・保留`に欠落・曖昧・重複を記録します。終了コード2でも比較ブックは作成済みの場合があります。JSONの状態とブック内の保留を確認し、人手レビューを継続してください。候補はすべて未承認です。

代理店名だけで条件を対応付けません。前月代理店帳票の`agency_id`、商品/プラン/ISP、住居・速度・申込区分・適用期間を担当者が確認し、条件軸がない・複数一致・読取不能なら保留します。粗利・直接原価・各税区分・前回確定報酬・代理店例外・掲載可否は上位原本だけでは確定しないため、承認済み社内根拠から別途入力が必要です。ラダーは現行P2入力スキーマに専用列がなく、ルールの設計・承認まで計算保留です。

この比較処理はP2の承認済み列へ値を自動転記しません。P2の金額行には代理店ID、正式項目ID、承認済み条件対応、税区分、根拠、ハッシュ付き承認が必要で、原本候補だけからそれらを一意に決められないためです。担当者は比較ブックを根拠索引にし、v4の`粗利原価`、`条件対応`、`代理店別例外`、`前回条件`へレビュー済みデータを転記し、`承認管理`に承認証跡を記録します。承認前の候補値を`APPROVED`へ変更してはなりません。

`prepare_monthly_comparison.py`はP1の現行抽出器を前月上位・前月代理店3種・当月上位の各資料へ適用し、上位原本間および種別ごとの条件候補を比較します。条件軸不足・同一キー重複・資料欠落・読取不能は候補ブックに保留として残します。P5の`prepare`/`finalize`には`--monthly-comparison`が必須で、対象期間・選択資料・元ファイルのSHA-256を再検証します。比較ブックのハッシュも月次承認Snapshotへ含め、prepare後の比較結果または元資料の変更を拒否します。

## P0/P1/P2モデルと解析

`src/au_hikari_conditions/models.py` に `SourceDocument`、`ConditionItem`、`Agency`、`DeliverableItemDefinition` と `DELIVERABLE_ITEM_ID_DICTIONARY` を定義しています。`document_reader.py` はPDFページ、Excelシート/セル、Word段落/表セルの根拠位置を保持します。`normalization.py` は成果項目候補に加えて、資料内で確認できる商材・プラス種別・ISP・住居種別・速度・申込／開通区分などの条件軸候補を抽出します。条件軸は候補の根拠位置とともに保留され、PDFページやExcelシートをまたいだ文脈補完は行いません。`comparison.py` は保留付き差分を行います。金額確定、代理店報酬計算、帳票生成、メール送信は対象外です。

`calculation.py` はP2の決定的な計算コア、`p2.py` は承認入力ブックとの接続です。入力シートの列・期間・状態・承認ID・承認レコードのSHA-256を検証し、承認済み前回報酬と前回／対象月の条件対応、粗利・原価・税区分・例外を既存計算コアへ渡します。条件差異、未承認・不明値、税抜でない金額、負額、最低粗利割れ、例外確認があれば計算結果を出さず、理由一覧へ保留します。計算結果は提案額と前回差分を含む内部用Excelです。既存のv3テンプレートは保持し、P2列（期間ルール、掲載可否、前回報酬）が追加されたv4を利用します。金額や承認は補完・推定しません。

承認済み入力ブックからP2結果を生成する場合は、入力・出力ともリポジトリ外の会社承認済み保管場所を指定します（既存出力は上書きしません）。

```powershell
py scripts/run_p2_calculation.py --input "C:\approved-store\input.xlsx" --target-period 2026-11 --output "C:\approved-store\p2-result.xlsx"
```

`record_hash`は承認対象レコードの業務列から算出したSHA-256です。承認者・承認日時はタイムゾーン付きISO-8601で記録してください。スクリプトは保留行があれば終了コード2で終了し、生成したブックに代理店別計算と保留理由を分けて出力します。P2の匿名化回帰テストでは10月条件の再現と11月更新を確認します。実資料・候補データは処理対象にせず、Zoho連携・外部送信は行いません。

### P3 代理店別帳票生成・読み戻し検証

`p3.py`はP2結果ブックを承認入力から再計算し、行数・代理店/成果項目・状態・税区分後の金額・前回差分・掲載可否を照合してから、代理店ごとのExcel/Word帳票を生成します。適用条件、代理店名、対象月、報酬、掲載可否、Excel表示形式/印刷範囲、Word表/書式を保存後に読み戻して検証します。表示可否が`WITHHOLD`の場合、金額は表示せず「非提示」とします。空欄の条件軸は非適用とみなさず、明示値または`NOT_APPLICABLE`等の入力が必要です。既存の10月資料は複雑な結合セル/複数セクション構成のため、原本を上書き・コピーせず、共通の構造化テンプレートとして再構成します。

生成物は必ずリポジトリ外の会社承認済み保管場所に置き、既存ファイルは上書きしません。P2に保留がある場合や、P2結果が入力からの再計算と一致しない場合は何も生成しません。生成結果の状態は常に内部確認用で、配信承認・最終QA済みを意味しません。

```powershell
py scripts/run_p3_generation.py --input "C:\approved-store\input.xlsx" --p2-result "C:\approved-store\p2-result.xlsx" --target-period 2026-11 --output-directory "C:\approved-store\p3-review"
```

この段階では10月32件との内容一致・目視レイアウト承認は未実施です。実データが承認済みの後、代理店別の差戻し確認と帳票目視QAを行ってください。

`workflow.py` は代理店ごとの改定状態遷移、成果物・本文・テンプレート・宛先・マスター版・対象期間を含む承認スナップショット照合、および送信冪等性キーの生成を行います。P4-Bの下書き処理履歴はSQLiteに保存しますが、代理店ID・宛先・本文・添付内容を記録しません。WorkDrive連携と実送信は未実装です。

### P4-B Zoho Mail 下書き準備（実機未実施）

`p4b.py` は、代理店別の宛先・件名・本文設定、期間と代理店情報を含む安全なプレビュー、承認済みWorkflow/Snapshot照合、P3 Excel/Wordの実ファイル・SHA-256照合、Zohoへの添付アップロードと `mode="draft"` 固定の下書き保存を提供します。TOが未設定、メールアドレス重複、テンプレート不正、承認やSnapshotの不一致、Excel/Wordの代理店・期間違い、内部確認/未承認マーカー、ファイル差替え、添付アップロード応答のファイル名不一致は通信前または下書き保存前に停止します。`save_draft()` は `mode` が `draft` 以外なら履歴登録・通信前に停止し、送信APIはありません。

下書きの一意キーは代理店・対象月のハッシュです。再実行を含む2回目はブロックします。添付アップロードの失敗は `FAILED`、下書きPOST後の通信/応答エラーは作成成否を判定できないため `OUTCOME_UNKNOWN` として再試行を禁止します。処理履歴DBはGitリポジトリ外の会社承認済み保管場所へ置き、実宛先・本文・代理店IDを保管しません。DBファイルとそのSQLite journal/WALファイルへのアクセス権を会社の秘密情報管理方針に従って制限してください。

**P5のP3最終化:** 内部確認用P3帳票を編集・警告削除して配布版へ変換することはありません。P5はP2から構造化条件を再取得し、別のGit外「承認候補」領域へ警告なしの候補を新規生成します。候補は承認されるまで配布・下書きに利用できません。月次承認表には、P2入力、Excel/Word候補、代理店別宛先、本文、テンプレートの各ハッシュを記録します。QA・承認後、P5は承認表と候補の現在の内容/ハッシュを再照合し、承認済み候補の同じバイト列を別のGit外配布先へ昇格します。内容を書き換えないため、昇格後も承認Snapshotのハッシュが一致します。代理店ID・社名・対象月・P2再計算内容・未承認表示の有無も読み戻し検証します。

代理店の実宛先・本文テンプレートはリポジトリに配置せず、会社承認済みGit外の配信管理ExcelからP5が読み込みます。`DraftHistoryStore` にはGit外のDBパスを渡します。実機使用前に会社・Zoho管理者から、Self Clientの `ZohoMail.messages.CREATE` 権限（添付APIにも必要な権限を含む）、対象アカウント、対象データの使用、下書き作成を明示承認してください。現在確認済みの `ZohoMail.accounts.READ` だけでは下書き作成できるとみなせません。

P5は承認済みのExcel・Snapshot・WorkflowからP4-Bプレビューを再構築します。設定値やプレビュー本文、宛先、API応答本文をコンソール・ログへ出力しないでください。P5はデフォルトでZoho APIを呼び出しません。実APIは外部承認フラグ付きの明示的な下書き実行時のみ呼ばれます。

P4-Bの自動テストは匿名ファイルとモックAPIだけを使います。実Zoho API、実トークン、実メールアドレス、実帳票のアップロード・下書き作成・送信は一切実行していません。

### P5 月次統合運用

`scripts/run_p5.py` の `prepare` は、前月比較ブック、リポジトリ内のP1ゲート、承認済みP2入力、P2計算、P3承認候補生成、承認済み配信管理Excelの宛先/テンプレートを連続検証し、月次承認行を追加した別Excelを作ります。前月比較ブックは`prepare_monthly_comparison.py`で生成したGit外ブックを`--monthly-comparison`に指定します。P1 blocker、資料ハッシュ不一致、P2保留、未承認/期限切れ宛先・テンプレート、本文/宛先不整合があれば停止します。P1の`review_holds`は業務承認済みを意味しません。P2の承認済み入力が別途必要です。

#### 毎月の担当者チェックリスト

1. **原本を配置（担当者）**: 前月の上位原本・代理店別条件表、当月上位原本を各`01_上位原本\YYYY-MM\`、`03_条件表\YYYY-MM\`へ置きます。P1が比較する上位キャンペーン通知書を各月1件に特定できるようにし、元ファイルは変更しません。
2. **前月比較を作成（自動）・確認（担当者）**: 上記`prepare_monthly_comparison.py`を実行し、出力ブックをGit外に保存します。抽出エラー・不足資料・複数一致・条件軸不明を確認し、代理店IDと条件対応を手動でレビューします。未解決項目は保留のままです。
3. **P2入力Excelを準備（担当者）**: [`承認済み業務データ入力テンプレート_v4.xlsx`](./02_取得データ/承認済み業務データ入力テンプレート_v4.xlsx)を会社承認済みのGit外作業領域へコピーし、比較ブックを根拠索引に、粗利・原価・税区分・条件対応・代理店別例外・前回条件を入力します。承認管理の状態、承認者、タイムゾーン付き承認日時、理由、レコードハッシュを整備します。不明値は推測せず、承認済み根拠がなければP2を停止します。
4. **配信管理Excelを準備（担当者）**: [`配信管理入力テンプレート_架空サンプル_v3.xlsx`](./05_配信管理/配信管理入力テンプレート_架空サンプル_v3.xlsx)の構造を参考に、会社承認済みGit外の配信管理Excelを準備します。代理店ID・社名・宛先・CC・送信元・宛名・件名/本文テンプレート・有効期間・承認状態を確認してください。架空サンプルの宛先・本文を実運用に流用しないでください。
5. **prepareを実行（自動）**: 以下の例の全パスをGit外の会社承認済み領域へ置き換えて実行します。`--monthly-comparison`には手順2の比較ブックを指定します。P2結果、承認依頼Excel、P3候補が新規作成されます。既存ファイルを指定しないでください。

```powershell
py scripts/run_p5.py prepare `
  --target-period 2026-11 `
  --batch-id batch-2026-11-run01 `
  --approved-input "C:\company-approved\2026-11\p2-input.xlsx" `
  --p2-result "C:\company-approved\2026-11\run01\p2-result.xlsx" `
  --monthly-comparison "C:\company-approved\2026-11\monthly-comparison.xlsx" `
  --delivery-workbook "C:\company-approved\distribution\delivery-master.xlsx" `
  --approval-workbook "C:\company-approved\2026-11\run01\monthly-approval.xlsx" `
  --candidate-directory "C:\company-approved\2026-11\run01\approval-candidates" `
  --requested-by operator-identifier
```

6. **候補をQA・承認（担当者／承認者）**: 各代理店のExcel/Word候補、P2結果、前月比較ブックと月次承認表を照合します。代理店、対象月、金額・適用条件、掲載内容、宛先、CC、送信元、件名/本文、Excel/Word両添付を確認します。問題がなければ対象行の`qa_status=PASSED`、`approval_status=APPROVED`、`approved_by`、タイムゾーン付きISO-8601の`approved_at`、`approval_reason`を入力します。ハッシュ列は変更しません。不一致・保留・不明があれば承認せず、修正後は新しい`batch-id`で最初から実行します。
7. **finalizeを実行（自動）**: P2入力・前月比較ブック・元資料ハッシュ・宛先/本文・候補のSnapshotを再照合し、すべて一致した場合だけ候補と同じ内容をGit外の最終配布先へ昇格します。通常の配布準備では`--save-drafts`を付けず、最終帳票を人が再確認します。
8. **下書き作成（別途会社承認後のみ）**: Zohoの必要スコープ・対象アカウント・データ利用の会社承認が確認済みの場合に限り、承認された実行環境で`--save-drafts`を付けます。下書き保存後、人がZoho画面で宛先・本文・添付を確認し、メール送信は社内手順に沿って人が行います。P5に送信機能はありません。

**匿名E2Eの再実行**: `py -m unittest tests.test_p5_integration.P5IntegrationTest.test_cli_prepare_human_approval_finalize_with_mock_draft` は一時領域に作った架空の前月原本・代理店条件・当月原本から比較CLIを実行し、P2入力、P5の`prepare`→模擬人手承認→`finalize --save-drafts`までを確認します。P1ゲートはテスト内で合格状態を注入し、Zohoクライアントはモックへ置換するため、実原本解析、実トークン、実API、実データ、メール送信は行いません。実運用の原本解析・承認判断を代替するテストではありません。

```powershell
py scripts/run_p5.py prepare `
  --target-period 2026-11 `
  --batch-id batch-2026-11 `
  --approved-input "C:\company-approved\2026-11\p2-input.xlsx" `
  --p2-result "C:\company-approved\2026-11\p2-result.xlsx" `
  --monthly-comparison "C:\company-approved\2026-11\monthly-comparison.xlsx" `
  --delivery-workbook "C:\company-approved\distribution\delivery-master.xlsx" `
  --approval-workbook "C:\company-approved\2026-11\monthly-approval.xlsx" `
  --candidate-directory "C:\company-approved\2026-11\approval-candidates" `
  --requested-by operator-identifier
```

候補と月次承認ブックを会社承認済みアクセス制御下で目視QAしてください。月次承認シートには、正しい代理店・期間・候補ハッシュ・宛先/本文/テンプレート/条件マスターの各ハッシュが記録されています。QA完了後、該当行の`qa_status`を`PASSED`、`approval_status`を`APPROVED`にし、承認者ID・タイムゾーン付きISO-8601承認日時・承認理由を入力します。ハッシュ列は変更しないでください。`送付先マスター`と`本文テンプレート`は事前承認済み・対象月全期間有効である必要があります。サンプル行を実運用に流用せず、サンプル本文テンプレートを実運用ブックに残さないでください。テンプレート`template_hash`は公開関数`compute_mail_template_hash(template_row)`で算出し、本文/件名/許可プレースホルダー等を変更した場合は再承認してください。`reply_to_address`は現在のZoho下書きAPIペイロードで未対応のため、値がある場合は停止します。

承認後の`finalize`はP1/P2を再検証し、全対象代理店のQA/承認、宛先/本文/テンプレート、P2入力、候補ファイルの内容とハッシュが一致する場合のみ配布版を昇格します。`--save-drafts`を付けない場合はファイル昇格だけで、Zohoへ接続しません。

```powershell
py scripts/run_p5.py finalize `
  --target-period 2026-11 `
  --batch-id batch-2026-11 `
  --approved-input "C:\company-approved\2026-11\p2-input.xlsx" `
  --p2-result "C:\company-approved\2026-11\p2-result.xlsx" `
  --monthly-comparison "C:\company-approved\2026-11\monthly-comparison.xlsx" `
  --delivery-workbook "C:\company-approved\distribution\delivery-master.xlsx" `
  --approval-workbook "C:\company-approved\2026-11\monthly-approval.xlsx" `
  --candidate-directory "C:\company-approved\2026-11\approval-candidates" `
  --final-directory "C:\company-approved\2026-11\approved-distribution"
```

実Zohoへの下書き作成は別途会社の明示承認、Self Clientの`ZohoMail.messages.CREATE`等必要スコープの承認、運用テスト承認後のみ実行できます。未承認時は`--save-drafts`を使わないでください。承認を得た後に限り、会社管理の実行環境へ`P5_ZOHO_DRAFTS_APPROVED=YES`、`P5_DRAFT_HISTORY_DB`（絶対パスのGit外SQLiteファイル）を設定して、上の`finalize`に`--save-drafts`を加えます。OAuth認証は[DPAPI token file configuration](#dpapi-token-file-selection)を参照します。スクリプトは結果件数と状態だけを表示し、メール本文・宛先・添付内容・認証情報を出力しません。メール送信APIは存在せず、人間がZoho画面で下書きを確認し、別の社内手順で手動送信します。

再実行は既存成果物を上書きしません。入力・宛先・テンプレート・帳票の修正後は、監査用に旧成果物を保持し、新しい`batch_id`と新規のP2結果/候補/承認ブック/出力先を使って`prepare`からやり直してください。同一代理店・対象月のZoho下書きはSQLite履歴で二重作成を拒否します。下書きPOST後に結果が不明な場合は再実行せず、Zoho上の状態を人間が確認してください。保留または承認取消時は`finalize`を実行せず、内容・宛先・添付を修正して再QA/再承認します。

P5自動テストは合成P2入力、架空宛先、候補帳票、Zohoモックを使い、P2→P3候補→承認Excel→最終化→プレビュー→モック下書き保存を検証します。実API・実トークン・実データでは動作させていません。

### P4-A Zoho Mail 認証・読み取り基盤

`zoho_mail.py` は会社管理のSelf Client OAuthクライアントID・クライアントシークレット・refresh tokenを使ってアクセストークンを更新し、Zoho Mailのアカウント一覧を読み取ります。`ZOHO_ACCOUNT_ID`を外部設定から読み、取得可能なアカウント一覧内にそのIDがあることを確認するAPIを備えます。データセンターは `ZOHO_REGION` で明示し、未対応値はURLを推測せず拒否します。アクセストークンはメモリ内でのみキャッシュし、有効期限前に更新します。Mail APIが401を返した場合は1回だけ更新後に再試行します。認証・権限・リージョン・レート制限・通信・応答形式のエラーは分類して通知し、レスポンス本文や認証情報を例外メッセージへ含めません。

標準の環境変数プロバイダーは以下の名前を読み取ります。`DpapiSecretProvider` はWindows PowerShellの `ConvertFrom-SecureString` が生成するCurrentUser DPAPIの16進形式を現在のWindowsユーザーの `CryptUnprotectData` でプロセス内復号します。Self Client JSONはUTF-8またはUTF-16のJSONオブジェクトで、キーは `client_id` と `client_secret` です。refresh tokenは別のDPAPIファイルから取得し、`account_id`・`region`は別の `SecretProvider`（既定では環境変数）へフォールバックします。Windows PowerShell 5.1の `Set-Content` は既定でBOMなしのシステムANSIテキストとCRLFを使いますが、DPAPI blobはASCIIの16進表現であるためANSIコードページでも文字を損ないません。Python側の読込はUTF-8/UTF-16 BOM付きテキストも許容します。指定された保存コマンドとファイル名を根拠にした方式であり、他のDPAPI形式、entropy付きデータ、別Windowsユーザー／別端末で保護されたデータには対応しません。秘密情報ファイルは読み取り専用で扱い、書換えません。別の秘密情報管理手段を使う場合も `SecretProvider.get_secret(name)` を実装して `ZohoMailConfig.from_secret_provider(...)` に渡せます。

- `ZOHO_CLIENT_ID`
- `ZOHO_CLIENT_SECRET`
- `ZOHO_REFRESH_TOKEN`
- `ZOHO_ACCOUNT_ID` — 送信元候補として確認したアカウントの `accountId`。アドレスから推測しません。
- `ZOHO_REGION` — `us`, `eu`, `in`, `au`, `cn`, `jp`, `ae`, `ca`, `sa`, `uk` のいずれか。契約中のZohoデータセンターを会社管理者に確認して設定します。

Mail APIのアカウント読み取りには `ZohoMail.accounts.READ` スコープが必要です。API利用可否とスコープはSelf Clientの認可時に確認してください。読み取り接続はアカウント取得のみを行います。P4-Bは下書き作成APIを実装していますが、P4-Aの読み取り権限だけでは実行できるとみなさず、別途 `ZohoMail.messages.CREATE` 等の承認・権限付与後に限ります。テストは合成応答を使い、実APIへ接続しません。

Windowsの検証端末では `DpapiSecretProvider()` を `ZohoMailConfig.from_secret_provider(...)` に渡して使用できます。既定では `%LOCALAPPDATA%\ZohoMailApiTest\refresh_token.dpapi` と `%LOCALAPPDATA%\ZohoMailApiTest\self_client.json` を読み込みます。Self Client JSONの期待スキーマは `{"client_id":"...","client_secret":"..."}` です。`ZOHO_ACCOUNT_ID` と `ZOHO_REGION` は会社承認済みの環境設定等から別途提供してください。DPAPI復号値はPythonプロセス内メモリで扱い、子プロセス、標準出力、コマンドライン、ログへ渡しません。

#### DPAPI token file selection

既存の`refresh_token.dpapi`は変更せず、別のDPAPI blobを使うときだけ`ZOHO_REFRESH_TOKEN_DPAPI_PATH`へGit外にある絶対パスを指定します。例:`C:\company-secrets\ZohoMailApiTest\refresh_token_v2.dpapi`。設定があればそのファイルだけを選び、ファイル欠損や復号失敗時に旧ファイルへ自動フォールバックしません。未設定なら従来の`%LOCALAPPDATA%\ZohoMailApiTest\refresh_token.dpapi`を使います。P5ではこの環境変数を設定するだけでV2ファイルを選択でき、既存ファイルの移動・変更・再発行は不要です。

自動テストは実ファイルを開かず、一時ディレクトリに合成Self Client JSONと、Windows PowerShell標準の `ConvertFrom-SecureString | Set-Content` で生成したダミーDPAPI blobのみを作成します。テストは実行時のBOM有無・ASCII互換文字コード・CRLF改行を検証し、Pythonプロセス内で復号します。テスト内のPowerShell起動にはダミー値のみを標準入力で渡し、実秘密情報は扱いません。Windows以外、またはWindows PowerShellが利用できない環境ではDPAPIラウンドトリップテストはスキップされます。

#### 読み取り専用の実機受入手順（未実行）

以下は会社の明示承認を得た後に、指定された検証端末・Zohoアカウントでのみ実行できます。この作業では実APIへ接続していません。実行前に、Zoho管理者がSelf Client、`ZohoMail.accounts.READ`、リージョン（手動確認済みのUSは `us`）、送信元候補の `accountId` を承認済み設定保管先で確認してください。実IDやトークンをGit・作業ログ・コマンド履歴へ記録しないでください。

1. 会社承認済みの実行環境から、`ZOHO_ACCOUNT_ID` と `ZOHO_REGION=us` をプロセスへ注入します。refresh tokenとSelf Client情報は既定のDPAPI/JSONファイルから読み取ります。PowerShell履歴、シェル起動引数、共有ログ、`.env`への秘密値の貼り付けは避けます。
2. 承認済みのPython実行環境で、出力・ログを追加せず、読み取り専用の処理だけを実行します。

   ```python
   from au_hikari_conditions.zoho_mail import DpapiSecretProvider, ZohoMailClient, ZohoMailConfig

   config = ZohoMailConfig.from_secret_provider(DpapiSecretProvider())
   ZohoMailClient(config).get_configured_account()
   ```

   返されたアカウント情報は表示・記録しません。アカウント一致／不一致の結果だけを、承認済み監査先へ秘密値を含めず記録します。
3. 401/403/404等では設定を確認して停止し、トークン値や応答本文を調査ログに複写しません。
4. 検証後はプロセスを終了して環境変数を破棄します。下書き作成、送信、実帳票生成、認証情報の再発行はこの受入手順に含みません。
