# auひかり条件表更新・配信管理システム

P0のデータモデルと、P1の原本読み取り・項目候補抽出・上位条件比較の基盤です。P1は未承認IDや業務判断が必要な候補を保留し、代理店向け帳票生成や配信は行いません。

## 構成

- `01_上位原本/`: 上位原本の保管
- `02_取得データ/`: 取得データの保管（仮置き）
- `03_条件表/`: 条件表の保管（仮置き）
- `04_成果物/`: 生成成果物の保管（仮置き）
- `05_配信管理/`: 配信管理データの保管（仮置き）
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

P1解析結果はJSONで標準出力されます。現在の入力資料に対する完了ゲートが未達の場合、このコマンドは終了コード1を返します。

## P0/P1モデルと解析

`src/au_hikari_conditions/models.py` に `SourceDocument`、`ConditionItem`、`Agency`、`DeliverableItemDefinition` と `DELIVERABLE_ITEM_ID_DICTIONARY` を定義しています。`document_reader.py` はPDFページ、Excelシート/セル、Word段落/表セルの根拠位置を保持します。`normalization.py` と `comparison.py` は候補抽出・正規化と保留付き差分を行います。金額確定、代理店報酬計算、帳票生成、メール送信は対象外です。
