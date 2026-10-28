# auひかり条件表更新・配信管理システム

P0フェーズの初期セットアップです。現段階では標準ライブラリのみを使用し、データモデルの骨組みと環境確認を提供します。

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

02〜06のフォルダー名とモデル項目は、仕様書の詳細を確認後に確定してください。成果項目ID辞書は実IDを推測せず、空の登録先として定義しています。

## 環境確認

Python 3.10以降を用意し、リポジトリルートで実行します。

```powershell
python scripts/check_environment.py
python -m unittest discover -s tests
```

パッケージとしてインストールする場合:

```powershell
python -m pip install -e .
```

## P0モデル

`src/au_hikari_conditions/models.py` に `SourceDocument`、`ConditionItem`、`Agency`、`DeliverableItemDefinition` と `DELIVERABLE_ITEM_ID_DICTIONARY` を定義しています。仕様確定前の骨組みのため、業務ルールや永続化処理は含みません。
