# Excel to Markdown Converter

Excel（XLSX）ファイルをMarkdown形式に変換するStreamlitベースのチャットアプリケーション

## 🎯 主な機能

- **直感的なチャットインターフェース**: ファイルをアップロードして対話的に変換
- **複数シート対応**: 変換したいシートを選択可能
- **自動プレビュー**: 変換完了後、結果を自動的にプレビュー表示
- **エラーハンドリング**: シートごとのエラー処理により、一部失敗でも他のシートは変換可能
- **大容量ファイル対応**: デフォルト200MBまで（設定で変更可能）

## 🚀 クイックスタート

### uvを使用する場合（推奨）

```bash
# リポジトリのクローン
git clone https://github.com/yourusername/excel-to-md-streamlit-chat.git
cd excel-to-md-streamlit-chat

# 依存関係のインストールと実行
uv sync
uv run streamlit run app.py
```

### pipを使用する場合

```bash
# リポジトリのクローン
git clone https://github.com/yourusername/excel-to-md-streamlit-chat.git
cd excel-to-md-streamlit-chat

# 仮想環境の作成
python -m venv .venv
source .venv/bin/activate  # Windowsの場合: .venv\Scripts\activate

# 依存関係のインストール
pip install -r requirements.txt

# アプリケーションの実行
streamlit run app.py
```

## 📝 使い方

1. **ファイルのアップロード**
   - XLSXファイルをドラッグ&ドロップまたは「Browse files」からアップロード

2. **シートの選択**
   - アップロード後、変換したいシートをチェックボックスで選択

3. **変換の実行**
   - 「選択したシートを変換」ボタンをクリック
   - 変換結果が自動的にプレビュー表示されます

4. **結果のダウンロード**
   - チャットで「ダウンロード」または `/download` と入力
   - Markdownファイルをダウンロード可能

## 💬 チャットコマンド

| コマンド | 説明 |
|---------|------|
| `ダウンロード` または `/download` | 変換結果をダウンロード |
| `プレビュー` または `/preview` | 変換結果をプレビュー表示 |
| `エラー` または `/error` | エラー詳細を表示 |

## ⚙️ 設定オプション

### アップロードサイズの変更

```bash
# 500MBに変更する場合
uv run streamlit run app.py --server.maxUploadSize=500
```

### ポート番号の変更

```bash
# ポート8502で起動
uv run streamlit run app.py --server.port=8502
```

### デバッグモード

```bash
# 詳細なログ出力を有効化
uv run streamlit run app.py --logger.level=debug
```

## 🛠️ 技術スタック

- **[Streamlit](https://streamlit.io/)** - Webアプリケーションフレームワーク
- **[pandas](https://pandas.pydata.org/)** - Excelファイルの読み込みとデータ処理
- **[openpyxl](https://openpyxl.readthedocs.io/)** - XLSXファイルの読み込みエンジン
- **[tabulate](https://github.com/astanin/python-tabulate)** - Markdownテーブル生成

## 📋 システム要件

- Python 3.12以上
- 200MB以上の空きメモリ（大きなファイルの場合はそれ以上）

## 🚧 制限事項

- **対応形式**: Excel 2007以降（.xlsx）のみ対応
- **書式**: セルの色、フォント、結合セルなどの書式は保持されません
- **オブジェクト**: グラフ、画像、図形は変換されません
- **保護**: パスワード保護されたファイルは読み込めません

## 📄 ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は[LICENSE](LICENSE)ファイルを参照してください。

## 🙏 謝辞

このプロジェクトは以下のオープンソースプロジェクトを使用しています：

- Streamlit
- pandas
- openpyxl
- tabulate
