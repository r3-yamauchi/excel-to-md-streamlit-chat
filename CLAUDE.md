# CLAUDE.md

このファイルは、Claude Code（claude.ai/code）がこのリポジトリのコードを操作する際のガイダンスを提供します。

## プロジェクト概要

ExcelファイルをMarkdown形式に変換するStreamlitチャットアプリケーション。XLSXファイルをアップロードし、選択したシートをMarkdown形式に変換、自動プレビュー表示、ダウンロードが可能。

## 開発環境のセットアップと実行

### uvを使用した環境構築（推奨）

```bash
# 依存関係のインストール
uv sync

# アプリケーションの実行
uv run streamlit run app.py

# デバッグモードでの実行
uv run streamlit run app.py --logger.level=debug

# アップロードサイズ制限の変更（デフォルト: 200MB）
uv run streamlit run app.py --server.maxUploadSize=500

# ポート変更
uv run streamlit run app.py --server.port=8502
```

### pipを使用した環境構築（代替方法）

```bash
# 仮想環境の作成
python -m venv .venv
source .venv/bin/activate  # Windowsの場合: .venv\Scripts\activate

# 依存関係のインストール
pip install -r requirements.txt

# アプリケーションの実行
streamlit run app.py
```

## 主要な依存関係とその役割

- **streamlit**: チャットインターフェースとWebアプリケーションフレームワーク
- **pandas (2.3.0)**: Excelファイルの読み込みとDataFrame処理
- **openpyxl (3.1.5)**: XLSXファイルの読み込みエンジン（pandas.read_excelのバックエンド）
- **tabulate (0.9.0)**: DataFrameをMarkdownテーブルに変換（pandas.to_markdown()のバックエンド）

## アプリケーションのアーキテクチャ

### コア実装ファイル: app.py

#### セッション状態管理 (app.py:16-28)

- `messages`: チャット履歴の保存（list of dict）
- `converted_markdown`: 変換結果の全体テキスト（string）
- `filename`: 出力ファイル名（string）
- `conversion_errors`: エラーメッセージのリスト（list）
- `markdown_results`: シートごとの変換結果を格納する辞書（dict[sheet_name: markdown_content]）
- `converted_sheets`: 正常に変換されたシート名のリスト（list）

#### ファイル検証フロー (app.py:39-55)

1. 拡張子チェック（.xlsx必須）
2. MIMEタイプ検証（3種類の有効なタイプを許可）
   - `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
   - `application/vnd.ms-excel.sheet.macroEnabled.12`
   - `application/octet-stream`（一部ブラウザでの互換性対応）
3. pandas.ExcelFileでの読み込み可能性チェック

#### 変換処理フロー (app.py:94-190)

1. 選択されたシートをループ処理
2. 各シートごとに個別のtry-exceptブロックで処理
3. DataFrameの前処理（object型を文字列に変換、NaN処理）
4. to_markdown()による変換、失敗時は代替フォーマットを使用
5. 結果をmarkdown_results辞書に保存
6. 変換完了後、自動的にプレビューを表示（app.py:177-188の新機能）

#### チャットコマンド処理 (app.py:210-245)

- `/download` または "ダウンロード": ダウンロードボタンを表示
- `/preview` または "プレビュー": 変換結果のプレビュー表示
- `/error` または "エラー": エラー詳細の表示
- 大文字小文字を区別しない処理（.lower()使用）

### 最新の実装変更点

#### 自動プレビュー機能 (app.py:177-188)

変換成功時に自動的に以下の処理を実行：
1. 各シートの変換結果を1000文字まで表示
2. Markdownコードブロックで整形表示
3. ダウンロード方法の案内を追加

```python
preview_content = "📋 **変換結果のプレビュー**\n\n"
for sheet_name in converted_sheets:
    sheet_content = markdown_results[sheet_name]
    preview_content += f"### 📄 シート: {sheet_name}\n\n"
    preview_content += f"```markdown\n{sheet_content[:1000]}{'...' if len(sheet_content) > 1000 else ''}\n```\n\n"
```

#### スラッシュコマンド対応 (app.py:210,217,224)

従来の日本語コマンドに加えて、スラッシュコマンドをサポート：
- `"/download" in prompt.lower()` で大文字小文字を区別しない
- 日本語コマンドとの OR 条件で両方に対応

### エラーハンドリング戦略

#### 1. ファイルレベルのエラー (app.py:195-201)

- `pd.errors.ParserError`: 破損ファイルや非Excel形式
- `PermissionError`: ファイルアクセス権限エラー
- `Exception`: その他の予期しないエラー

#### 2. シートレベルのエラー (app.py:143-147)

- `MemoryError`: 大きなシートの処理時
- 個別シートの読み込みエラー
- 空のDataFrameの処理（app.py:110-112）

#### 3. 変換レベルのエラー (app.py:126-139)

- to_markdown()失敗時の代替Markdown生成
- 警告メッセージとして記録し、処理は継続

### パフォーマンス考慮事項

- シート情報表示時は最初の5行のみプレビュー (app.py:72)
- プレビュー表示は最初の1000文字に制限（app.py:185, app.py:227-230）
- プログレスバーによる進捗表示 (app.py:149)
- シートごとの個別エラー処理により、一部失敗でも他のシートは変換可能

## 実装上の制約と注意事項

### Streamlitの制約

- デフォルトのアップロードサイズ制限: 200MB
- セッション状態はタブごとに独立
- リロード時にセッション状態がリセットされる
- file_uploaderは同時に1つのファイルのみ対応

### Excel処理の制約

- 対応形式: XLSX（Excel 2007以降）のみ
- 複雑な書式（結合セル、画像、グラフ）は無視される
- パスワード保護されたファイルは読み込み不可
- 非常に大きなファイルはメモリエラーの可能性

### Markdown変換の制約

- pandas.to_markdown()のmaxcolwidthsパラメータで列幅制御
- 特殊文字を含むセルは適切にエスケープされない場合がある
- 代替フォーマット使用時は基本的なテーブル形式のみ

## よくある拡張要求と実装方針

### CSV/XLS形式への対応

- file_uploaderのtypeパラメータに追加
- pd.read_csv()やxlrd/xlwtライブラリの追加が必要

### バッチ変換機能

- 複数ファイルのアップロードを許可
- ファイルごとにセッション状態を管理する辞書構造に変更

### 書式保持機能

- openpyxlの直接使用でセルの書式情報を取得
- Markdown拡張（色、太字等）での表現を検討

### エクスポート形式の追加

- HTMLエクスポート: pandas.to_html()の活用
- PDFエクスポート: markdownからPDFへの変換ライブラリ（例：markdown2pdf）の追加

## 開発時の注意点

### コミット時の確認事項

- セッション状態の初期化漏れがないか確認
- エラーハンドリングが適切に実装されているか確認
- 新しい依存関係を追加した場合は requirements.txt を更新
- MITライセンスとの互換性を確認

### テスト時の確認項目

- 空のExcelファイル
- 1シートのみのファイル
- 100シート以上の大量シートファイル
- 特殊文字を含むシート名
- 数式を含むセル
- 日付・時刻形式のセル
- 非常に長いテキストを含むセル

### コード品質基準

- 適切なエラーハンドリングとユーザーへのフィードバック
- 日本語と英語の混在を避ける（UIメッセージは日本語統一）
- 型ヒントの積極的な使用
- docstringによる関数の説明（日本語可）

## 今後の改善提案

### パフォーマンス最適化

- 大規模ファイルの分割処理
- 非同期処理の導入でUIのレスポンス改善
- キャッシュ機能の実装（@st.cacheデコレータの活用）

### ユーザビリティ向上

- 変換設定のプリセット機能
- 変換履歴の保存と再利用
- プレビュー表示のカスタマイズオプション

### セキュリティ強化

- ファイルサイズによる制限の動的調整
- アップロードファイルのウイルススキャン連携
- ユーザーセッションの適切な管理

---

最終更新: 2025年6月13日