import streamlit as st
import pandas as pd
import io
from datetime import datetime
import mimetypes
import os

# ページ設定
st.set_page_config(
    page_title="Excel to Markdown Converter",
    page_icon="📊",
    layout="wide"
)

# セッション状態の初期化
if "messages" not in st.session_state:
    st.session_state.messages = []
if "converted_markdown" not in st.session_state:
    st.session_state.converted_markdown = None
if "filename" not in st.session_state:
    st.session_state.filename = None
if "conversion_errors" not in st.session_state:
    st.session_state.conversion_errors = []
if "markdown_results" not in st.session_state:
    st.session_state.markdown_results = {}
if "converted_sheets" not in st.session_state:
    st.session_state.converted_sheets = []

# タイトル
st.title("📊 Excel to Markdown Converter Chat")
st.markdown("ExcelファイルをMarkdown形式に変換してダウンロードできます")

# サイドバー
with st.sidebar:
    st.header("📁 ファイルアップロード")
    uploaded_file = st.file_uploader(
        "XLSXファイルを選択してください",
        type=['xlsx'],
        help="Excel形式のファイル（.xlsx）をアップロードしてください"
    )
    
    if uploaded_file is not None:
        # ファイル形式の検証
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()
        mime_type = uploaded_file.type
        
        # MIMEタイプの検証
        valid_mime_types = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel.sheet.macroEnabled.12',
            'application/octet-stream'  # 一部のブラウザではこのMIMEタイプになることがある
        ]
        
        if file_extension != '.xlsx':
            st.error(f"❌ エラー: アップロードされたファイル '{uploaded_file.name}' はXLSX形式ではありません。")
            st.info("📝 サポートされているファイル形式: .xlsx")
        elif mime_type not in valid_mime_types and mime_type != 'application/octet-stream':
            st.warning(f"⚠️ 警告: ファイルのMIMEタイプ ({mime_type}) が想定と異なります。処理を継続します。")
        
        try:
            # Excelファイルを読み込む
            with st.spinner("ファイルを読み込んでいます..."):
                excel_data = pd.ExcelFile(uploaded_file)
                sheet_names = excel_data.sheet_names
            
            if not sheet_names:
                st.error("❌ エラー: ファイルにシートが含まれていません。")
            else:
                st.success(f"✅ ファイル '{uploaded_file.name}' を読み込みました")
                st.info(f"📊 シート数: {len(sheet_names)}")
                
                # 各シートの情報を表示
                with st.expander("シート情報を表示"):
                    for sheet_name in sheet_names:
                        try:
                            df_preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, nrows=5)
                            st.write(f"**{sheet_name}** - {len(df_preview.columns)}列")
                        except Exception as e:
                            st.write(f"**{sheet_name}** - 読み込みエラー: {str(e)}")
                
                # シート選択
                if len(sheet_names) > 1:
                    selected_sheets = st.multiselect(
                        "変換するシートを選択してください",
                        sheet_names,
                        default=sheet_names,
                        help="複数のシートを選択できます"
                    )
                else:
                    selected_sheets = sheet_names
                
                # 変換オプション
                st.subheader("⚙️ 変換オプション")
                include_index = st.checkbox("インデックスを含める", value=False)
                max_col_width = st.number_input("最大列幅", min_value=10, max_value=100, value=30)
                
                # 変換ボタン
                if st.button("🔄 Markdownに変換", type="primary", disabled=not selected_sheets):
                    markdown_content = ""
                    conversion_errors = []
                    successful_sheets = 0
                    markdown_results = {}
                    converted_sheets = []
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for idx, sheet_name in enumerate(selected_sheets):
                        try:
                            status_text.text(f"処理中: {sheet_name} ({idx + 1}/{len(selected_sheets)})")
                            
                            # シートを読み込む
                            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                            
                            # 空のDataFrameチェック
                            if df.empty:
                                conversion_errors.append(f"シート '{sheet_name}' は空です")
                                continue
                            
                            # データ型の処理（複雑なデータ型を文字列に変換）
                            for col in df.columns:
                                if df[col].dtype == 'object':
                                    df[col] = df[col].astype(str).replace('nan', '')
                            
                            # シート用のMarkdownを作成
                            sheet_markdown = f"## {sheet_name}\n\n"
                            sheet_markdown += f"*行数: {len(df)}, 列数: {len(df.columns)}*\n\n"
                            
                            # DataFrameをMarkdownテーブルに変換
                            try:
                                markdown_table = df.to_markdown(
                                    index=include_index,
                                    maxcolwidths=max_col_width
                                )
                                sheet_markdown += markdown_table
                            except Exception as e:
                                # to_markdownが失敗した場合の代替処理
                                sheet_markdown += "| " + " | ".join(df.columns) + " |\n"
                                sheet_markdown += "|" + "|".join(["---" for _ in df.columns]) + "|\n"
                                for _, row in df.iterrows():
                                    sheet_markdown += "| " + " | ".join(str(val) for val in row) + " |\n"
                                conversion_errors.append(f"シート '{sheet_name}': Markdown変換で警告 - 代替形式を使用")
                            
                            # 結果を保存
                            markdown_results[sheet_name] = sheet_markdown
                            markdown_content += sheet_markdown + "\n\n"
                            successful_sheets += 1
                            converted_sheets.append(sheet_name)
                            
                        except MemoryError:
                            conversion_errors.append(f"シート '{sheet_name}': メモリ不足エラー - データが大きすぎます")
                        except Exception as e:
                            conversion_errors.append(f"シート '{sheet_name}': {str(e)}")
                        
                        # プログレスバー更新
                        progress_bar.progress((idx + 1) / len(selected_sheets))
                    
                    # 完了処理
                    status_text.text("変換完了！")
                    progress_bar.empty()
                    
                    if successful_sheets > 0:
                        # セッション状態に保存
                        st.session_state.converted_markdown = markdown_content
                        st.session_state.filename = uploaded_file.name.replace('.xlsx', '.md')
                        st.session_state.conversion_errors = conversion_errors
                        st.session_state.markdown_results = markdown_results
                        st.session_state.converted_sheets = converted_sheets
                        
                        # 結果表示
                        st.success(f"✅ {successful_sheets}/{len(selected_sheets)} シートの変換に成功しました！")
                        
                        # エラーがあれば表示
                        if conversion_errors:
                            with st.expander("⚠️ 変換時の警告・エラー"):
                                for error in conversion_errors:
                                    st.warning(error)
                        
                        # 成功メッセージ
                        success_message = f"✅ ファイル '{uploaded_file.name}' から {successful_sheets} シートをMarkdown形式に変換しました！\n\n変換されたシート: {', '.join(converted_sheets)}"
                        st.session_state.messages.append({
                            "role": "assistant",
                            "content": success_message
                        })
                        
                        # 自動的に変換結果をプレビュー表示
                        preview_content = "📋 **変換結果のプレビュー**\n\n"
                        for sheet_name in converted_sheets:
                            sheet_content = markdown_results[sheet_name]
                            preview_content += f"### 📄 シート: {sheet_name}\n\n"
                            preview_content += f"```markdown\n{sheet_content[:1000]}{'...' if len(sheet_content) > 1000 else ''}\n```\n\n"
                        
                        preview_content += "\n💡 全体をダウンロードするには `/download` と入力してください。"
                        
                        st.session_state.messages.append({
                            "role": "assistant",
                            "content": preview_content
                        })
                    else:
                        st.error("❌ すべてのシートの変換に失敗しました")
                        with st.expander("エラー詳細"):
                            for error in conversion_errors:
                                st.error(error)
                
        except pd.errors.ParserError as e:
            st.error(f"❌ ファイル解析エラー: このファイルは正しいExcel形式ではない可能性があります。\n詳細: {str(e)}")
        except PermissionError:
            st.error("❌ ファイルアクセスエラー: ファイルが他のプログラムで開かれている可能性があります。")
        except Exception as e:
            st.error(f"❌ 予期しないエラーが発生しました: {str(e)}")
            st.info("💡 ヒント: ファイルが破損していないか、パスワード保護されていないか確認してください。")

# チャットインターフェース
chat_container = st.container()

with chat_container:
    # 既存のメッセージを表示
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])
    
    # ユーザー入力
    if prompt := st.chat_input("質問やコマンドを入力してください..."):
        # ユーザーメッセージを追加
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        
        # アシスタントの応答
        with st.chat_message("assistant"):
            if ("/download" in prompt.lower() or "ダウンロード" in prompt) and st.session_state.converted_markdown:
                # ダウンロードリンクを提供
                st.markdown("📥 Markdownファイルをダウンロードできます：")
                st.download_button(
                    label="Markdownファイルをダウンロード",
                    data=st.session_state.converted_markdown,
                    file_name=st.session_state.filename,
                    mime="text/markdown"
                )
                response = "ダウンロードボタンを表示しました。クリックしてMarkdownファイルをダウンロードしてください。"
            elif ("/preview" in prompt.lower() or "プレビュー" in prompt) and st.session_state.converted_markdown:
                # Markdownプレビューを表示
                st.markdown("### 📄 Markdownプレビュー")
                # プレビューの最初の1000文字を表示
                preview_text = st.session_state.converted_markdown[:1000]
                if len(st.session_state.converted_markdown) > 1000:
                    preview_text += "\n\n... (以下省略)"
                st.code(preview_text, language="markdown")
                response = "Markdownのプレビューを表示しました。"
            elif ("/error" in prompt.lower() or "エラー" in prompt) and st.session_state.conversion_errors:
                # エラー情報を表示
                st.markdown("### ⚠️ 変換時のエラー・警告")
                for error in st.session_state.conversion_errors:
                    st.warning(error)
                response = "変換時のエラー・警告を表示しました。"
            elif st.session_state.converted_markdown is None:
                response = "まだExcelファイルが変換されていません。サイドバーからファイルをアップロードして変換してください。"
            else:
                response = "以下のコマンドが使用できます:\n- `/download` または 'ダウンロード' - 変換したMarkdownファイルをダウンロード\n- `/preview` または 'プレビュー' - Markdownの内容を表示\n- `/error` または 'エラー' - 変換時のエラー・警告を表示"
            
            st.markdown(response)
            st.session_state.messages.append({"role": "assistant", "content": response})

# フッター
st.markdown("---")
st.markdown("💡 **使い方**: サイドバーからExcelファイルをアップロードし、変換ボタンをクリックしてください。")
st.markdown("📝 **対応形式**: .xlsx (Excel 2007以降)")
st.markdown("⚠️ **注意**: 大きなファイルや複雑な書式のファイルは変換に時間がかかる場合があります。")