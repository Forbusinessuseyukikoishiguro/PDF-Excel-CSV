import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pdfplumber
import shutil
import os
from datetime import datetime
import sys

class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to CSV/Excel 変換ツール")
        self.root.geometry("600x500")
        
        # メインフレーム
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # タイトル
        title_label = ttk.Label(main_frame, text="PDF to CSV/Excel 変換ツール", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 入力ファイル選択
        ttk.Label(main_frame, text="変換元PDFファイル:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.input_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.input_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="参照", command=self.browse_input_file).grid(row=1, column=2, padx=5)
        
        # 出力フォルダ選択
        ttk.Label(main_frame, text="出力フォルダ:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.output_dir = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="参照", command=self.browse_output_dir).grid(row=2, column=2, padx=5)
        
        # 変換形式選択
        ttk.Label(main_frame, text="変換形式:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.output_format = tk.StringVar(value="csv")
        format_frame = ttk.Frame(main_frame)
        format_frame.grid(row=3, column=1, sticky=tk.W, pady=5)
        ttk.Radiobutton(format_frame, text="CSV", variable=self.output_format, value="csv").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(format_frame, text="Excel", variable=self.output_format, value="excel").pack(side=tk.LEFT, padx=5)
        
        # オプション設定
        options_frame = ttk.LabelFrame(main_frame, text="オプション設定", padding="10")
        options_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        self.create_backup = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="元ファイルのバックアップを作成", 
                       variable=self.create_backup).pack(anchor=tk.W)
        
        self.add_timestamp = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="ファイル名にタイムスタンプを追加", 
                       variable=self.add_timestamp).pack(anchor=tk.W)
        
        # 詳細設定
        details_frame = ttk.LabelFrame(main_frame, text="PDF読み取り設定", padding="10")
        details_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(details_frame, text="開始ページ (空白で全ページ):").pack(anchor=tk.W)
        self.start_page = tk.StringVar()
        ttk.Entry(details_frame, textvariable=self.start_page, width=10).pack(anchor=tk.W, pady=2)
        
        ttk.Label(details_frame, text="終了ページ (空白で最後まで):").pack(anchor=tk.W)
        self.end_page = tk.StringVar()
        ttk.Entry(details_frame, textvariable=self.end_page, width=10).pack(anchor=tk.W, pady=2)
        
        # 実行ボタン
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="変換実行", command=self.convert_pdf, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="クリア", command=self.clear_fields).pack(side=tk.LEFT, padx=5)
        
        # プログレスバー
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # ログ表示エリア
        log_frame = ttk.LabelFrame(main_frame, text="ログ", padding="5")
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初期値設定
        self.output_dir.set(os.getcwd())
        
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="PDFファイルを選択",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.input_path.set(filename)
    
    def browse_output_dir(self):
        dirname = filedialog.askdirectory(title="出力フォルダを選択")
        if dirname:
            self.output_dir.set(dirname)
    
    def clear_fields(self):
        self.input_path.set("")
        self.start_page.set("")
        self.end_page.set("")
        self.log_text.delete(1.0, tk.END)
    
    def log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def create_backup_file(self, input_file):
        try:
            backup_dir = os.path.join(os.path.dirname(input_file), "backup")
            os.makedirs(backup_dir, exist_ok=True)
            
            filename = os.path.basename(input_file)
            name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{name}_backup_{timestamp}{ext}"
            backup_path = os.path.join(backup_dir, backup_filename)
            
            shutil.copy2(input_file, backup_path)
            self.log(f"バックアップ作成: {backup_path}")
            return backup_path
        except Exception as e:
            self.log(f"バックアップ作成エラー: {str(e)}")
            return None
    
    def extract_tables_from_pdf(self, pdf_path, start_page=None, end_page=None):
        tables = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                self.log(f"PDF総ページ数: {total_pages}")
                
                # ページ範囲の設定
                start = start_page - 1 if start_page else 0
                end = end_page if end_page else total_pages
                
                self.log(f"処理ページ範囲: {start + 1} - {end}")
                
                for page_num in range(start, min(end, total_pages)):
                    page = pdf.pages[page_num]
                    self.log(f"ページ {page_num + 1} を処理中...")
                    
                    # テーブル抽出を試行
                    page_tables = page.extract_tables()
                    
                    if page_tables:
                        for i, table in enumerate(page_tables):
                            if table:  # 空でないテーブルのみ
                                df = pd.DataFrame(table[1:], columns=table[0])  # 最初の行をヘッダーとして使用
                                df.name = f"Page_{page_num + 1}_Table_{i + 1}"
                                tables.append(df)
                                self.log(f"テーブル発見: ページ{page_num + 1}, テーブル{i + 1} ({len(df)}行)")
                    
                    # テーブルが見つからない場合、テキストとして抽出
                    if not page_tables:
                        text = page.extract_text()
                        if text:
                            # 改行で分割して簡易的なテーブル作成
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            if lines:
                                # スペースまたはタブで分割
                                rows = []
                                for line in lines:
                                    # 複数のスペースまたはタブで分割
                                    row = [cell.strip() for cell in line.split() if cell.strip()]
                                    if row:
                                        rows.append(row)
                                
                                if rows:
                                    # 最大列数を求める
                                    max_cols = max(len(row) for row in rows)
                                    # 全行を同じ列数にする
                                    for row in rows:
                                        while len(row) < max_cols:
                                            row.append("")
                                    
                                    df = pd.DataFrame(rows[1:] if len(rows) > 1 else rows, 
                                                    columns=rows[0] if len(rows) > 1 else [f"列{i+1}" for i in range(max_cols)])
                                    df.name = f"Page_{page_num + 1}_Text"
                                    tables.append(df)
                                    self.log(f"テキスト抽出: ページ{page_num + 1} ({len(df)}行)")
                
        except Exception as e:
            self.log(f"PDF読み取りエラー: {str(e)}")
            raise e
        
        return tables
    
    def convert_pdf(self):
        # 入力チェック
        if not self.input_path.get():
            messagebox.showerror("エラー", "PDFファイルを選択してください")
            return
        
        if not os.path.exists(self.input_path.get()):
            messagebox.showerror("エラー", "選択されたPDFファイルが存在しません")
            return
        
        if not self.output_dir.get():
            messagebox.showerror("エラー", "出力フォルダを指定してください")
            return
        
        try:
            self.progress.start()
            self.log("変換処理を開始します...")
            
            # バックアップ作成
            if self.create_backup.get():
                self.create_backup_file(self.input_path.get())
            
            # ページ範囲の解析
            start_page = None
            end_page = None
            
            if self.start_page.get():
                try:
                    start_page = int(self.start_page.get())
                except ValueError:
                    messagebox.showerror("エラー", "開始ページは数値で入力してください")
                    return
            
            if self.end_page.get():
                try:
                    end_page = int(self.end_page.get())
                except ValueError:
                    messagebox.showerror("エラー", "終了ページは数値で入力してください")
                    return
            
            # PDFからテーブル抽出
            self.log("PDFからデータを抽出中...")
            tables = self.extract_tables_from_pdf(self.input_path.get(), start_page, end_page)
            
            if not tables:
                messagebox.showwarning("警告", "PDFからテーブルデータが見つかりませんでした")
                return
            
            # 出力ファイル名の生成
            input_filename = os.path.splitext(os.path.basename(self.input_path.get()))[0]
            timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S") if self.add_timestamp.get() else ""
            
            if self.output_format.get() == "csv":
                if len(tables) == 1:
                    # 単一テーブルの場合
                    output_filename = f"{input_filename}{timestamp}.csv"
                    output_path = os.path.join(self.output_dir.get(), output_filename)
                    tables[0].to_csv(output_path, index=False, encoding='utf-8-sig')
                    self.log(f"CSV出力完了: {output_path}")
                else:
                    # 複数テーブルの場合、個別ファイルとして保存
                    for i, table in enumerate(tables):
                        output_filename = f"{input_filename}_{table.name}{timestamp}.csv"
                        output_path = os.path.join(self.output_dir.get(), output_filename)
                        table.to_csv(output_path, index=False, encoding='utf-8-sig')
                        self.log(f"CSV出力完了: {output_path}")
            
            elif self.output_format.get() == "excel":
                output_filename = f"{input_filename}{timestamp}.xlsx"
                output_path = os.path.join(self.output_dir.get(), output_filename)
                
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for table in tables:
                        sheet_name = table.name if hasattr(table, 'name') else f"Sheet_{tables.index(table) + 1}"
                        # シート名の長さ制限（Excel仕様）
                        sheet_name = sheet_name[:31]
                        table.to_excel(writer, sheet_name=sheet_name, index=False)
                
                self.log(f"Excel出力完了: {output_path}")
            
            self.log(f"変換完了！ {len(tables)}個のテーブルを処理しました。")
            messagebox.showinfo("完了", "変換が完了しました")
            
        except Exception as e:
            error_msg = f"変換エラー: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("エラー", error_msg)
        
        finally:
            self.progress.stop()

def main():
    # 必要なライブラリのインストール確認
    try:
        import pdfplumber
        import pandas
        import openpyxl
    except ImportError as e:
        print(f"必要なライブラリがインストールされていません: {e}")
        print("以下のコマンドでインストールしてください:")
        print("pip install pdfplumber pandas openpyxl")
        return
    
    root = tk.Tk()
    app = PDFConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
