import csv
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def load_default_speaker_data():
    """
    デフォルトの発言者データを辞書形式で返します。
    """
    return {
        "エマ": {"color": "ff69b4","length": 3},
        "ヒロ": {"color": "dc143c","length": 3},
        "アンアン": {"color": "6a5acd","length": 3},
        "シェリー": {"color": "6495ed","length": 3},
        "ハンナ": {"color": "9acd32","length": 3},
        "ノア": {"color": "87cefa","length": 3},
        "レイア": {"color": "ff7f50","length": 3},
        "ミリア": {"color": "ffa500","length": 3},
        "ココ": {"color": "ff8c00","length": 3},
        "マーゴ": {"color": "8a2be2","length": 3},
        "ナノカ": {"color": "696969","length": 3},
        "アリサ": {"color": "800000","length": 3},
        "メルル": {"color": "dda0dd","length": 3},
        "シロ": {"color": "867ba9","length": 3},
        "ユキ": {"color": "e6e6fa","length": 3},
        "ゴクチョー": {"color": "778899","length": 3},
    }

def load_speaker_data(csv_filepath):
    """
    CSVファイルから発言者データを読み込み、辞書形式で返します。
    """
    # ★ tkinter版では、スクリプトの場所を __file__ から取得します
    script_dir = os.path.dirname(os.path.abspath(__file__))
    absolute_csv_path = os.path.join(script_dir, csv_filepath)
    
    speaker_data = {}
    try:
        with open(absolute_csv_path, mode='r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if 'name' in row and 'color' in row and 'length' in row:
                    try:
                        speaker_data[row['name']] = {
                            "color": row['color'],
                            "length": int(row['length'])
                        }
                    except ValueError:
                        print(f"警告: name '{row['name']}' の length '{row['length']}' は数値ではありません。スキップします。")
                else:
                    print(f"警告: CSVの行が不正です（name, color, length が必要）: {row}")

    except FileNotFoundError:
        print(f"エラー: CSVファイルが見つかりません: {absolute_csv_path}")
        # ★ エラー時は None を返します
        return None
    except Exception as e:
        print(f"エラー: CSVファイルの読み込み中にエラーが発生しました: {e}")
        return None
        
    return speaker_data

def highlight_speaker_text(text, speaker_data):
    """
    発言者ハイライト処理
    """
    for name, data in speaker_data.items():
        for bracket in ["「", "『"]:
            prefix = name + bracket
            if text.startswith(prefix):
                color = data['color']
                length = data['length']
                remaining_text = text[len(prefix):]
                highlight_part = remaining_text[:length]
                if '{' in highlight_part or '}' in highlight_part:
                    return remaining_text
                else:
                    rest_part = remaining_text[length:]
                    # 開始括弧に合わせて閉じ括弧も変える
                    close_bracket = "」" if bracket == "「" else "』"
                    return f"{bracket}《color:#{color}》{highlight_part}《/color》{rest_part}{close_bracket if not rest_part.endswith(close_bracket) else ''}"
    return text

def format_custom_links(text):
    """
    カスタムリンク置換処理
    """
    pattern = r'\{([^,]+),([^,]+),([^\}]+)\}'
    
    def replacer(match):
        link_id = match.group(1).strip()
        display_text = match.group(2).strip()
        ref_text = match.group(3).strip()
        return f"《link:#{link_id}》《red》{display_text}《/red》《/link》《ref》{ref_text}《/ref》"
        
    return re.sub(pattern, replacer, text)

def main_gui():
    
    # --- CSVファイル名はここで固定 ---
    csv_filename = "speakers.csv"

    # --- (A) ボタンが押されたときの関数 ---
    
    def select_input_file():
        """ 「入力ファイル選択...」ボタンが押されたときの処理 """
        # ファイル選択ダイアログを開く
        # .txtファイルのみ表示
        filepath = filedialog.askopenfilename(
            title="処理するテキストファイルを選択",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if filepath:
            # 選択されたファイルパスをテキストボックスに表示
            # (StringVar を使って値を設定)
            input_path_var.set(filepath)
            status_var.set("ファイルが選ばれました。「実行」を押してください。")

    def run_processing():
        """ 「実行」ボタンが押されたときの処理 """
        
        # 1. 入力ファイルパスを取得 (StringVar から値を取得)
        input_filepath = input_path_var.get()

        # 2. ファイルが選ばれているかチェック
        if not input_filepath:
            messagebox.showerror("エラー", "まず「入力ファイル選択...」ボタンからファイルを選んでください。")
            return # 処理を中断

        # 3. 出力ファイルパスを自動で決める
        base, ext = os.path.splitext(input_filepath)
        output_filepath = f"{base}_output{ext}"

        # 4. ステータスを「処理中」に更新
        status_var.set("処理中...")
        root.update_idletasks() # 表示を強制更新

        try:
            # 5. CSVを読み込む
            #speaker_data = load_default_speaker_data()
            speaker_data = load_speaker_data(csv_filename)

            # CSVの読み込みに失敗した場合
            if speaker_data is None:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                abs_csv_path = os.path.join(script_dir, csv_filename)
                messagebox.showerror("CSVエラー", f"{csv_filename} が見つかりません。\nスクリプトと同じ場所に置いてください。\n\n(検索パス: {abs_csv_path})")
                status_var.set("CSVファイルが見つかりません。")
                return

            # 6. メインのファイル処理
            line_count = 0
            with open(output_filepath, mode='w', encoding='utf-8') as outfile:
                with open(input_filepath, mode='r', encoding='utf-8') as infile:
                    for line in infile:
                        line_count += 1
                        original_text = line.strip()
                        
                        if not original_text:
                            outfile.write('\n')
                            continue
                        
                        # 処理1: ハイライト
                        processed_text_step1 = highlight_speaker_text(original_text, speaker_data)
                        
                        # 処理2: リンク
                        final_processed_text = format_custom_links(processed_text_step1)
                        
                        # ファイルに書き込み
                        outfile.write(final_processed_text + '\n')
            
            # 7. 完了メッセージ
            status_var.set(f"完了！ ({line_count}行処理)")
            messagebox.showinfo("処理完了", f"処理が完了しました。\n{output_filepath}\nに保存されました。")

        except FileNotFoundError:
            messagebox.showerror("エラー", f"入力ファイルが見つかりません:\n{input_filepath}")
            status_var.set("エラー: 入力ファイルが見つかりません。")
        except Exception as e:
            # その他の予期せぬエラー
            messagebox.showerror("予期せぬエラー", f"エラーが発生しました:\n{e}")
            status_var.set("予期せぬエラーが発生しました。")

    # --- (B) ウィンドウと部品の作成 ---
    
    root = tk.Tk()
    root.title("会話テキスト置換スクリプト")
    root.geometry("450x180") # ウィンドウサイズ

    # メインフレーム (ウィンドウの余白(padding)設定)
    main_frame = ttk.Frame(root, padding="15")
    main_frame.pack(fill="both", expand=True)

    # --- 部品を配置 ---
    
    # 1行目: 説明ラベル
    instruction_label = ttk.Label(main_frame, text="処理したいテキストファイルを選んでください。")
    instruction_label.pack(anchor="w")

    # 2行目: ファイルパス表示
    # (tk.StringVar: tkinter用の特別な変数で、中身が変わると表示も連動して変わる)
    input_path_var = tk.StringVar()
    input_path_entry = ttk.Entry(main_frame, textvariable=input_path_var, state="readonly", width=60)
    input_path_entry.pack(fill="x", pady=5)

    # 3行目: ボタン類 (横並び)
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill="x", pady=5)
    
    # 3-1: ファイル選択ボタン
    select_button = ttk.Button(button_frame, text="入力ファイル選択...", command=select_input_file)
    select_button.pack(side="left")

    # 3-2: 実行ボタン (右寄せ)
    run_button = ttk.Button(button_frame, text="実行", command=run_processing, style="Accent.TButton")
    run_button.pack(side="right", padx=(10, 0))

    # (実行ボタンのスタイルを定義)
    style = ttk.Style()
    style.configure("Accent.TButton", font=("", 10, "bold"))

    # 4行目: ステータス表示ラベル
    status_var = tk.StringVar()
    status_var.set("ファイルを選んでください。")
    status_label = ttk.Label(main_frame, textvariable=status_var, foreground="blue")
    status_label.pack(anchor="center", pady=10)

    # --- (C) ウィンドウの実行 ---
    root.mainloop()

# -----------------------------------------------------------------
# (3) スクリプトの実行
# -----------------------------------------------------------------
if __name__ == "__main__":
    main_gui()