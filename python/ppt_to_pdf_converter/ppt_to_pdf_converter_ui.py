import os
import sys
import platform
import datetime
import time
import argparse
from pathlib import Path
import logging
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import subprocess # macOS 用

# pywin32 の条件付きインポート (Windowsのみ)
HAS_PYWIN32 = False
if platform.system() == "Windows":
    try:
        import win32com.client
        import pythoncom # COM初期化に必要
        HAS_PYWIN32 = True
    except ImportError:
        print("警告: 'pywin32' ライブラリが見つかりません。Windows環境でのPPT/PPTX変換には必要です。`pip install pywin32` でインストールしてください。", file=sys.stderr)

# --- UIウィジェットクラス (context_folder_extracter_ui.py より流用・改変) ---
class RoundedButton(tk.Canvas):
    """角丸のモダンなボタン"""
    def __init__(self, parent, text, command=None, width=100, height=36,
                 bg_color="#2F4F4F", fg_color="white", font_family="Arial",
                 font_size=10, font_weight="bold", corner_radius=10, hover_color="#3A6363", **kwargs):
        super().__init__(parent, width=width, height=height,
                        bg=parent["bg"], bd=0, highlightthickness=0, **kwargs)

        self.bg_color = bg_color
        self.fg_color = fg_color
        self.hover_color = hover_color
        self.corner_radius = corner_radius
        self.current_color = bg_color
        self.command = command
        self.text = text

        # ボタンの描画
        self.button_bg = self.create_rounded_rect(0, 0, width, height, corner_radius, fill=bg_color)
        self.button_text = self.create_text(width//2, height//2, text=text, fill=fg_color,
                                           font=(font_family, font_size, font_weight))

        # イベントバインド
        self.bind("<Enter>", self.on_hover)
        self.bind("<Leave>", self.on_leave)
        self.bind("<ButtonPress-1>", self.on_press)
        self.bind("<ButtonRelease-1>", self.on_release)

    def create_rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        """角丸の長方形を描画"""
        points = [
            x1+r, y1,
            x2-r, y1,
            x2, y1,
            x2, y1+r,
            x2, y2-r,
            x2, y2,
            x2-r, y2,
            x1+r, y2,
            x1, y2,
            x1, y2-r,
            x1, y1+r,
            x1, y1
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

    def on_hover(self, event):
        """ホバー状態の処理"""
        if getattr(self, '_state', 'normal') == 'disabled': return # 無効時は何もしない
        self.current_color = self.hover_color
        self.itemconfig(self.button_bg, fill=self.hover_color)

    def on_leave(self, event):
        """通常状態に戻る処理"""
        if getattr(self, '_state', 'normal') == 'disabled': return
        self.current_color = self.bg_color
        self.itemconfig(self.button_bg, fill=self.bg_color)

    def on_press(self, event):
        """ボタンプレス時の処理"""
        if getattr(self, '_state', 'normal') == 'disabled': return
        self.itemconfig(self.button_bg, fill=self.fg_color)
        self.itemconfig(self.button_text, fill=self.bg_color)

    def on_release(self, event):
        """ボタンリリース時の処理"""
        if getattr(self, '_state', 'normal') == 'disabled': return
        self.itemconfig(self.button_bg, fill=self.current_color)
        self.itemconfig(self.button_text, fill=self.fg_color)
        if self.command:
            self.command()

    # state プロパティを追加して、標準ウィジェットのように扱えるようにする
    def config(self, **kwargs):
        if 'state' in kwargs:
            new_state = kwargs.pop('state')
            setattr(self, '_state', new_state) # カスタム属性として保存
            if new_state == 'disabled':
                # 無効化時の見た目を変更
                super().itemconfig(self.button_bg, fill="#AAA", outline="#AAA")
                super().itemconfig(self.button_text, fill="#777")
            elif new_state == 'normal':
                # 有効化時の見た目に戻す
                super().itemconfig(self.button_bg, fill=self.bg_color, outline=self.bg_color)
                super().itemconfig(self.button_text, fill=self.fg_color)
                self.current_color = self.bg_color # Leave状態の色に戻す
            else:
                 pass
        super().config(**kwargs)

    # configure メソッドもオーバーライドして state に対応させる
    def configure(self, **kwargs):
        self.config(**kwargs)


class ModernEntryFrame(tk.Frame):
    """モダンな入力欄を持つフレーム"""
    def __init__(self, parent, label_text, default_value="", width=250, button_text=None, button_command=None, **kwargs):
        super().__init__(parent, bg=parent["bg"], **kwargs)

        self.var = tk.StringVar(value=default_value)

        # ラベル
        self.label = tk.Label(self, text=label_text, bg=self["bg"],
                            fg="#333", font=("Arial", 10, "bold"))
        self.label.pack(anchor="w", pady=(5, 2))

        # 入力とボタンを横に並べるフレーム
        input_button_frame = tk.Frame(self, bg=parent["bg"])
        input_button_frame.pack(fill="x")

        # 入力フレーム（影のエフェクト用）
        self.entry_frame = tk.Frame(input_button_frame, bg="#DDD", padx=1, pady=1)
        self.entry_frame.pack(side=tk.LEFT, fill="x", expand=True, pady=(0, 5))

        # 実際の入力欄
        self.entry = tk.Entry(self.entry_frame, textvariable=self.var,
                            font=("Arial", 10), bd=0, bg="white", fg="black")
        self.entry.pack(fill="x", ipady=8, padx=1, pady=1)

        # 参照ボタン (もしあれば)
        if button_text and button_command:
            self.browse_button = RoundedButton(input_button_frame, button_text,
                                        command=button_command, width=60, height=30,
                                        bg_color="#999", hover_color="#777")
            self.browse_button.pack(side=tk.RIGHT, padx=(5, 0), pady=(0, 5), anchor='s')


    def get(self):
        """入力された値を取得"""
        return self.var.get()

    def set(self, value):
        """値を設定"""
        self.var.set(value)
# --- UIウィジェットクラス ここまで ---

# --- ロガー設定 ---
try:
    # スクリプトのおかれた場所を基準にoutputを探す
    script_dir = Path(__file__).resolve().parent
    output_dir_default = script_dir / "output"
    output_dir_default.mkdir(exist_ok=True) # なければ作成
    log_file_path = output_dir_default / "ppt_to_pdf_converter.log"
except Exception as e:
    # フォールバック
    script_dir = Path.cwd() # カレントディレクトリ
    output_dir_default = script_dir / "output"
    # output作成試行、失敗してもログはカレントに
    try:
        output_dir_default.mkdir(exist_ok=True)
    except Exception:
        pass # 作成失敗時はログファイルだけカレントに
    log_file_path = script_dir / "ppt_to_pdf_converter.fallback.log"
    print(f"警告: outputディレクトリの準備に失敗しました。ログは {log_file_path} に出力されます。エラー: {e}", file=sys.stderr)


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, mode="a", encoding="utf-8"), # 追記モード: UI版はファイルにもログを書き込む
        logging.StreamHandler() # コンソールにも出す (デバッグ用)
    ]
)
logger = logging.getLogger(__name__)

logger.info("-" * 50)
logger.info("PPT to PDF Converter (UI) 起動")
logger.info(f"スクリプトディレクトリ: {script_dir}")
logger.info(f"デフォルト出力ディレクトリ: {output_dir_default}")
logger.info(f"ログファイルパス: {log_file_path}")
logger.info(f"プラットフォーム: {platform.system()}")
if platform.system() == "Windows":
    logger.info(f"pywin32 利用可能: {HAS_PYWIN32}")
# --- ロガー設定 ここまで ---


# --- PDF変換ロジック (CLI版と共通化してもよいが、ここではUI版内に保持) ---
def convert_ppt_to_pdf_windows(ppt_path: Path, pdf_path: Path):
    """
    Windows環境でPowerPointファイルをPDFに変換する（COM使用）
    """
    if not HAS_PYWIN32:
        raise ImportError("pywin32ライブラリが必要です。")

    powerpoint = None
    presentation = None
    abs_ppt_path = str(ppt_path.resolve())
    abs_pdf_path = str(pdf_path.resolve())
    logger.debug(f"Windows変換開始: {abs_ppt_path} -> {abs_pdf_path}") # フルパス表示

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(abs_ppt_path, ReadOnly=True, Untitled=False, WithWindow=False)
        presentation.SaveAs(abs_pdf_path, 32) # ppFixedFormatTypePDF = 32
        logger.info(f"変換成功 (Win): {abs_ppt_path} -> {abs_pdf_path}") # フルパス表示
        return True
    except pythoncom.com_error as e:
         logger.error(f"COMエラー (Win): {abs_ppt_path} の変換中にエラー発生: {e}", exc_info=True) # UI版は詳細ログを出す
         hresult = getattr(e, 'hresult', 'N/A')
         if hresult == -2147221005:
             raise RuntimeError(f"PowerPointがインストールされていないか、COM登録に問題があります。({abs_ppt_path})") from e # フルパス表示
         elif hresult == -2147023174:
              raise RuntimeError(f"PowerPointのプロセスに接続できませんでした。({abs_ppt_path})") from e # フルパス表示
         else:
              raise RuntimeError(f"PowerPointでの変換中にCOMエラー (HRESULT: {hresult}) ({abs_ppt_path})") from e # フルパス表示
    except Exception as e:
        logger.error(f"予期せぬエラー (Win): {abs_ppt_path} の変換中にエラー発生: {e}", exc_info=True) # UI版は詳細ログを出す, フルパス表示
        raise RuntimeError(f"PowerPointでの変換中に予期せぬエラー: {e} ({abs_ppt_path})") from e # フルパス表示
    finally:
        try:
            if presentation: presentation.Close()
        except Exception as close_e:
            logger.warning(f"PowerPointプレゼンテーション終了処理中にエラー: {close_e}")
        pythoncom.CoUninitialize()

def convert_ppt_to_pdf_macos(ppt_path: Path, pdf_path: Path):
    """
    macOS環境でPowerPointファイルをPDFに変換する（AppleScript使用）
    """
    logger.debug(f"macOS変換開始: {ppt_path.resolve()} -> {pdf_path.resolve()}") # フルパス表示
    abs_ppt_path = str(ppt_path.resolve())
    abs_pdf_path = str(pdf_path.resolve())

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # 1. PowerPoint for Mac を試す
    powerpoint_script = f'''
    tell application "Microsoft PowerPoint"
        activate
        open "{abs_ppt_path}"
        try
            save active presentation in "{abs_pdf_path}" as save as PDF
            set success to true
        on error errMsg number errNum
            set success to false
            log "PowerPoint Error: " & errMsg & " (" & errNum & ")"
        end try
        close active presentation saving no
        if success then return "PowerPoint Success"
    end tell
    return "PowerPoint Failed"
    '''
    try:
        logger.debug("AppleScript (PowerPoint) を実行")
        process = subprocess.run(['osascript', '-e', powerpoint_script], capture_output=True, text=True, check=False, timeout=120)
        logger.debug(f"AppleScript (PowerPoint) 結果: stdout='{process.stdout.strip()}', stderr='{process.stderr.strip()}'")
        if process.returncode == 0 and "PowerPoint Success" in process.stdout:
            logger.info(f"変換成功 (Mac/PowerPoint): {abs_ppt_path} -> {abs_pdf_path}") # フルパス表示
            return True
        else:
            logger.warning(f"PowerPoint for Mac での変換失敗/見つからず。Keynote試行。エラー: {process.stderr.strip()}")
    except FileNotFoundError:
        logger.warning("osascriptが見つかりません。Keynote試行。")
    except subprocess.TimeoutExpired:
        logger.error(f"AppleScript (PowerPoint) タイムアウト: {abs_ppt_path}") # フルパス表示
    except Exception as e:
        logger.error(f"AppleScript (PowerPoint) 予期せぬエラー: {e}", exc_info=True) # UI版は詳細ログ

    # 2. Keynote を試す
    keynote_script = f'''
    tell application "Keynote"
        activate
        try
            set theDoc to open "{abs_ppt_path}"
            delay 1
            export theDoc to file "{abs_pdf_path}" as PDF with properties {{image quality:Best}}
            set success to true
        on error errMsg number errNum
            set success to false
            log "Keynote Error: " & errMsg & " (" & errNum & ")"
        end try
        if success then
            close theDoc saving no
            return "Keynote Success"
        end if
    end tell
    return "Keynote Failed"
    '''
    try:
        logger.debug("AppleScript (Keynote) を実行")
        process = subprocess.run(['osascript', '-e', keynote_script], capture_output=True, text=True, check=False, timeout=120)
        logger.debug(f"AppleScript (Keynote) 結果: stdout='{process.stdout.strip()}', stderr='{process.stderr.strip()}'")
        if process.returncode == 0 and "Keynote Success" in process.stdout:
            logger.info(f"変換成功 (Mac/Keynote): {abs_ppt_path} -> {abs_pdf_path}") # フルパス表示
            return True
        else:
            logger.error(f"Keynote変換失敗。エラー: {process.stderr.strip()}")
            raise RuntimeError(f"macOS変換失敗 (両方)。({abs_ppt_path}) エラー: {process.stderr.strip()}") # フルパス表示
    except FileNotFoundError:
        raise RuntimeError("macOS変換に必要なosascriptが見つかりません。")
    except subprocess.TimeoutExpired:
         raise RuntimeError(f"Keynote変換タイムアウト。({abs_ppt_path})") # フルパス表示
    except Exception as e:
        raise RuntimeError(f"Keynote変換中に予期せぬエラー: {e} ({abs_ppt_path})") from e # フルパス表示

# --- PDF変換ロジック ここまで ---


# --- メインUIクラス ---
class ModernPPTtoPDFConverterUI:
    """PPT/PPTX to PDF Converter UI"""
    def __init__(self, root):
        self.root = root
        self.root.title("PPT/PPTX to PDF Converter (UI)") # タイトルにUI版と明記

        # ウィンドウサイズと位置を設定
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = 700 # 幅を少し広げる
        window_height = 700 # 高さを少し広げる
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # 背景色を設定
        self.bg_color = "#F5F5F7"
        self.accent_color = "#D2691E" # Chocolate color for accent
        self.root.configure(bg=self.bg_color)

        # メインコンテナ
        self.main_container = tk.Frame(root, bg=self.bg_color, padx=20, pady=20)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        # ヘッダーフレーム
        self.header_frame = tk.Frame(self.main_container, bg=self.bg_color)
        self.header_frame.pack(fill=tk.X, pady=(0, 10))

        # タイトル
        title_label = tk.Label(self.header_frame, text="PPT/PPTX to PDF Converter (UI)", # タイトルにUI版と明記
                             font=("Arial", 16, "bold"), bg=self.bg_color, fg=self.accent_color)
        title_label.pack(side=tk.LEFT)

        # --- 入力設定フレーム (カード風) ---
        self.input_frame_outer = tk.Frame(self.main_container, bg=self.bg_color, pady=5)
        self.input_frame_outer.pack(fill=tk.X)
        self.input_frame_card = tk.Frame(self.input_frame_outer, bg="white", padx=15, pady=15,
                                         highlightbackground="#DDD", highlightthickness=1)
        self.input_frame_card.pack(fill=tk.X)

        # --- 入力パスリスト部分 --- 
        input_list_label = tk.Label(self.input_frame_card, text="変換対象のファイル/フォルダ",
                                  bg="white", fg="#333", font=("Arial", 10, "bold"))
        input_list_label.pack(anchor="w", pady=(0, 5))

        input_list_frame = tk.Frame(self.input_frame_card, bg="white")
        input_list_frame.pack(fill=tk.X, expand=True)

        # リストボックス
        self.input_listbox = tk.Listbox(input_list_frame, selectmode=tk.EXTENDED,
                                        font=("Arial", 9), height=6,
                                        bg="#FFF", fg="#222", bd=1, relief="solid")
        self.input_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # リストボックス用スクロールバー
        listbox_scrollbar = tk.Scrollbar(input_list_frame, orient=tk.VERTICAL)
        listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.input_listbox.config(yscrollcommand=listbox_scrollbar.set)
        listbox_scrollbar.config(command=self.input_listbox.yview)

        # --- 入力リスト操作ボタン --- 
        input_button_frame = tk.Frame(self.input_frame_card, bg="white", pady=5)
        input_button_frame.pack(fill=tk.X)

        self.add_file_button = RoundedButton(input_button_frame, "ファイル追加",
                                        command=self.add_files, width=100, height=30,
                                        bg_color="#88A", hover_color="#AAB")
        self.add_file_button.pack(side=tk.LEFT, padx=2)

        self.add_folder_button = RoundedButton(input_button_frame, "フォルダ追加",
                                          command=self.add_folder, width=100, height=30,
                                          bg_color="#8A8", hover_color="#ABA")
        self.add_folder_button.pack(side=tk.LEFT, padx=2)

        self.remove_button = RoundedButton(input_button_frame, "選択削除",
                                      command=self.remove_selected, width=80, height=30,
                                      bg_color="#A88", hover_color="#BAA")
        self.remove_button.pack(side=tk.LEFT, padx=2)

        self.clear_button = RoundedButton(input_button_frame, "全削除",
                                     command=self.clear_list, width=70, height=30,
                                     bg_color="#AAA", hover_color="#BBB")
        self.clear_button.pack(side=tk.LEFT, padx=2)

        # --- 区切り線 --- 
        separator = tk.Frame(self.input_frame_card, height=1, bg="#EEE")
        separator.pack(fill=tk.X, pady=10)

        # --- 出力フォルダ指定 --- 
        self.output_dir_entry = ModernEntryFrame(self.input_frame_card, "PDF出力先フォルダ",
                                                 default_value=str(output_dir_default), # デフォルトを設定
                                                 button_text="参照",
                                                 button_command=self.browse_output_dir)
        self.output_dir_entry.pack(fill=tk.X, pady=(5, 0)) # 下のpadyを0に

        # --- ボタンフレーム (変換開始) --- 
        self.action_frame = tk.Frame(self.main_container, bg=self.bg_color, pady=15)
        self.action_frame.pack(fill=tk.X)

        # 変換開始ボタン
        self.convert_button = RoundedButton(self.action_frame, "変換開始",
                                          command=self.run_conversion, width=120, height=40,
                                          bg_color=self.accent_color, hover_color="#E5975E", # Lighter hover
                                          font_size=12)
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        self.convert_button.configure(state='normal') # 初期状態を明示

        # --- ログ表示フレーム ---
        self.log_frame = tk.Frame(self.main_container, bg=self.bg_color, pady=10)
        self.log_frame.pack(fill=tk.BOTH, expand=True)

        log_label = tk.Label(self.log_frame, text="ログ:",
                             font=("Arial", 10), bg=self.bg_color, fg="#333")
        log_label.pack(anchor="w", pady=(0, 5))

        self.log_container = tk.Frame(self.log_frame, bg="#DDD", padx=1, pady=1)
        self.log_container.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(self.log_container, font=("Consolas", 9),
                                  wrap=tk.WORD, bd=0, padx=10, pady=10,
                                  bg="white", fg="black", height=10, state=tk.DISABLED) # 高さを調整
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        log_scrollbar = tk.Scrollbar(self.log_container)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        log_scrollbar.config(command=self.log_text.yview)

        # --- ステータスバー ---
        self.status_var = tk.StringVar(value="準備完了")
        self.status_bar = tk.Label(self.root, textvariable=self.status_var,
                                  bg="#E5E5E7", fg="#666", font=("Arial", 9),
                                  bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=10, pady=2)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # キューの初期化
        self.status_queue = queue.Queue()
        self.log_queue = queue.Queue() # ログ用キュー

        # キューチェックのスケジュール
        self.root.after(100, self.check_queues)

        # 初期フォーカス (リストボックスか最初のボタンに)
        self.input_listbox.focus_set()

    # --- リストボックス操作メソッド --- 
    def add_files(self):
        """ファイル選択ダイアログを開き、選択されたPPT/PPTXファイルをリストに追加"""
        # filetypes で PPT/PPTX のみをフィルタリング
        filetypes = [("PowerPoint ファイル", "*.ppt *.pptx"), ("すべてのファイル", "*.*")]
        filenames = filedialog.askopenfilenames(title="変換するファイルを選択", filetypes=filetypes)
        if filenames:
            current_items = set(self.input_listbox.get(0, tk.END))
            added_count = 0
            for filename in filenames:
                abs_path = str(Path(filename).resolve())
                if abs_path not in current_items:
                    self.input_listbox.insert(tk.END, abs_path)
                    current_items.add(abs_path)
                    added_count += 1
                    logger.info(f"入力リストにファイル追加: {abs_path}")
            self.set_status(f"{added_count} 個のファイルを入力リストに追加しました。")

    def add_folder(self):
        """フォルダ選択ダイアログを開き、選択されたフォルダをリストに追加"""
        directory = filedialog.askdirectory(title="変換対象フォルダを選択")
        if directory:
            abs_path = str(Path(directory).resolve())
            current_items = set(self.input_listbox.get(0, tk.END))
            if abs_path not in current_items:
                self.input_listbox.insert(tk.END, abs_path)
                logger.info(f"入力リストにフォルダ追加: {abs_path}")
                self.set_status(f"フォルダを入力リストに追加しました: {abs_path}")
            else:
                self.set_status(f"フォルダは既に入力リストに存在します: {abs_path}")

    def remove_selected(self):
        """リストボックスで選択されている項目を削除"""
        selected_indices = self.input_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("情報", "削除する項目が選択されていません。")
            return

        # 後ろから削除しないとインデックスがずれる
        removed_count = 0
        for i in reversed(selected_indices):
            item_path = self.input_listbox.get(i)
            self.input_listbox.delete(i)
            logger.info(f"入力リストから削除: {item_path}")
            removed_count += 1
        self.set_status(f"{removed_count} 個の項目を入力リストから削除しました。")

    def clear_list(self):
        """リストボックスのすべての項目を削除"""
        if self.input_listbox.size() > 0:
            if messagebox.askyesno("確認", "入力リストを空にしますか？"):
                self.input_listbox.delete(0, tk.END)
                logger.info("入力リストをクリアしました。")
                self.set_status("入力リストをクリアしました。")
        else:
            messagebox.showinfo("情報", "入力リストは既に空です。")

    # --- その他UIメソッド --- 
    def browse_output_dir(self):
        """出力フォルダを選択する"""
        directory = filedialog.askdirectory(initialdir=self.output_dir_entry.get(), title="出力先フォルダを選択")
        if directory:
            abs_dir = str(Path(directory).resolve()) # 絶対パスに変換
            self.output_dir_entry.set(abs_dir)
            logger.info(f"出力フォルダ変更: {abs_dir}")

    def add_log(self, message, level="INFO"):
        """ログエリアにメッセージを追加 (キュー経由で)"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp} {level}] {message}\n"
        self.log_queue.put(log_entry)
        # loggerにも記録 (ファイルログ用)
        if level == "ERROR":
            logger.error(message)
        elif level == "WARNING":
            logger.warning(message)
        else:
            logger.info(message)

    def update_log_display(self, message):
        """ログテキストエリアを更新"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END) # 自動スクロール
        self.log_text.config(state=tk.DISABLED)

    def set_status(self, message):
        """ステータスバーを更新 (キュー経由で)"""
        self.status_queue.put(message)

    def check_queues(self):
        """キューをチェックしてUIを更新"""
        # ステータスキュー
        try:
            while not self.status_queue.empty():
                status_message = self.status_queue.get_nowait()
                self.status_var.set(status_message)
        except queue.Empty:
            pass
        except Exception as e:
            logger.error(f"ステータスキュー処理エラー: {e}")

        # ログキュー
        try:
            while not self.log_queue.empty():
                log_message = self.log_queue.get_nowait()
                self.update_log_display(log_message)
        except queue.Empty:
            pass
        except Exception as e:
            logger.error(f"ログキュー処理エラー: {e}")

        # 再スケジュール
        self.root.after(100, self.check_queues)

    def run_conversion(self):
        """変換処理を開始"""
        input_paths_str = self.input_listbox.get(0, tk.END)
        output_dir_str = self.output_dir_entry.get().strip()

        if not input_paths_str:
            messagebox.showerror("エラー", "変換対象のファイルまたはフォルダをリストに追加してください。")
            return
        if not output_dir_str:
            messagebox.showerror("エラー", "出力フォルダを指定してください。")
            return

        # 入力パスをPathオブジェクトのリストに (絶対パスになっているはず)
        input_paths = [Path(p) for p in input_paths_str]
        output_dir = Path(output_dir_str).resolve() # 常に絶対パスで扱う

        # 出力フォルダがなければ作成試行
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("エラー", f"出力フォルダの作成に失敗しました:\n{output_dir}\nエラー: {e}")
            logger.error(f"出力フォルダ作成失敗: {output_dir}, エラー: {e}")
            return

        # ログエリアをクリア
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        self.set_status("変換処理を開始します...")
        self.add_log(f"入力パス ({len(input_paths)}件):")
        for p in input_paths:
            self.add_log(f"  - {p}") # フルパス表示
        self.add_log(f"出力フォルダ: {output_dir}") # フルパス表示
        self.add_log(f"プラットフォーム: {platform.system()}")

        # ボタンを無効化
        self.convert_button.configure(state='disabled')
        self.add_file_button.configure(state='disabled')
        self.add_folder_button.configure(state='disabled')
        self.remove_button.configure(state='disabled')
        self.clear_button.configure(state='disabled')

        # スレッドで処理を実行 (引数を変更)
        threading.Thread(target=self.conversion_thread, args=(input_paths, output_dir), daemon=True).start()

    def conversion_thread(self, input_paths: list[Path], output_dir: Path):
        """ファイル検索と変換を実行するスレッド (複数入力パス対応)"""
        success_count = 0
        fail_count = 0
        skipped_count = 0
        target_files = set() # 処理対象ファイル (絶対パス) を格納

        try:
            self.set_status("入力パスを処理中...")

            # 入力パスリストから処理対象ファイルを収集
            for p in input_paths:
                abs_p = p.resolve() # 念のため再度 resolve
                if not abs_p.exists():
                    self.add_log(f"スキップ: パスが存在しません: {abs_p}", "WARNING")
                    continue

                if abs_p.is_file():
                    if abs_p.suffix.lower() in ['.ppt', '.pptx']:
                        target_files.add(abs_p)
                        self.add_log(f"ファイル発見: {abs_p}", "DEBUG")
                    else:
                        self.add_log(f"スキップ: PPT/PPTXファイルではありません: {abs_p}", "WARNING")
                elif abs_p.is_dir():
                    self.add_log(f"フォルダ検索中: {abs_p}", "INFO")
                    found_in_dir = 0
                    for ext in ['*.ppt', '*.PPT', '*.pptx', '*.PPTX']:
                        for found_file in abs_p.rglob(ext):
                            resolved_found_file = found_file.resolve()
                            target_files.add(resolved_found_file)
                            self.add_log(f"ファイル発見 (フォルダ内): {resolved_found_file}", "DEBUG")
                            found_in_dir += 1
                    self.add_log(f"{found_in_dir} 件のファイルを {abs_p} から追加しました。", "INFO")
                else:
                    self.add_log(f"スキップ: 不明なパスタイプ: {abs_p}", "WARNING")

            # 変換処理開始
            ppt_files_list = sorted(list(target_files)) # ソートして処理順を安定させる
            total_files = len(ppt_files_list)

            if total_files == 0:
                self.add_log("変換対象のPPT/PPTXファイルが見つかりませんでした。")
                self.set_status("完了 (対象ファイルなし)")
                return # スレッド終了

            self.add_log(f"{total_files} 件のPPT/PPTXファイルを変換対象とします。変換を開始します...")

            # プラットフォーム固有の変換関数を選択 (変更なし)
            converter_func = None
            os_name = platform.system()
            if os_name == "Windows":
                if not HAS_PYWIN32:
                     self.add_log("Windows環境ですが、pywin32が見つからないため変換できません。", "ERROR")
                     self.set_status("エラー (pywin32なし)")
                     fail_count = total_files
                     return
                converter_func = convert_ppt_to_pdf_windows
            elif os_name == "Darwin": # macOS
                converter_func = convert_ppt_to_pdf_macos
            else:
                self.add_log(f"未対応のOSです: {os_name}。変換できません。", "ERROR")
                self.set_status("エラー (未対応OS)")
                fail_count = total_files
                return

            # 変換処理ループ (変更なし、ただし ppt_file は絶対パス)
            for i, ppt_file_abs in enumerate(ppt_files_list):
                pdf_filename = ppt_file_abs.stem + ".pdf"
                pdf_file = output_dir / pdf_filename
                pdf_file_abs = pdf_file.resolve()
                # processed_files リストはUI版では不要かも

                self.set_status(f"変換中 ({i+1}/{total_files}): {ppt_file_abs.name}")
                self.add_log(f"変換開始: {ppt_file_abs}")

                if pdf_file_abs.exists():
                    self.add_log(f"スキップ: 出力先に同名ファイルが存在します -> {pdf_file_abs}", "WARNING")
                    skipped_count += 1
                    continue

                try:
                    start_time = time.time()
                    # 変換関数には Path オブジェクトを渡す
                    if converter_func(ppt_file_abs, pdf_file):
                        elapsed_time = time.time() - start_time
                        self.add_log(f"変換成功: {ppt_file_abs} -> {pdf_file_abs} ({elapsed_time:.2f}秒)")
                        success_count += 1
                    else:
                        self.add_log(f"変換失敗: {ppt_file_abs}", "ERROR")
                        fail_count += 1
                except (ImportError, RuntimeError, FileNotFoundError, subprocess.TimeoutExpired) as e:
                    self.add_log(f"変換エラー ({ppt_file_abs}): {e}", "ERROR")
                    fail_count += 1
                except Exception as e:
                    self.add_log(f"予期せぬ変換エラー ({ppt_file_abs}): {e}", "ERROR")
                    logger.error(f"予期せぬ変換エラー詳細 ({ppt_file_abs}):", exc_info=True)
                    fail_count += 1

        except Exception as e:
            # スレッド全体の予期せぬエラー
            self.add_log(f"処理中に予期せぬエラーが発生しました: {e}", "ERROR")
            logger.error("変換スレッド全体でエラー:", exc_info=True)
            self.set_status("エラー発生")
        finally:
            # 完了メッセージ
            completion_message = f"完了: 成功 {success_count}件, 失敗 {fail_count}件, スキップ {skipped_count}件"
            self.add_log(completion_message)
            self.set_status(completion_message)
            # ボタンを再度有効化
            self.convert_button.configure(state='normal')
            self.add_file_button.configure(state='normal')
            self.add_folder_button.configure(state='normal')
            self.remove_button.configure(state='normal')
            self.clear_button.configure(state='normal')


# --- メイン実行部 ---
def main():
    """アプリケーションのメインエントリポイント"""
    root = tk.Tk()
    # スタイル設定（MacでのComboboxなどの表示改善のため）
    style = ttk.Style(root)
    # OSによってテーマを設定
    if platform.system() == "Windows":
        style.theme_use('vista') # または 'xpnative', 'clam' など
    elif platform.system() == "Darwin":
        try:
            style.theme_use('aqua') # macOSのネイティブテーマ
        except tk.TclError:
            style.theme_use('default') # フォールバック
    else:
        style.theme_use('clam') # Linuxなどでのデフォルト候補

    app = ModernPPTtoPDFConverterUI(root)
    root.mainloop()
    logger.info("PPT to PDF Converter (UI) 終了")

if __name__ == "__main__":
    main()
