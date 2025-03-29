#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Context Folder Extracter (Command Line)

指定したフォルダパスとそのサブフォルダ内のファイル内容を抽出。
すべてのファイル内容を1つのMarkdownファイルに集約する。

使用例:
        python context_folder_extracter.py <target_folder_path> [options]

引数:
        target_folder_path: 内容を抽出するフォルダのパス (必須)

オプション:
        --days, -d DAYS: 過去何日以内のファイルを対象にするか (デフォルト: 14)
        --all, -a:     すべての期間のファイルを対象にする (--daysより優先)
        --exclude, -e [PATTERN ...]: 除外するフォルダやファイルのパターン
        --verbose, -v: 詳細なログを出力する

出力:
        output ディレクトリ内に、<フォルダ名>_<タイムスタンプ>_contents.md というファイル名で出力
"""

import os
import sys
import platform
import datetime
import time
import argparse
from pathlib import Path
import logging
import re

# output ディレクトリの存在確認または作成
# __file__ から絶対パスを取得し、その親ディレクトリを基準にする
try:
    # sys.argv[0]を使用してスクリプトの絶対パスを取得（MacとWindowsの両方で動作）
    script_dir = Path(os.path.abspath(sys.argv[0])).parent
    output_dir = script_dir / "output"
    output_dir.mkdir(exist_ok=True)
    log_file_path = output_dir / "context_folder_extracter.log"
except Exception as e:
    # outputディレクトリ作成やログファイルパス設定に失敗した場合
    # スクリプトと同じディレクトリにログを出力するなどのフォールバック処理
    script_dir = Path(__file__).resolve().parent
    log_file_path = script_dir / "context_folder_extracter.fallback.log"
    print(f"警告: outputディレクトリの準備に失敗しました。ログは {log_file_path} に出力されます。エラー: {e}", file=sys.stderr)
    # output_dir が未定義の場合に備え、スクリプトディレクトリを割り当てておく
    if 'output_dir' not in locals():
        output_dir = script_dir

# ロガーの設定
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, mode="w", encoding="utf-8"), # 上書きモードに設定
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

logger.info(f"スクリプトディレクトリ: {script_dir}")
logger.info(f"出力ディレクトリ: {output_dir}")
logger.info(f"ログファイルパス: {log_file_path}")

def get_bms_root():
    """
    BMSのルートディレクトリを取得する。環境に合わせて調整し、存在確認を行う。
    """
    system = platform.system()
    bms_root = None
    if system == "Windows":
        # Windowsの場合は D:/BMS を想定
        bms_root = Path("D:/BMS")
    elif system == "Darwin": # macOS の正式名は Darwin
        # macOSの場合はホームディレクトリ下の Desktop/BMS を想定
        home_dir = Path.home() # ホームディレクトリを取得
        bms_root = home_dir / "Desktop" / "BMS"
        logger.info(f"macOS用BMSルートディレクトリを設定: {bms_root}")
    else: # その他のLinuxなど
        # 必要であれば他のOS用のパスも定義
        logger.warning(f"未対応のOSです: {system}. デフォルトパスは設定されません。")
        return None # PathオブジェクトではなくNoneを返す

    # パスが存在するか確認
    if bms_root and not bms_root.exists():
        logger.warning(f"デフォルトのBMSルートディレクトリが見つかりません: {bms_root}")
        # 存在しない場合でもパス自体は返す

    return bms_root # Pathオブジェクトを返す

def is_binary_file(file_path):
    """
    ファイルがバイナリかテキストかを判定
    """
    textchars = bytearray({7, 8, 9, 10, 12, 13, 27} | set(range(0x20, 0x100)) - {0x7f})
    try:
        with open(file_path, 'rb') as f:
            is_binary = bool(f.read(1024).translate(None, textchars))
        return is_binary
    except Exception as e:
        logger.warning(f"ファイル {file_path} の判定中にエラー発生: {e}")
        return True  # エラーが発生した場合もバイナリとして扱う

def is_file_in_date_range(file_path, days=None):
    """
    ファイルが指定された日数以内かどうか確認
    daysがNoneの場合は、すべてのファイルが対象
    """
    if days is None:
        return True
    
    try:
        # ファイルの最終更新日時を取得
        file_mtime = os.path.getmtime(file_path)
        now = time.time()
        
        # 日数を秒に変換
        days_in_seconds = days * 24 * 60 * 60
        
        # 指定期間内かチェック
        return (now - file_mtime) <= days_in_seconds
    except Exception as e:
        logger.warning(f"ファイル日付の確認中にエラー発生: {file_path} - {e}")
        return True  # エラーが発生した場合は含める

def match_exclude_pattern(path, exclude_patterns):
    """
    指定された除外パターンにパスが一致するかどうか確認
    いずれかのパターンに一致すればTrueを返す
    """
    if not exclude_patterns:
        return False
    
    path_lower = str(path).lower()
    for pattern in exclude_patterns:
        if pattern.lower() in path_lower:
            logger.debug(f"除外パターン '{pattern}' がパス '{path}' に一致")
            return True
    return False

def read_file_content(file_path):
    """
    ファイルの内容を読む。エラーで対応
    バイナリファイルの場合は、[バイナリファイル]と表示
    """
    if is_binary_file(file_path):
        return f"[バイナリファイル: {file_path}]"
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(file_path, 'r', encoding='shift-jis') as f:
                return f.read()
        except Exception as e:
            logger.warning(f"ファイル {file_path} の読込中にエラー発生: {e}")
            return f"[ファイル読込エラー: {file_path}]"
    except Exception as e:
        logger.warning(f"ファイル {file_path} の読込中にエラー発生: {e}")
        return f"[ファイル読込エラー: {file_path}]"

def find_folders_and_files(target_folder_path, days=None, exclude_patterns=None):
    """
    指定されたフォルダパスとそのサブフォルダ内のファイルを検索
    days: 指定された日数以内のファイルのみを含める（Noneの場合はすべてのファイル）
    exclude_patterns: 除外するパターンのリスト
    """
    found_files = [] # ファイルパスのリスト
    
    # target_folder_pathがPathオブジェクトの場合は文字列に変換
    if isinstance(target_folder_path, Path):
        target_folder_path = str(target_folder_path)

    log_message = f"検索開始: フォルダパス '{target_folder_path}' 内を検索"
    logger.info(log_message)

    if days is not None:
        log_message = f"期間制限: 過去{days}日以内に更新されたファイルのみを対象"
        logger.info(log_message)

    if exclude_patterns:
        log_message = f"除外パターン: {', '.join(exclude_patterns)}"
        logger.info(log_message)

    # 指定されたフォルダパスとそのサブフォルダを検索
    for dirpath, dirnames, filenames in os.walk(target_folder_path):
        current_dir = Path(dirpath)

        # 除外パターンに一致するフォルダはスキップ (サブフォルダも含む)
        if match_exclude_pattern(current_dir, exclude_patterns):
            log_message = f"除外パターンに一致するフォルダをスキップ: {current_dir}"
            logger.info(log_message)
            # このディレクトリ配下を無視するために dirnames をクリア
            dirnames[:] = []
            continue

        # フォルダ内のファイルを処理
        for file in filenames:
            file_path = current_dir / file
            full_path = str(file_path)

            # 除外パターンに一致する場合はスキップ
            if match_exclude_pattern(file_path, exclude_patterns):
                logger.debug(f"除外パターンに一致するファイルをスキップ: {file_path}")
                continue

            # 期間チェック
            if is_file_in_date_range(full_path, days):
                found_files.append(full_path)
                logger.debug(f"ファイル追加: {full_path}")
            else:
                logger.debug(f"期間外ファイルスキップ: {full_path}")

    log_message = f"検索完了: {len(found_files)} 個のファイルを発見"
    logger.info(log_message)

    # ファイルが見つかった場合のみ、フォルダ情報を返す形式に合わせる
    if found_files:
        return [{"folder_path": str(target_folder_path), "files": found_files}]
    else:
        return []

def create_markdown_content(target_folder_path, found_folders_info): # 引数名変更
    """
    検索結果をMarkdown形式で出力 (指定フォルダパス基準)
    """
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    target_folder_name = Path(target_folder_path).name # フォルダ名を取得

    # 生成日時とファイル数はログにのみ記録
    logger.info(f"生成日時: {timestamp}")

    # found_folders_info はリストだが、要素は1つのはず
    if not found_folders_info:
         md_content = f"# {target_folder_name} フォルダ内容集約\n\n"
         md_content += f"**指定されたフォルダパス**: `{target_folder_path}`\n\n"
         md_content += "**期間内のファイルは見つかりませんでした。**\n"
         return md_content

    # 最初の要素（唯一のはず）からファイルリストを取得
    files = found_folders_info[0].get("files", [])
    file_count = len(files)

    logger.info(f"対象ファイル数: {file_count}")

    md_content = f"# {target_folder_name} フォルダ内容集約\n\n"
    md_content += f"**指定されたフォルダパス**: `{target_folder_path}`\n\n"
    md_content += f"**抽出されたファイル数**: {file_count}\n\n"

    if not files:
        md_content += "**期間内のファイルは見つかりませんでした。**\n"
        return md_content

    md_content += "## 目次\n\n"

    # ファイルを最終更新日時で降順ソート
    sorted_files = sorted(files, key=os.path.getmtime, reverse=True)

    # 目次作成 - ファイル名のみ表示
    for i, file_path in enumerate(sorted_files, 1):
        file_name = Path(file_path).name
        file_anchor = f"file-{i}-{file_name.lower().replace(' ', '-').replace('.', '-')}"
        md_content += f"{i}. [{file_name}](#{file_anchor})\n"

    md_content += "\n---\n\n"

    # ファイルごとの詳細 - 新しいものから順に
    md_content += "## ファイル内容\n\n"
    for i, file_path in enumerate(sorted_files, 1):
        file_extension = Path(file_path).suffix.lower()
        file_name = Path(file_path).name
        file_mtime = datetime.datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
        file_anchor = f"file-{i}-{file_name.lower().replace(' ', '-').replace('.', '-')}"

        # Markdownの見出しにアンカーを追加
        md_content += f"### {i}. {file_name} <a name='{file_anchor}'></a>\n\n"
        md_content += f"**パス**: `{file_path}`\n"
        md_content += f"**最終更新日時**: {file_mtime}\n\n"

        # ファイル内容をコードブロックで表示
        file_content = read_file_content(file_path)

        # ファイル拡張子に応じた言語指定 (主要なもののみ抜粋)
        lang = ""
        if file_extension in ['.py', '.pyw']:
            lang = "python"
        elif file_extension in ['.md', '.markdown']:
            lang = "markdown"
        elif file_extension in ['.js', '.jsx']:
            lang = "javascript"
        elif file_extension in ['.html', '.htm']:
            lang = "html"
        elif file_extension in ['.css']:
            lang = "css"
        elif file_extension in ['.json']:
            lang = "json"
        elif file_extension in ['.xml']:
            lang = "xml"
        elif file_extension in ['.sh', '.bash']:
            lang = "bash"
        elif file_extension in ['.bat', '.cmd', '.ps1']:
            lang = "powershell"
        elif file_extension in ['.sql']:
            lang = "sql"

        md_content += f"```{lang}\n{file_content}\n```\n\n"
        md_content += "---\n\n"

    return md_content

def main():
    """
    メイン処理
    """
    parser = argparse.ArgumentParser(description='指定したフォルダパス内のファイル内容を集約するツール')
    parser.add_argument('target_folder_path', help='内容を抽出するフォルダのパス') # 必須の位置引数に変更
    # parser.add_argument('root_dir', nargs='?', default=None, help='...') # 削除
    parser.add_argument('--verbose', '-v', action='store_true', help='詳細なログを出力する')
    parser.add_argument('--days', '-d', type=int, default=14, help='過去何日以内のファイルを対象にするか指定（デフォルト: 14日）')
    parser.add_argument('--all', '-a', action='store_true', help='すべての期間のファイルを対象にする（--daysよりも優先）')
    parser.add_argument('--exclude', '-e', nargs='+', help='除外するフォルダやファイルのパターンを指定（スペース区切りで複数指定可能）')

    args = parser.parse_args()

    # 詳細ログが有効な場合はDEBUGレベルに設定
    if args.verbose:
        logger.setLevel(logging.DEBUG)
        # ファイルハンドラもDEBUGレベルに設定
        for handler in logger.handlers:
            if isinstance(handler, logging.FileHandler):
                handler.setLevel(logging.DEBUG)

    # フォルダパスの取得と検証
    target_folder_path_str = args.target_folder_path
    target_folder_path = Path(target_folder_path_str)

    if not target_folder_path.exists():
        logger.error(f"指定されたフォルダパスが見つかりません: {target_folder_path}")
        print(f"エラー: 指定されたフォルダパスが見つかりません: {target_folder_path}")
        return 1
    if not target_folder_path.is_dir():
        logger.error(f"指定されたパスはフォルダではありません: {target_folder_path}")
        print(f"エラー: 指定されたパスはフォルダではありません: {target_folder_path}")
        return 1

    # 期間指定の処理
    days = None if args.all else args.days

    # 除外パターンの処理
    exclude_patterns = args.exclude if args.exclude else []

    # output_dir の存在確認と出力ファイルパスの設定
    global output_dir
    if not output_dir.exists():
        logger.warning(f"出力ディレクトリが見つかりません: {output_dir}. 作成を試みます。")
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as mkdir_e:
            logger.error(f"出力ディレクトリの作成に失敗しました: {output_dir}\nエラー: {mkdir_e}")
            print(f"エラー: 出力ディレクトリの作成に失敗しました: {output_dir}\nエラー: {mkdir_e}")
            return 1

    # 出力ファイル名に禁止文字が含まれていないか確認・置換
    folder_name = target_folder_path.name
    if not folder_name: # ルートディレクトリなどが指定された場合
        folder_name = "root_context"
    safe_folder_name = re.sub(r'[\\/*?:"<>|]', '_', folder_name)
    timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") # タイムスタンプ追加
    output_filename = f"{safe_folder_name}_{timestamp_str}_contents.md" # タイムスタンプもファイル名に含める
    output_file = output_dir / output_filename

    logger.info(f"処理開始: フォルダパス '{target_folder_path}' の内容抽出")
    if days is not None:
        logger.info(f"期間制限: 過去 {days} 日以内")
    if exclude_patterns:
        logger.info(f"除外パターン: {exclude_patterns}")
    logger.info(f"出力ファイル: {output_file}")

    try:
        # フォルダとファイルの検索 (関数名と引数を変更)
        logger.info(f"ターゲットフォルダの絶対パス: {target_folder_path.resolve()}")
        found_folders_info = find_folders_and_files(target_folder_path, days, exclude_patterns)

        # Markdown ファイルの作成 (引数を変更)
        md_content = create_markdown_content(target_folder_path, found_folders_info)

        # ファイルに書き出し
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(md_content)

        logger.info(f"処理完了: 出力ファイル '{output_file}' を作成しました")
        print(f"\n出力ファイル: {output_file}")

    except Exception as e:
        logger.error(f"エラーが発生しました: {e}", exc_info=True)
        print(f"\nエラーが発生しました: {e}")
        return 1

    return 0

if __name__ == "__main__":
    sys.exit(main())
