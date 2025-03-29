# Context Folder Extracter

指定したフォルダパスとそのサブフォルダ内のすべてのファイルを検索し、その内容（コンテキスト）を1つのマークダウンファイルにまとめるPythonツールです。

このツールはコマンドライン版とGUI版の二つのバージョンが用意されています。

- コマンドライン版: `context_folder_extracter.py`
- GUI版: `context_folder_extracter_ui.py`

## 機能

- 指定したフォルダパスとそのサブフォルダ内のすべてのファイルを検索
- すべてのファイル内容を1つのマークダウンファイルにまとめる（コンテキスト抽出）
- 自動的にファイル拡張子に応じた言語シンタックスハイライトを適用
- バイナリファイルは「バイナリファイル」と表示し、テキストファイルの内容は全文表示
- 目次やファイル一覧を含む見やすいドキュメント構造
- 指定した期間内のファイルのみを対象にする機能（デフォルト: 過去14日以内）
- 特定のフォルダやファイルを除外する機能
- WindowsとMacの両環境で動作

## 使い方

### コマンドライン版の使用方法

```bash
# Macの場合
cd /Users/matuni__/Desktop/BMS/tools/python_tools/context_folder_extracter
python context_folder_extracter.py <target_folder_path>

# Windowsの場合
cd D:\BMS\tools\python_tools\context_folder_extracter
python context_folder_extracter.py <target_folder_path>
```

### GUI版の使用方法

```bash
# Macの場合
cd /Users/matuni__/Desktop/BMS/tools/python_tools/context_folder_extracter
python context_folder_extracter_ui.py

# Windowsの場合
cd D:\BMS\tools\python_tools\context_folder_extracter
python context_folder_extracter_ui.py
```

### コマンドライン版の引数

- `target_folder_path`: 内容を抽出するフォルダのパス (必須)
- `--verbose`, `-v`: 詳細なログ出力を有効にするオプション
- `--days`, `-d`: 過去何日以内のファイルを対象にするか指定（デフォルト: 14日）
- `--all`, `-a`: すべての期間のファイルを対象にする（--daysよりも優先）
- `--exclude`, `-e`: 除外するフォルダやファイルのパターンを指定（スペース区切りで複数指定可能）

### コマンドライン版の使用例

```bash
# Macの場合、指定パスのフォルダを抽出
cd /Users/matuni__/Desktop/BMS/tools/python_tools/context_folder_extracter
python context_folder_extracter.py /Users/matuni__/Desktop/BMS/workspace/project1

# 期間を過去30日以内に変更して抽出する場合（Windows）
cd D:\BMS\tools\python_tools\context_folder_extracter
python context_folder_extracter.py D:\BMS\workspace\project1 --days 30

# すべての期間のファイルを対象にする場合（Windows）
cd D:\BMS\tools\python_tools\context_folder_extracter
python context_folder_extracter.py D:\BMS\workspace\project1 --all

# 特定のパターンを含むフォルダやファイルを除外する場合（Windows）
cd D:\BMS\tools\python_tools\context_folder_extracter
python context_folder_extracter.py D:\BMS\workspace\project1 --exclude temp backup test

# 詳細なログを出力する場合（Mac）
cd /Users/matuni__/Desktop/BMS/tools/python_tools/context_folder_extracter
python context_folder_extracter.py /Users/matuni__/Desktop/BMS/workspace/project1 --verbose

# 複数のオプションを組み合わせる場合（Windows）
cd D:\BMS\tools\python_tools\context_folder_extracter
python context_folder_extracter.py D:\BMS\workspace\project1 --all --exclude test temp --verbose
```

### GUI版の使い方

GUI版では、以下の操作が可能です：

1. フォルダパスを入力または「参照」ボタンで選択
2. 期間設定（日数を指定、または「すべての期間」をチェック）
3. 除外パターンを指定（スペース区切りで複数指定可能）
4. 「抽出開始」ボタンをクリック
5. 対象フォルダ内のファイルが抽出され、結果が表示されます
6. 「コピー」ボタンで結果をクリップボードにコピー、または「保存」ボタンでファイルに保存できます

大量のテキストを処理する場合は、「180K文字コピー」または「900K文字コピー」ボタンを使用して、サイズ制限のあるシステムに対応します。

## 出力

- 出力ファイルは `output` ディレクトリに `<フォルダ名>_<タイムスタンプ>_contents.md` という名前で保存されます
- 例：`project1_20250329_120000_contents.md`

## 注意事項

- MacとWindowsの両方で動作するように設計されています
- 非常に大きなファイルや多数のファイルがある場合、処理に時間がかかる場合があります
- バイナリファイルや読み込みに失敗したファイルは内容表示をスキップします
- すべてのフォルダとファイルが検索対象に含まれます（隠しフォルダや特殊なフォルダも含みます）
- デフォルトでは過去14日（2週間）以内に更新されたファイルのみが対象となります
- `--all` オプションを使用すると、すべての期間のファイルが対象になります
- 除外パターンは指定されたパス内に含まれる場合に一致します（部分一致）
- 除外パターンは大文字小文字を区別しません

## 出力ファイルの構成

出力されるマークダウンファイルは以下の構造になっています：

```
# [フォルダ名] フォルダ内容集約

**指定されたフォルダパス**: `パス`

**抽出されたファイル数**: [N]

## 目次

1. [ファイル名1](#file-1-ファイル名1-txt)
2. [ファイル名2](#file-2-ファイル名2-py)
...

---

## ファイル内容

### 1. ファイル名1

**パス**: `ファイルパス1`
**最終更新日時**: YYYY-MM-DD HH:MM:SS

```[言語]
ファイル内容
```

---

### 2. ファイル名2
...
```

## 要件

- Python 3.6以上
- 標準ライブラリのみ使用（追加のインストールは不要）