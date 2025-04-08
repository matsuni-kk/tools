# PPT/PPTX to PDF Converter

指定されたフォルダ内にある PowerPoint ファイル (.ppt, .pptx) を再帰的に検索し、PDF ファイルに変換するツールです。

このツールには、グラフィカルユーザーインターフェース (GUI) 版とコマンドラインインターフェース (CLI) 版の2種類があります。

## 機能

*   指定した入力フォルダ内のすべての .ppt および .pptx ファイルを検索します（サブフォルダ含む）。
*   見つかったファイルを PDF 形式に変換します。
*   変換された PDF ファイルを指定した出力フォルダに保存します（入力フォルダの階層構造は維持されません）。
*   Windows と macOS の両プラットフォームに対応しています。
*   GUI版: モダンなグラフィカルユーザーインターフェースを提供します。
*   GUI版: 処理ログをウィンドウ内に表示します。
*   CLI版: コマンドライン引数で操作します。オプションで詳細ログ出力や上書き指定が可能です。
*   どちらのバージョンも、処理ログを出力フォルダ内の `ppt_to_pdf_converter.log` ファイルに記録します。

## 使い方

### GUI版 (`ppt_to_pdf_converter_ui.py`)

1.  Python スクリプト `ppt_to_pdf_converter_ui.py` を実行します。
    ```bash
    python ppt_to_pdf_converter_ui.py
    ```
2.  GUI が表示されたら、以下の手順に従います。
    *   **PPT/PPTXのあるフォルダ**: [参照] ボタンをクリックして、変換したい PowerPoint ファイルが含まれるフォルダを選択します。
    *   **PDF出力先フォルダ**: [参照] ボタンをクリックして、変換された PDF ファイルを保存するフォルダを選択します。（デフォルトはスクリプトと同じ階層の `output` フォルダです）
    *   **変換開始**: ボタンをクリックして変換処理を開始します。
3.  処理の進捗状況と結果は、ウィンドウ下部のログエリアとステータスバーに表示されます。
4.  変換された PDF ファイルは、指定した出力フォルダに保存されます。

### CLI版 (`ppt_to_pdf_converter.py`)

1.  コマンドプロンプトまたはターミナルを開き、以下の形式でスクリプトを実行します。
    ```bash
    python ppt_to_pdf_converter.py <入力パス...> <出力フォルダパス> [オプション]
    ```
    *   `<入力パス...>`: 変換したい .ppt/.pptx **ファイル**または**フォルダ**のパスを**1つ以上**指定します。フォルダを指定した場合、そのフォルダ内のすべてのサブフォルダも再帰的に検索されます。
    *   `<出力フォルダパス>`: 変換後の PDF ファイルを保存するフォルダのパスを指定します。フォルダが存在しない場合は自動的に作成されます。

2.  **オプション:**
    *   `--overwrite`: このフラグを付けると、出力フォルダに同名の PDF ファイルが存在する場合に上書きします。デフォルトでは上書きしません。
    *   `-v` または `--verbose`: このフラグを付けると、より詳細なデバッグログをコンソールに出力します。

3.  **実行例:**
    ```bash
    # フォルダを指定 (サブフォルダも対象)
    python ppt_to_pdf_converter.py "D:\\プレゼン資料" "D:\\PDF出力"

    # 個別のファイルを複数指定
    python ppt_to_pdf_converter.py file1.pptx "folder/file2.ppt" /path/to/file3.pptx ./pdf_output

    # フォルダとファイルを混在して指定し、上書きを有効にする
    python ppt_to_pdf_converter.py /Users/user/Documents/Slides presentation.pptx /Users/user/Documents/PDFs --overwrite

    # 詳細ログを出力する
    python ppt_to_pdf_converter.py C:\\Input C:\\Output -v
    ```
4.  処理の進捗と結果はコンソールに表示されます。ログはフルパスで表示されます。

## 依存関係

*   **Python 3.6+**
*   **Tkinter**: GUI版 (`ppt_to_pdf_converter_ui.py`) を使用する場合に必要です。（通常 Python に同梱）
*   **Windows のみ**:
    *   **Microsoft PowerPoint**: PDF への変換に使用するため、インストールされている必要があります。
    *   **pywin32**: PowerPoint をプログラムから操作するために必要です。インストールされていない場合は、以下のコマンドでインストールしてください。
        ```bash
        pip install pywin32
        ```
*   **macOS のみ**:
    *   **Microsoft PowerPoint for Mac** または **Keynote**: PDF への変換に使用するため、どちらかがインストールされている必要があります。（PowerPoint が優先されます）

## 注意事項

*   変換処理には時間がかかる場合があります（特にファイル数が多い場合やファイルサイズが大きい場合）。
*   Windows 環境では、PowerPoint アプリケーションがバックグラウンドで起動・操作されます。
*   macOS 環境では、PowerPoint または Keynote アプリケーションが前面に出て操作される場合があります。
*   デフォルトでは、出力フォルダに既に同名の PDF ファイルが存在する場合、そのファイルの変換はスキップされます（CLI版で `--overwrite` を指定した場合を除く）。
*   パスワードで保護されたファイルや破損したファイルの変換は失敗する可能性があります。
*   変換の品質は、使用される PowerPoint / Keynote のバージョンやファイルの内容に依存する場合があります。
*   ログファイル (`ppt_to_pdf_converter.log`) は、ツールを起動するたびに追記されます（UI版、CLI版共通）。
