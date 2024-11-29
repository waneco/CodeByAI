import os
import sys
import datetime

def remove_leading_spaces(file_path):
    """
    テキストファイルの各行の先頭にある半角スペースを削除し、
    タイムスタンプ付きの新しいファイルに保存する。
    """
    try:
        # 入力ファイルのパスを確認
        if not os.path.isfile(file_path):
            print(f"エラー: 指定されたファイルが存在しません: {file_path}")
            return

        # 入力ファイル名とディレクトリを取得
        dir_name, file_name = os.path.split(file_path)
        base_name, ext = os.path.splitext(file_name)

        # タイムスタンプ付きの新しいファイル名を生成
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        new_file_name = f"{base_name}_{timestamp}{ext}"
        new_file_path = os.path.join(dir_name, new_file_name)

        # ファイルを処理
        with open(file_path, "r", encoding="utf-8") as f_in:
            lines = f_in.readlines()

        # 行ごとに先頭の半角スペースを削除
        stripped_lines = [line.lstrip() for line in lines]

        # 新しいファイルに書き込む
        with open(new_file_path, "w", encoding="utf-8") as f_out:
            f_out.writelines(stripped_lines)

        print(f"処理完了: 新しいファイルが作成されました: {new_file_path}")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    # ドラッグ＆ドロップされたファイルを取得
    if len(sys.argv) < 2:
        print("使用方法: テキストファイルをこのスクリプトにドラッグ＆ドロップしてください。")
        sys.exit(1)

    for file_path in sys.argv[1:]:
        print(f"処理中: {file_path}")
        remove_leading_spaces(file_path)

    print("すべてのファイルの処理が完了しました。")
