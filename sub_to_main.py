import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import shutil
from datetime import datetime
import threading
import os
import sys
import json
import pyexiv2
from PIL import Image
from ttkthemes import ThemedStyle
import winreg

class CopyImagesApp:
    def __init__(self, app):
        self.app = app
        app.title("記念サブ＝＞メイン")
        app.geometry("450x800")
        
        # レジストリキー設定
        self.registry_key = r"SOFTWARE\PhotoCopyTool"
        
        # アイコンを設定（タスクバー表示用）
        try:
            # 実行ファイルと同じディレクトリのアイコンを探す
            if getattr(sys, 'frozen', False):
                # PyInstallerでパッケージ化されている場合
                base_path = sys._MEIPASS
            else:
                # 通常のPythonスクリプトの場合
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_path, "icon.ico")
            if os.path.exists(icon_path):
                app.iconbitmap(icon_path)
            elif os.path.exists("icon.ico"):
                app.iconbitmap("icon.ico")
        except:
            pass

        self.copy_in_progress = False
        self.source_folder1_path = ""
        self.source_folder2_path = ""
        self.destination_folder_path = ""
        self.setup_gui()

        # ThemedStyleを使用してテーマを適用
        self.style = ThemedStyle(self.app)
        self.style.set_theme("yaru")  # テーマを選択（例: "plastik"）
        
        saved_settings = self.load_settings_from_registry()
        if saved_settings:
            self.source_folder1_path = saved_settings.get("source_folder1_path", "")
            self.source_folder2_path = saved_settings.get("source_folder2_path", "")
            self.destination_folder_path = saved_settings.get("destination_folder_path", "")
            self.update_folder_labels()


    def setup_gui(self):
        font_settings = ("Meiryo", 12)
        self.app.option_add("*TButton*Font", font_settings)
        self.app.option_add("*TLabel*Font", font_settings)
        self.app.option_add("*TEntry*Font", font_settings)
        self.app.option_add("*TCombobox*Font", font_settings)

        folder_number_label = ttk.Label(self.app, text="撮影No.を4桁で入力", font=("Meiryo", 16))
        folder_number_label.grid(row=0, column=0, columnspan=2, sticky="n")

        folder_number_label = ttk.Label(self.app, text="※複数ある場合は,（カンマ）で区切ってください", font=("Meiryo", 10))
        folder_number_label.grid(row=1, column=0, columnspan=2, sticky="n")

        self.folder_number_entry = ttk.Entry(self.app, font=("Meiryo", 14), width=15)
        self.folder_number_entry.grid(row=2, column=0, columnspan=2, sticky="ew")

        spacer_above = ttk.Label(self.app, text="", font=("Meiryo", 12))
        spacer_above.grid(row=2, column=0, pady=30)

        folder_number_label = ttk.Label(self.app, text="クリックは1回のみ", font=("Meiryo", 10))
        folder_number_label.grid(row=3, column=0, columnspan=2, sticky="n")

        frame = ttk.Frame(self.app)
        frame.grid(row=4, column=0, columnspan=2, sticky="n", pady=5)
        self.app.grid_columnconfigure(0, weight=1)  # 列0を拡張して中央に配置

        copy_button = ttk.Button(frame, text="写真をコピー", command=self.copy_images_parallel, width=15)
        copy_button.grid(row=0, column=0, sticky="ew")

        spacer = ttk.Label(frame, text="", font=("Arial", 12), width=2)
        spacer.grid(row=0, column=1)  # ボタン間にスキマを挿入

        cancel_button = ttk.Button(frame, text="作業中止", command=self.cancel_copy, width=15)
        cancel_button.grid(row=0, column=2, sticky="ew")

        self.app.grid_columnconfigure(1, weight=1)  # 列1も拡張して中央に配置

        self.notifications_frame = ttk.LabelFrame(self.app, text="通知")
        self.notifications_frame.grid(row=9, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")

        
        self.scrollbar = ttk.Scrollbar(self.notifications_frame, orient=tk.VERTICAL)
        self.notification_listbox = tk.Listbox(self.notifications_frame, yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.notification_listbox.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.notification_listbox.pack(fill=tk.BOTH, expand=True)

        spacer_above = ttk.Label(self.app, text="", font=("Meiryo", 12))
        spacer_above.grid(row=4, column=0, pady=20)

        spacer_folders = ttk.Label(self.app, text="", font=("Meiryo", 8))
        spacer_folders.grid(row=5, column=0, pady=10)

        folder_instruction_label = ttk.Label(self.app, text="対象の年度フォルダを選択してください。", font=("Meiryo", 12), foreground="blue")
        folder_instruction_label.grid(row=6, column=0, columnspan=2, sticky="n")

        folder_frame = ttk.LabelFrame(self.app, text="フォルダ設定")
        folder_frame.grid(row=7, column=0, columnspan=2, pady=5, padx=10, sticky="ew")

        source1_button = ttk.Button(folder_frame, text="セレクトPC側", command=self.select_source_folder1, width=12)
        source1_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.source1_label = ttk.Label(folder_frame, text="セレクトPC側: 未選択", font=("Meiryo", 9))
        self.source1_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        source2_button = ttk.Button(folder_frame, text="記念スタジオ側", command=self.select_source_folder2, width=12)
        source2_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.source2_label = ttk.Label(folder_frame, text="記念スタジオ側: 未選択", font=("Meiryo", 9))
        self.source2_label.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        dest_button = ttk.Button(folder_frame, text="メインPC", command=self.select_destination_folder, width=12)
        dest_button.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
        self.destination_label = ttk.Label(folder_frame, text="メインPC: 未選択", font=("Meiryo", 9))
        self.destination_label.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        folder_frame.grid_columnconfigure(1, weight=1)

        info_frame = tk.Frame(self.app, bg="yellow")
        info_frame.grid(row=10, column=0, columnspan=2, sticky="ew", pady=5, padx=10)
        info_label = tk.Label(info_frame, text="セレクトで星５のみを抽出してコピーしています。\nコピー完了後は必ず伝票と目視で\n記念メインを確認してください。", font=("Meiryo", 10), justify="center", foreground="black", background="yellow")
        info_label.pack(expand=True)

    def add_notification(self, message):
        timestamp = datetime.now().strftime("[%H:%M:%S] ")
        message = timestamp + message
        self.notification_listbox.insert(tk.END, message)
        self.notification_listbox.see(tk.END)

    # コピー処理の中に進捗情報を表示するための関数を追加
    def update_copy_progress(self, folder_number, image_filename):
        message = f"撮影No. {folder_number}: 写真のコピーが完了しました.\nコピー済み: {image_filename}"
        self.add_notification(message)

    def cancel_copy(self):
        if self.copy_in_progress:
            self.add_notification("コピー処理を中止します.")
            self.copy_in_progress = False
        else:
            self.add_notification("コピー処理は実行されていません.")

    def find_files_recursively(self, base_folder):
        all_files = []

        try:
            for root, dirs, files in os.walk(base_folder):
                  # .DS_Store ファイルを除外
                dirs[:] = [d for d in dirs if not d.startswith('.')]
                files = [f for f in files if not f.startswith('.DS_Store')]

                for file in files:
                    file_path = os.path.join(root, file)
                    all_files.append(file_path)
        except OSError as e:
            error_msg = f"フォルダアクセスエラー {base_folder}: {str(e)} (errno: {getattr(e, 'errno', 'unknown')})"
            self.add_notification(error_msg)
            print(f"os.walk error: {e}")

        return all_files
    

    def find_folder_with_number_recursive_common(self,base_folder, target_number):
        for root, dirs, files in os.walk(base_folder):
            # .DS_Store ファイルを除外
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            files = [f for f in files if not f.startswith('.DS_Store')]

            for dir_name in dirs:
                full_dir_path = os.path.join(root, dir_name)
                # 4桁の数字で始まり、その後にハイフンが含まれていない条件
                if dir_name.startswith(target_number) and '-' not in dir_name:
                    image_files = self.find_image_files_in_folder(full_dir_path)
                    if image_files:
                        # 画像ファイルが存在する場合はそのまま返す
                        return full_dir_path
                # 入れ子のフォルダがある場合、再帰的に探索を行う
                found_path = self.find_folder_with_number_recursive_common(full_dir_path, target_number)
                if found_path:
                    return found_path
        return None


    def find_image_files_in_folder(self, folder):
        image_files = []
        for root, dirs, files in os.walk(folder):
            # .DS_Store ファイルを除外
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            files = [f for f in files if not f.startswith('.DS_Store')]

            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg')):
                    image_files.append(os.path.join(root, file))
        return image_files

    def get_xmp_rating(self, image_path):
        import tempfile
        import shutil
        
        try:
            # pyexiv2のログレベルを設定してSony1警告を抑制
            pyexiv2.set_log_level(4)  # エラーレベルのみ
            
            # 日本語パスの問題を解決するため一時ファイルにコピー
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_file:
                temp_path = temp_file.name
            
            # 元ファイルを一時ファイルにコピー
            shutil.copy2(image_path, temp_path)
            
            try:
                with pyexiv2.Image(temp_path) as set_image:
                    xmp = set_image.read_xmp()
                    if 'Xmp.xmp.Rating' in xmp:
                        rating = xmp['Xmp.xmp.Rating']
                        return rating
                    else:
                        # レーティングが見つからない場合は0を返す
                        return "0"
            finally:
                # 一時ファイルを削除
                try:
                    os.unlink(temp_path)
                except:
                    pass
                    
        except Exception as e:
            error_message = str(e)
            self.add_notification(f"Error while extracting XMP Rating from {image_path}: {error_message}")
        return "0"




     # コピー処理の中で進捗情報を表示するたびに呼び出す
    def copy_images(self, folder_numbers, source_folder1_path, source_folder2_path, destination_folder_path):
        for folder_number in folder_numbers:
            source_folder1 = self.find_folder_with_number_recursive_common(source_folder1_path, folder_number)
            source_folder2 = self.find_folder_with_number_recursive_common(source_folder2_path, folder_number)
            
            if source_folder1 is None or source_folder2 is None:
                self.add_notification(f"警告: 撮影No.{folder_number}のフォルダが見つかりませんでした.")
                continue

            self.add_notification(f"撮影No.{folder_number}の処理を開始します...")
            
            # 全ファイルを取得して進捗計算用にカウント
            all_select_files = self.find_files_recursively(source_folder1)
            total_files = len(all_select_files)
            processed_files = 0
            rating5_found = 0
            copied_images = 0

            self.add_notification(f"セレクト側ファイル数: {total_files}枚")

            for image_path in all_select_files:
                if not self.copy_in_progress:
                    self.add_notification("コピー処理が中止されました.")
                    return
                
                processed_files += 1
                self.add_notification(f"進捗: {processed_files}/{total_files} - {os.path.basename(image_path)}")
                
                rating1 = self.get_xmp_rating(image_path)
                if rating1 == "5":
                    rating5_found += 1
                    image_filename1 = os.path.basename(image_path)
                    corresponding_image = os.path.join(source_folder2, image_filename1)
                    corresponding_image = os.path.normpath(corresponding_image)
                    
                    if os.path.exists(corresponding_image):
                        # 一時フォルダ内にコピー先フォルダを作成（Studio側のフォルダ名を使用）
                        temp_folder = os.path.join(self.get_temp_folder(), os.path.basename(source_folder2))
                        if not os.path.exists(temp_folder):
                            os.makedirs(temp_folder, exist_ok=True)

                        destination_path = os.path.join(temp_folder, image_filename1)
                        shutil.copyfile(corresponding_image, destination_path)
                        copied_images += 1
                        self.add_notification(f"★コピー完了: {image_filename1}")
                    else:
                        self.add_notification(f"対応する画像が存在しません: {image_filename1}")

            self.move_temp_folders(destination_folder_path)
            self.add_notification(f"撮影No.{folder_number}完了 - レーティング5: {rating5_found}枚, コピー: {copied_images}枚")

    def get_temp_folder(self):
        # アプリケーションの実行ディレクトリを取得
        app_directory = os.path.dirname(__file__)

        # 一時フォルダを返す
        return os.path.join(app_directory, "temp_copy_folder")

    def move_temp_folders(self, destination_folder_path):
        # 一時フォルダからコピー先フォルダへ移動
        temp_folder = self.get_temp_folder()
        
        # 一時フォルダが存在しない場合は何もしない
        if not os.path.exists(temp_folder):
            return
            
        for temp_subfolder in os.listdir(temp_folder):
            temp_subfolder_path = os.path.join(temp_folder, temp_subfolder)
            if os.path.exists(temp_subfolder_path) and os.path.isdir(temp_subfolder_path):
                destination_subfolder = os.path.join(destination_folder_path, temp_subfolder)
                os.makedirs(destination_subfolder, exist_ok=True)
                for file in os.listdir(temp_subfolder_path):
                    file_path = os.path.join(temp_subfolder_path, file)
                    if os.path.isfile(file_path):
                        destination_path = os.path.join(destination_subfolder, file)
                        shutil.move(file_path, destination_path)
                os.rmdir(temp_subfolder_path)  # 一時フォルダを削除



    def copy_images_parallel(self):
        if self.copy_in_progress:
            self.add_notification("エラー: 既にコピー処理が進行中です。")
            return

        if not self.source_folder1_path or not self.source_folder2_path or not self.destination_folder_path:
            self.add_notification("エラー: すべてのフォルダを選択してください。")
            return

        folder_numbers_input = self.folder_number_entry.get()
        folder_numbers = [number.strip() for number in folder_numbers_input.split(",")]

        self.copy_in_progress = True
        self.add_notification("コピー処理を開始します...")

        def copy_images_thread():
            try:
                print(f"destination_folder_path: {self.destination_folder_path}")
                print(f"Source folder 1 path: {self.source_folder1_path}")
                print(f"Source folder 2 path: {self.source_folder2_path}")
                
                # フォルダアクセス可能性をチェック
                self.add_notification("フォルダへのアクセスを確認中...")
                if not os.path.exists(self.source_folder1_path):
                    self.add_notification(f"エラー: セレクトPC側フォルダにアクセスできません: {self.source_folder1_path}")
                    return
                if not os.path.exists(self.source_folder2_path):
                    self.add_notification(f"エラー: 記念スタジオ側フォルダにアクセスできません: {self.source_folder2_path}")
                    return
                if not os.path.exists(self.destination_folder_path):
                    self.add_notification(f"エラー: メインPCフォルダにアクセスできません: {self.destination_folder_path}")
                    return
                
                self.copy_images(folder_numbers, self.source_folder1_path, self.source_folder2_path, self.destination_folder_path)
                self.add_notification("コピー処理が完了しました.")
            except Exception as e:
                error_msg = f"コピー処理中にエラーが発生しました: {str(e)} (errno: {getattr(e, 'errno', 'unknown')})"
                self.add_notification(error_msg)
                print(f"Error details: {e}")
            finally:
                self.copy_in_progress = False

        copy_thread = threading.Thread(target=copy_images_thread)
        copy_thread.start()

    def on_closing(self):
        self.save_settings_to_registry()
        self.app.destroy()

    def save_settings_to_registry(self):
        try:
            # HKEY_CURRENT_USERに設定を保存
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.registry_key)
            winreg.SetValueEx(key, "source_folder1_path", 0, winreg.REG_SZ, self.source_folder1_path)
            winreg.SetValueEx(key, "source_folder2_path", 0, winreg.REG_SZ, self.source_folder2_path)
            winreg.SetValueEx(key, "destination_folder_path", 0, winreg.REG_SZ, self.destination_folder_path)
            winreg.CloseKey(key)
        except Exception as e:
            print(f"レジストリ保存エラー: {e}")

    def load_settings_from_registry(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.registry_key)
            settings = {}
            settings["source_folder1_path"] = winreg.QueryValueEx(key, "source_folder1_path")[0]
            settings["source_folder2_path"] = winreg.QueryValueEx(key, "source_folder2_path")[0]
            settings["destination_folder_path"] = winreg.QueryValueEx(key, "destination_folder_path")[0]
            winreg.CloseKey(key)
            return settings
        except (FileNotFoundError, OSError, Exception):
            return {}

    def select_source_folder1(self):
        folder_path = filedialog.askdirectory(title="セレクトPC側フォルダを選択してください")
        if folder_path:
            self.source_folder1_path = folder_path
            self.source1_label.config(text=f"セレクトPC側: {os.path.basename(folder_path)}")
            self.save_settings_to_registry()

    def select_source_folder2(self):
        folder_path = filedialog.askdirectory(title="記念スタジオ側フォルダを選択してください")
        if folder_path:
            self.source_folder2_path = folder_path
            self.source2_label.config(text=f"記念スタジオ側: {os.path.basename(folder_path)}")
            self.save_settings_to_registry()

    def select_destination_folder(self):
        folder_path = filedialog.askdirectory(title="メインPCフォルダを選択してください")
        if folder_path:
            self.destination_folder_path = folder_path
            self.destination_label.config(text=f"メインPC: {os.path.basename(folder_path)}")
            self.save_settings_to_registry()

    def update_folder_labels(self):
        if hasattr(self, 'source1_label'):
            if self.source_folder1_path:
                self.source1_label.config(text=f"セレクトPC側: {os.path.basename(self.source_folder1_path)}")
            else:
                self.source1_label.config(text="セレクトPC側: 未選択")
        
        if hasattr(self, 'source2_label'):
            if self.source_folder2_path:
                self.source2_label.config(text=f"記念スタジオ側: {os.path.basename(self.source_folder2_path)}")
            else:
                self.source2_label.config(text="記念スタジオ側: 未選択")
        
        if hasattr(self, 'destination_label'):
            if self.destination_folder_path:
                self.destination_label.config(text=f"メインPC: {os.path.basename(self.destination_folder_path)}")
            else:
                self.destination_label.config(text="メインPC: 未選択")

if __name__ == "__main__":
    app = tk.Tk()
    copy_images_app = CopyImagesApp(app)
    app.protocol("WM_DELETE_WINDOW", copy_images_app.on_closing)
    app.mainloop()
