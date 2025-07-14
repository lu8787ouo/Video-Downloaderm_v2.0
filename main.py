'''
ver2.0
'''

from __future__ import annotations
import customtkinter as ctk
from customtkinter import CTkImage
from tkinter import filedialog, messagebox
from PIL import Image, ImageOps
import requests
import threading
import io
import os
from CTkTable import CTkTable
import pywinstyles
from logging_config import setup_logger, log_and_show_error
from Page1 import get_video_info, download_video_audio
from Page2 import parse_playlist, download_video_audio_playlist_with_retry
from Page3 import convert_video, convert_audio, get_media_duration, time_to_seconds
from Page4 import fetch_voice_names, convert_text_to_speech
from config_manager import load_config, save_config
from concurrent.futures import ThreadPoolExecutor, as_completed
import subprocess
import json
import asyncio

# ------------------------------
# 初始化 Logger
# ------------------------------
logger = setup_logger(__name__)

# ------------------------------
# 語言設定資料
# ------------------------------
LANGUAGES = {}
index_path = os.path.join(os.path.dirname(__file__), "locale/index.txt")

with open(index_path, "r", encoding="utf-8") as f:
    lang_list = [line.strip() for line in f.read().split(",") if line.strip()]

for lang in lang_list:
    lang_file = os.path.join(os.path.dirname(__file__), f"locale/{lang}.json")
    if os.path.exists(lang_file):
        with open(lang_file, "r", encoding="utf-8") as f:
            LANGUAGES[lang] = json.load(f)
    else:
        print(f"Warning: {lang_file} not found.")

LANGUAGE_OPTIONS = {
    "zh-TW": "繁體中文",
    "zh-CN": "简体中文",
    "en": "English",
    "ja": "日本語",
    "es": "Español",
}

# ------------------------------
# 字體設定資料
# ------------------------------

def get_font(lang, key):
    font_info = LANGUAGES[lang]["font"][key]
    family = font_info.get("family", "Arial")
    size = font_info.get("size", 14)
    weight = font_info.get("weight", "")
    return (family, size, weight) if weight else (family, size)

# ------------------------------
# 設定視窗
# ------------------------------
class Setting(ctk.CTkToplevel):
    def __init__(self, master: 'MainApp'):
        super().__init__(master)
        self.title(LANGUAGES[self.master.current_language]["setting"]["setting_title"])
        self.geometry("480x360")
        self.resizable(False, False)
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=6)
        self.grid_rowconfigure(2, weight=1)
        
        # text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"

        # Background image initialization
        self.bg_label = ctk.CTkLabel(self, text="")
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        # ====== 頁面頂部 ======
        self.frame_top = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"),
        )
        self.frame_top.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.frame_top.grid_rowconfigure(0, weight=1)
        self.frame_top.grid_columnconfigure(0, weight=1)

        self.tabs = [
            {"key":"theme",   "label": LANGUAGES[self.master.current_language]["setting"]["segmentedBtn"]["theme"]},
            {"key":"language", "label": LANGUAGES[self.master.current_language]["setting"]["segmentedBtn"]["language"]},
            {"key":"resolution", "label": LANGUAGES[self.master.current_language]["setting"]["segmentedBtn"]["resolution"]},
            {"key":"picture", "label": LANGUAGES[self.master.current_language]["setting"]["segmentedBtn"]["picture"]},
            {"key":"others", "label": LANGUAGES[self.master.current_language]["setting"]["segmentedBtn"]["others"]},
        ]
        # 抽出 labels，並做反查 map
        self.tab_labels   = [tab["label"] for tab in self.tabs]
        self.label_to_key = {tab["label"]: tab["key"] for tab in self.tabs}

        self.nav = ctk.CTkSegmentedButton(
            master=self.frame_top,
            values=self.tab_labels,
            command=self.select_tab
        )
        self.nav.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.nav.set(LANGUAGES[self.master.current_language]["setting"]["segmentedBtn"]["theme"])  # set initial value

        # 依 key 建 frame
        self.frames = {}
        for tab in self.tabs:
            f = ctk.CTkFrame(
                self,
                bg_color=("#FFFFFF", "#000001"),  # 明暗主題雙色
                fg_color=("#FFFFFF", "#000001"),  # 明暗主題雙色
                )
            f.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
            f.grid_rowconfigure((0,1,2,3), weight=1)
            f.grid_columnconfigure((0,1,2,3), weight=1)
            f.grid_remove()
            self.frames[tab["key"]] = f

        # ====== 主題設定頁面 ======
        # 主題設定
        self.theme_label = ctk.CTkLabel(self.frames["theme"])
        self.theme_label.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        self.theme_combobox = ctk.CTkComboBox(self.frames["theme"], values=["Light", "Dark"])
        self.theme_combobox.grid(row=0, column=3, padx=5, pady=10, sticky="ew")
        self.theme_combobox.set(self.master.config.get("theme", "Dark")) # 預設值

        self.theme_color_label = ctk.CTkLabel(self.frames["theme"])
        self.theme_color_label.grid(row=1, column=0, padx=5, pady=10, sticky="ew")
        self.theme_color_btn = ctk.CTkButton(self.frames["theme"], command=self.import_theme_json)
        self.theme_color_btn.grid(row=1, column=3, padx=5, pady=10, sticky="w")
        
        # ===== 語言設定頁面 ======
        # 語言設定
        self.language_label = ctk.CTkLabel(self.frames["language"])
        self.language_label.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        self.language_combobox = ctk.CTkComboBox(self.frames["language"], values=list(LANGUAGE_OPTIONS.values()))
        self.language_combobox.grid(row=0, column=3, padx=5, pady=10, sticky="ew")
        self.language_combobox.set(LANGUAGE_OPTIONS[self.master.config.get("language", "zh-TW")])
        
        # ====== 解析度設定頁面 ======
        # 視窗解析度設定
        self.resolution_label = ctk.CTkLabel(self.frames["resolution"])
        self.resolution_label.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        self.resolution_combobox = ctk.CTkComboBox(self.frames["resolution"], values=["1280x720"])
        self.resolution_combobox.grid(row=0, column=3, padx=5, pady=10, sticky="ew")
        self.resolution_combobox.set(self.master.config.get("resolution", "1280x720")) # 預設值
        
        # ===== 圖片設定頁面 ======
        # 廣告路徑設定
        self.ad_label = ctk.CTkLabel(self.frames["picture"])
        self.ad_label.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        self.ad_button = ctk.CTkButton(self.frames["picture"], command=self.import_ad_image)  # 若有locale可再補
        self.ad_button.grid(row=0, column=3, padx=5, pady=10, sticky="w")

        # 新增背景圖片設定
        self.background_label = ctk.CTkLabel(self.frames["picture"])
        self.background_label.grid(row=1, column=0, padx=5, pady=10, sticky="ew")
        self.bg_button = ctk.CTkButton(self.frames["picture"], command=self.import_bg_image)
        self.bg_button.grid(row=1, column=3, padx=5, pady=10, sticky="w")        

        # ====== 其他設定頁面 ======
        # 透明度設定
        self.transparency_label = ctk.CTkLabel(self.frames["others"])
        self.transparency_label.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        self.transparency_entry = ctk.CTkEntry(self.frames["others"])
        self.transparency_entry.insert(0, self.master.config.get("transparency", "1"))
        self.transparency_entry.grid(row=0, column=3, padx=5, pady=10, sticky="ew")

        # cookies 設定
        self.cookies_label = ctk.CTkLabel(self.frames["others"])
        self.cookies_label.grid(row=1, column=0, padx=5, pady=10, sticky="ew")
        self.cookies_reset_button = ctk.CTkButton(self.frames["others"], width=80, command=self.reset_cookies)
        self.cookies_reset_button.grid(row=1, column=3, padx=5, pady=10, sticky="e")
        self.cookies_button = ctk.CTkButton(self.frames["others"], width=80, command=self.import_cookies)
        self.cookies_button.grid(row=1, column=3, padx=5, pady=10, sticky="w")

        # ====== 頁面下方 ======
        self.frame_bottom = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),  # 明暗主題雙色
            fg_color=("#FFFFFF", "#000001"),  # 明暗主題雙色
        )
        self.frame_bottom.grid(row=2, column=0, sticky="nsew")
        self.frame_bottom.grid_rowconfigure(0, weight=1)
        self.frame_bottom.grid_columnconfigure((0,1,2,3,4,5,6,7), weight=1)

        # 儲存按鈕
        self.save_button = ctk.CTkButton(self.frame_bottom, width=70, command=self.save_settings)
        self.save_button.grid(row=0, column=7, pady=5)

        self.update_all_objects()

    def select_tab(self, selected_label):
        selected_key = self.label_to_key[selected_label]
        for key, frame in self.frames.items():
            if key == selected_key:
                frame.grid()
            else:
                frame.grid_remove()
    
    def get_cur_nav_index(self):
        try:
            curent_nav_value = self.nav.get()
            if curent_nav_value in self.tab_labels:
                cur_index = self.tab_labels.index(curent_nav_value)
            else:
                cur_index = 0  # 預設為第一個
        except Exception as e:
            log_and_show_error(f"Error getting current nav index: {e}", self.master)
            cur_index = 0
        return cur_index

    def update_theme(self, choice):
        ctk.set_appearance_mode(choice)
        self.master.config["theme"] = choice

    def import_theme_json(self):
        # 讓使用者選擇圖片檔案，限定 json 格式
        file_path = filedialog.askopenfilename(
            title=LANGUAGES[self.master.current_language]["setting"]["choose_theme_color"],
            filetypes=[("Image Files", "*.json*")]
        )
        if file_path:
            # 儲存選擇的圖片路徑到 config 中
            self.master.config["theme_color"] = file_path
            save_config(self.master.config)
    
    def import_bg_image(self):
        file_path = filedialog.askopenfilename(
            title=LANGUAGES[self.master.current_language]["setting"]["choose_bg_image"],
            filetypes=[("Image Files", "*.jpg;*.jpeg;*.png;*.gif")]
        )
        if file_path:
            self.master.config["bg_image"] = file_path
            save_config(self.master.config)

    def change_language(self, choice):
        if choice == "繁體中文":
            self.master.current_language = "zh-TW"
        elif choice == "简体中文":
            self.master.current_language = "zh-CN"
        elif choice == "English":
            self.master.current_language = "en"
        elif choice == "日本語":
            self.master.current_language = "ja"
        elif choice == "Español":
            self.master.current_language = "es"    
        self.master.config["language"] = self.master.current_language
    
    def import_ad_image(self):
        # 讓使用者選擇圖片檔案，限定 jpg、jpeg、png 與 gif 格式
        file_path = filedialog.askopenfilename(
            title=LANGUAGES[self.master.current_language]["setting"]["choose_ad_image"],
            filetypes=[("Image Files", "*.jpg;*.jpeg;*.png;*.gif")]
        )
        if file_path:
            # 儲存選擇的圖片路徑到 config 中
            self.master.config["ad_image"] = file_path
            save_config(self.master.config)
    
    def import_bg_image(self):
        file_path = filedialog.askopenfilename(
            title=LANGUAGES[self.master.current_language]["setting"]["choose_bg_image"],
            filetypes=[("Image Files", "*.jpg;*.jpeg;*.png;*.gif")]
        )
        if file_path:
            self.master.config["bg_image"] = file_path
            save_config(self.master.config)

    def import_cookies(self):
        # 讓使用者選擇 cookies 檔案，限定 txt 格式
        file_path = filedialog.askopenfilename(
            title=LANGUAGES[self.master.current_language]["setting"]["choose_cookies_file"],
            filetypes=[("Text Files", "*.txt")]
        )
        if file_path:
            self.master.config["cookies"] = file_path
            save_config(self.master.config)
    
    def reset_cookies(self):
        """
        重設 cookies 路徑，將其設為空字串。
        """
        self.master.config["cookies"] = ""
        save_config(self.master.config)
        messagebox.showinfo(
            LANGUAGES[self.master.current_language]['setting']["setting_title"],
            LANGUAGES[self.master.current_language]['setting']["cookies_reset"]
        )
    
    def save_settings(self):
        self.update_theme(self.theme_combobox.get())
        self.change_language(self.language_combobox.get())
        self.master.config["resolution"] = self.resolution_combobox.get()
        self.master.config["transparency"] = self.transparency_entry.get()
        save_config(self.master.config)
        self.update_all_objects()
        self.master.update_all_pages_objects()
        messagebox.showinfo(
            LANGUAGES[self.master.current_language]['setting']["setting_title"],
            LANGUAGES[self.master.current_language]['setting']["saved_success"]
        )

    def update_bg_image(self):
        bg_image_path = self.master.bg_image_path 
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                img = Image.open(bg_image_path)
                img = ImageOps.fit(img, (self.winfo_width(), self.winfo_height()), Image.LANCZOS) # 圖片比例不變但填滿
                self.bg_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.bg_label.configure(text="", image=self.bg_image)
            except Exception as e:
                log_and_show_error(f"Background image load failed: {e}", self.master)
                self.bg_label.configure(text="", image=None)

    def update_text(self):
        lang = self.master.current_language
        self.title(LANGUAGES[lang]["setting"]["setting_title"])

        # 重新產生 tabs
        self.tabs = [
            {"key": "theme", "label": LANGUAGES[lang]["setting"]["segmentedBtn"]["theme"]},
            {"key": "language", "label": LANGUAGES[lang]["setting"]["segmentedBtn"]["language"]},
            {"key": "resolution", "label": LANGUAGES[lang]["setting"]["segmentedBtn"]["resolution"]},
            {"key": "picture", "label": LANGUAGES[lang]["setting"]["segmentedBtn"]["picture"]},
            {"key": "others", "label": LANGUAGES[lang]["setting"]["segmentedBtn"]["others"]},
        ]
        self.tab_labels   = [tab["label"] for tab in self.tabs]
        self.label_to_key = {tab["label"]:tab["key"] for tab in self.tabs}
        self.nav.configure(values=self.tab_labels, font=self.master.FONT_BUTTON)

        self.theme_label.configure(text=LANGUAGES[lang]["setting"]["theme_label"], font=self.master.FONT_BODY)
        self.theme_combobox.configure(values=["Light", "Dark"], font=self.master.FONT_BODY)
        self.language_combobox.configure(values=list(LANGUAGE_OPTIONS.values()), font=self.master.FONT_BODY)
        self.resolution_combobox.configure(values=["1280x720"], font=self.master.FONT_BODY)
        self.theme_color_label.configure(text=LANGUAGES[lang]["setting"]["theme_color_label"], font=self.master.FONT_BODY)
        self.theme_color_btn.configure(text=LANGUAGES[lang]["setting"]["theme_color_btn"], font=self.master.FONT_BUTTON)
        self.language_label.configure(text=LANGUAGES[lang]["setting"]["language_label"], font=self.master.FONT_BODY)
        self.resolution_label.configure(text=LANGUAGES[lang]["setting"]["resolution_label"], font=self.master.FONT_BODY)
        self.ad_label.configure(text=LANGUAGES[lang]["setting"]["set_ad_image"], font=self.master.FONT_BODY)
        self.ad_button.configure(text=LANGUAGES[lang]["setting"]["import_image"], font=self.master.FONT_BUTTON)
        self.background_label.configure(text=LANGUAGES[lang]["setting"]["set_bg_image"], font=self.master.FONT_BODY)
        self.bg_button.configure(text=LANGUAGES[lang]["setting"]["import_image"], font=self.master.FONT_BUTTON)
        self.transparency_label.configure(text=LANGUAGES[lang]["setting"]["transparency_label"], font=self.master.FONT_BODY)
        self.cookies_label.configure(text=LANGUAGES[lang]["setting"]["cookies_label"], font=self.master.FONT_BODY)
        self.cookies_reset_button.configure(text=LANGUAGES[lang]["setting"]["reset"], font=self.master.FONT_BUTTON)
        self.cookies_button.configure(text=LANGUAGES[lang]["setting"]["import_cookies"], font=self.master.FONT_BUTTON)
        self.save_button.configure(text=LANGUAGES[lang]["setting"]["save_button"], font=self.master.FONT_BUTTON)

    def update_frame_tranparency(self):
        # 根據主題設定物件透明度
        # 注意! 若frame已經消除透明度，內部物件則不需要再設定
        background_color = "#FFFFFF" if self.master.config.get("theme", "Dark") == "Light" else "#000001"
        pywinstyles.set_opacity(self.frame_top, value=self.master.transparency, color=background_color)
        for key, frame in self.frames.items():
            pywinstyles.set_opacity(frame, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_bottom, value=self.master.transparency, color=background_color)

    def update_all_objects(self):
        cur_nav_index = self.get_cur_nav_index()
        self.update_bg_image()
        self.update_text()
        self.update_frame_tranparency()
        self.select_tab(self.tab_labels[cur_nav_index])
        self.nav.set(self.tab_labels[cur_nav_index])

# ------------------------------
# 主視窗
# ------------------------------

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        # 讀取設定檔
        self.config = load_config()
        # 從設定檔中取得設定，若無則採用預設值
        self.current_language = self.config.get("language", "zh-TW")
        self.current_theme = self.config.get("theme", "Dark")
        resolution = self.config.get("resolution", "1280x720")
        self.download_path = self.config.get("download_path", os.getcwd())
        self.bg_image_path = self.config.get("bg_image", "")
        self.transparency = float(self.config.get("transparency", "1"))
        self.cookies_path = self.config.get("cookies", "")

        # 根據語言載入字體
        self.FONT_LOGO = get_font(self.current_language, "logo")
        self.FONT_TITLE = get_font(self.current_language, "title")
        self.FONT_BODY = get_font(self.current_language, "body")
        self.FONT_BUTTON = get_font(self.current_language, "button")

        ctk.set_appearance_mode(self.current_theme)
        self.title("Video DownloadErm")
        icon_path = os.path.join(os.path.dirname(__file__), 'assets/icon/icon.ico')
        self.iconbitmap(icon_path)
        self.geometry(resolution)
        self.resizable(False, False)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frames = {}
        for Page in (HomePage, Page1, Page2, Page3, Page4):
            page = Page(self)
            self.frames[Page] = page
            page.grid(row=0, column=0, sticky="nsew")

        self.show_frame(HomePage)
        self.setting_window = None

    def show_frame(self, page):
        frame = self.frames[page]
        frame.tkraise()
    
    def open_Setting(self):
        if self.setting_window is None or not self.setting_window.winfo_exists():
            self.setting_window = Setting(self)  # create window if its None or destroyed
            self.after(300, lambda: self.setting_window.focus()) # focus the window
        else:
            self.setting_window.focus()  # if window exists focus it
    
    def update_all_pages_objects(self):
        """
        被 Setting 視窗呼叫，用來更新整個主視窗以及所有 Page 裏頭的物件。
        """
        # 再次讀取設定檔
        self.config = load_config()
        # 從設定檔中取得設定，若無則採用預設值
        self.current_language = self.config.get("language", "zh-TW")
        self.current_theme = self.config.get("theme", "Dark")
        resolution = self.config.get("resolution", "1280x720")
        self.download_path = self.config.get("download_path", os.getcwd())
        self.bg_image_path = self.config.get("bg_image", "")
        #self.fg_color = self.config.get("fg_color", "#000000")
        self.transparency = float(self.config.get("transparency", "0.85"))
        self.cookies_path = self.config.get("cookies", "")

        # 根據語言再次載入字體
        self.FONT_LOGO = get_font(self.current_language, "logo")
        self.FONT_TITLE = get_font(self.current_language, "title")
        self.FONT_BODY = get_font(self.current_language, "body")
        self.FONT_BUTTON = get_font(self.current_language, "button")

        # 更新Setting 視窗
        if self.setting_window is not None and self.setting_window.winfo_exists():
            self.setting_window.update_all_objects()

        # 更新每一個 Page 的物件
        for page_class, page_obj in self.frames.items():
            page_obj.update_all_objects()

class HomePage(ctk.CTkFrame):  # 主页
    def __init__(self, master):
        super().__init__(master)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=7)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Background image initialization
        self.bg_label = ctk.CTkLabel(self, text="")
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        # ====== 頁面頂部 ======
        self.frame_top = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_top.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.frame_top.grid_columnconfigure((0,1,2,3), weight=1)
        self.frame_top.grid_rowconfigure(0, weight=1)

        self.edit_icon = ctk.CTkButton(
            self.frame_top, 
            text="\U00002699", 
            width=50, 
            command=master.open_Setting, 
            bg_color=("#FFFFFF", "#000001"), 
            text_color=("#000001", "#FFFFFF"), 
            fg_color=("transparent"),
            border_width=0
        )
        self.edit_icon.grid(row=0, column=3, padx=5, sticky="e")

        # ====== 頁面內容 ======
        self.frame_main =  ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_main.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        self.frame_main.grid_columnconfigure((0,1,2,3), weight=1)
        self.frame_main.grid_rowconfigure(0, weight=2)
        self.frame_main.grid_rowconfigure((1,2), weight=1)

        # Logo 區域
        self.logo_label = ctk.CTkLabel(
            self.frame_main,
            bg_color=("#FFFFFF", "#000001"),
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=5, columnspan=4, sticky="nsew")

        # 跳轉按鈕
        self.btn1 = ctk.CTkButton(
            self.frame_main,
            command=lambda: master.show_frame(Page1),
            width=256,
            height=40,
            corner_radius=12, 
            bg_color=("#FFFFFF", "#000001"), 
        )
        self.btn1.grid(row=1, column=1, pady=5)

        self.btn2 = ctk.CTkButton(
            self.frame_main, 
            command=lambda: master.show_frame(Page2),
            width=256,
            height=40,
            corner_radius=12,
            bg_color=("#FFFFFF", "#000001"), 
        )
        self.btn2.grid(row=1, column=2, pady=5)

        self.btn3 = ctk.CTkButton(
            self.frame_main, 
            command=lambda: master.show_frame(Page3),
            width=256,
            height=40,
            corner_radius=12, 
            bg_color=("#FFFFFF", "#000001"), 
        )
        self.btn3.grid(row=2, column=1, pady=5)

        self.btn4 = ctk.CTkButton(
            self.frame_main, 
            command=lambda: master.show_frame(Page4),
            width=256,
            height=40,
            corner_radius=12,
            bg_color=("#FFFFFF", "#000001"), 
        )
        self.btn4.grid(row=2, column=2, pady=5)

        # 初始化
        self.update_all_objects()

    def update_bg_image(self):
        bg_image_path = self.master.bg_image_path 
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                img = Image.open(bg_image_path)
                img = ImageOps.fit(img, (1280, 720), Image.LANCZOS) # 圖片比例不變但填滿
                self.bg_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.bg_label.configure(text="", image=self.bg_image)
            except Exception as e:
                log_and_show_error(f"Background image load failed: {e}", self.master)
                self.bg_label.configure(text="", image=None)

    def update_logo_area(self):
        icon_path = os.path.join(os.path.dirname(__file__), 'assets/icon/icon_r.png')
        if icon_path and os.path.exists(icon_path):
            try:
                img = Image.open(icon_path)
                img = ImageOps.contain(img, (336, 216))
                self.logo_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.logo_label.configure(image=self.logo_image, text=LANGUAGES[self.master.current_language]["homePage"]["title_label"], compound="left", padx=15)
            except Exception as e:
                log_and_show_error(f"Logo image load failed: {e}", self.master)
                self.logo_label.configure(text=LANGUAGES[self.master.current_language]["homePage"]["title_label"], image=None, font=self.master.FONT_BODY)
        else:
            self.logo_label.configure(text=LANGUAGES[self.master.current_language]["homePage"]["title_label"], image=None, font=self.master.FONT_BODY)
       
    def update_text(self):
        lang = self.master.current_language
        text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"

        self.logo_label.configure(text=LANGUAGES[lang]["homePage"]["title_label"], font=self.master.FONT_LOGO)
        self.btn1.configure(text=f"🎬 {LANGUAGES[lang]['page1']['page1_title']}", font=self.master.FONT_BUTTON)
        self.btn2.configure(text=f"📋 {LANGUAGES[lang]['page2']['page2_title']}", font=self.master.FONT_BUTTON)
        self.btn3.configure(text=f"🔄 {LANGUAGES[lang]['page3']['page3_title']}", font=self.master.FONT_BUTTON)
        self.btn4.configure(text=f"🔊 {LANGUAGES[lang]['page4']['page4_title']}", font=self.master.FONT_BUTTON)

    def update_frame_tranparency(self):
        # 根據主題設定物件透明度
        # 注意! 若frame已經消除透明度，內部物件則不需要再設定
        background_color = "#FFFFFF" if self.master.config.get("theme", "Dark") == "Light" else "#000001"
        pywinstyles.set_opacity(self.frame_top, color=background_color)
        pywinstyles.set_opacity(self.frame_main, color=background_color)

    def update_all_objects(self):
        self.update_bg_image()  # 更新背景圖片
        self.update_text() # 初始化文字
        self.update_logo_area() # 更新 logo 區域
        self.update_frame_tranparency() # 更新透明度

class Page1(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        
        self.download_path = self.master.config.get("download_path") or os.getcwd()
        self.video_url = ""
        
        # 設定 Grid 權重
        self.grid_columnconfigure(0, weight=7)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1) 
        self.grid_rowconfigure(1, weight=2) 
        self.grid_rowconfigure(2, weight=5)
        self.grid_rowconfigure(3, weight=2)

        text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"

        # Background image initialization
        self.bg_label = ctk.CTkLabel(self, text="")
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        # ====== 頁面頂部 ======
        self.frame_top = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_top.grid(row=0, column=0, sticky="nsew", columnspan=2, padx=5)
        self.frame_top.grid_rowconfigure(0, weight=1)
        self.frame_top.grid_columnconfigure((0,2), weight=1)
        self.frame_top.grid_columnconfigure(1, weight=8)

        self.back_icon = ctk.CTkButton(
            self.frame_top, 
            text="\U00002190", 
            width=50, 
            command=lambda: master.show_frame(HomePage),
            bg_color=("#FFFFFF", "#000001"), 
            fg_color="transparent",
            text_color=text_color, 
            border_width=0, 
        )
        self.back_icon.grid(row=0, column=0, padx=5, sticky="w")

        self.page_title = ctk.CTkLabel(
            self.frame_top,
            bg_color=("#FFFFFF", "#000001"),  
        )
        self.page_title.grid(row=0, column=1, padx=5, sticky="ew")

        self.edit_icon = ctk.CTkButton(
            self.frame_top, 
            text="\U00002699", 
            width=50, 
            command=master.open_Setting,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color="transparent", 
            text_color=text_color,
            border_width=0, 
        )
        self.edit_icon.grid(row=0, column=2, padx=5, sticky="e")
        
        # ====== 左側 Frame（影片資訊 & 下載選項）======
        self.frame_left = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_left.grid(row=1, column=0, sticky="nsew", rowspan=2, padx=5, pady=5)
        self.frame_left.grid_rowconfigure(0, weight=4)
        self.frame_left.grid_rowconfigure((1,2,3), weight=2)
        self.frame_left.grid_columnconfigure(0, weight=2)
        self.frame_left.grid_columnconfigure(1, weight=5)
        self.frame_left.grid_columnconfigure(2, weight=3)
        
        self.thumbnail_label = ctk.CTkLabel(self.frame_left, text="", width=400, height=300)
        self.thumbnail_label.grid(row=0, column=0, columnspan=3, pady=3)
        
        self.video_title_label = ctk.CTkLabel(self.frame_left, wraplength=400, text="")
        self.video_title_label.grid(row=1, column=0, columnspan=3, pady=1)
        
        self.resolution_combobox = ctk.CTkComboBox(self.frame_left, values=["No resolutions"])
        self.resolution_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        self.format_var = ctk.StringVar(value="mp4")
        self.radio_mp4 = ctk.CTkRadioButton(self.frame_left, text="MP4", variable=self.format_var, value="mp4")
        self.radio_mp3 = ctk.CTkRadioButton(self.frame_left, text="MP3", variable=self.format_var, value="mp3")
        self.radio_mp4.grid(row=2, column=2, padx=30, pady=5, sticky="w")
        self.radio_mp3.grid(row=3, column=2, padx=30, pady=5, sticky="w")
        
        self.subtitle_combobox = ctk.CTkComboBox(self.frame_left, values=["No subtitle"])
        self.subtitle_combobox.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        self.subtitle_combobox.grid_remove()  # 隱藏字幕選項
        
        # ====== 右側 Frame（URL 輸入 & 下載位置 & 其他操作 & 廣告）======
        # 上半部分
        self.frame_first_right = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_first_right.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        self.frame_first_right.grid_rowconfigure((0,1,2), weight=1)
        self.frame_first_right.grid_columnconfigure(0, weight=8)
        self.frame_first_right.grid_columnconfigure(1, weight=2)
        
        self.url_entry = ctk.CTkEntry(self.frame_first_right)
        self.url_entry.grid(row=0, column=0, padx=10, pady=2, sticky="ew")

        self.paste_button = ctk.CTkButton(self.frame_first_right, text="📋", width=30, font=("Arial", 16), command=self.paste_url) 
        self.paste_button.grid(row=0, column=1, padx=0, pady=2, sticky="w")
        
        self.submit_button = ctk.CTkButton(self.frame_first_right, command=self.fetch_video_info)
        self.submit_button.grid(row=0, column=1, padx=10, pady=2, sticky="e")
        
        self.download_path_textbox = ctk.CTkTextbox(self.frame_first_right, height=30, activate_scrollbars=False)
        self.download_path_textbox.configure(state="disabled")
        self.download_path_textbox.grid(row=1, column=0, padx=10, pady=2, sticky="ew")
        
        self.change_path_button = ctk.CTkButton(self.frame_first_right, command=self.change_download_path)
        self.change_path_button.grid(row=1, column=1, padx=10, pady=2, sticky="e")

        self.download_sub_var = ctk.BooleanVar(value=False)
        self.download_sub_checkbox = ctk.CTkCheckBox(
            self.frame_first_right, 
            variable=self.download_sub_var, 
            command=self.toggle_subtitle_combobox
        )
        self.download_sub_checkbox.grid(row=2, column=0, padx=10, pady=2, sticky="w")

        # ====== 右下 Frame（廣告區）======
        self.frame_second_right = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"),
        )
        self.frame_second_right.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
        self.frame_second_right.grid_propagate(False)  # 固定 frame 尺寸
        self.frame_second_right.grid_rowconfigure(0, weight=1)
        self.frame_second_right.grid_columnconfigure(0, weight=1)

        self.ad_label = ctk.CTkLabel(self.frame_second_right, text='')
        self.ad_label.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)

        # ====== 底部 Frame ======
        self.frame_bottom = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"),
        )
        self.frame_bottom.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.frame_bottom.grid_rowconfigure(0, weight=4)
        self.frame_bottom.grid_rowconfigure(1, weight=6)
        self.frame_bottom.grid_columnconfigure(0, weight=7)
        self.frame_bottom.grid_columnconfigure(1, weight=3)

        self.progress_bar_label = ctk.CTkLabel(self.frame_bottom)
        self.progress_bar_label.grid(row=0, column=0, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self.frame_bottom)
        self.progress_bar.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.progress_bar.set(0)

        self.download_button = ctk.CTkButton(self.frame_bottom, command=self.download_video)
        self.download_button.grid(row=1, column=1, pady=5)
        
        self.update_all_objects()

    def toggle_subtitle_combobox(self):
        if self.download_sub_var.get():
            self.subtitle_combobox.grid()  # 顯示字幕下拉選單
        else:
            self.subtitle_combobox.grid_remove()  # 隱藏字幕下拉選單

    def paste_url(self):
        try:
            clipboard = self.clipboard_get()
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, clipboard)
        except Exception as e:
            log_and_show_error(f"Failed to paste URL from clipboard: {e}", self.master)

    def fetch_video_info(self):
        """獲取影片資訊並更新 UI（thread 版本）"""
        url = self.url_entry.get()
        if not url:
            return
        self.video_url = url

        self.info_stop_event = threading.Event()
        self.info_thread = None

        def get_info_task():
            try:
                self.master.after(0, lambda: self.submit_button.configure(state="disabled"))  
                # Main task: 獲取影片資訊
                title, thumbnail_url, resolutions, subtitles = get_video_info(
                    url, self.format_var.get(), self.master.cookies_path
                )
                if self.info_stop_event.is_set():
                    # 被用戶終止，不更新 UI
                    return
                # 回到主執行緒更新 UI
                def update_ui():
                    self.video_title_label.configure(text=title)
                    self.resolution_combobox.configure(values=resolutions)
                    if resolutions:
                        self.resolution_combobox.set(resolutions[0])
                    else:
                        self.resolution_combobox.set("No resolutions")
                    self.subtitle_combobox.configure(values=subtitles)
                    self.subtitle_combobox.set(subtitles[0])
                    # 更新封面圖
                    response = requests.get(thumbnail_url)
                    img_data = Image.open(io.BytesIO(response.content))
                    self.thumbnail_image = CTkImage(light_image=img_data, dark_image=img_data, size=(400, 300))
                    self.thumbnail_label.configure(image=self.thumbnail_image, text="")
                    # 啟用提交按鈕
                    self.submit_button.configure(state="normal")
                self.master.after(0, update_ui)
            except Exception as e:
                if not self.info_stop_event.is_set():
                    log_and_show_error(f"Failed to fetch video info: {e}", self.master)
                self.master.after(0, lambda: self.submit_button.configure(state="normal"))

        def ask_cancel():
            if self.info_thread.is_alive():
                result = messagebox.askyesno(
                    LANGUAGES[self.master.current_language]["page1"]["timeout_title"],
                    LANGUAGES[self.master.current_language]["page1"]["timeout_message"]
                )
                if result:
                    self.info_stop_event.set()
                    self.submit_button.configure(state="normal")

        self.info_thread = threading.Thread(target=get_info_task)
        self.info_thread.start()
        # 10秒後詢問是否終止
        self.master.after(10000, ask_cancel)

    def update_progress(self, progress):
        if progress != -1:
            percent = int(progress * 100)
            # 使用 after 確保 GUI 更新在主執行緒中執行
            self.master.after(0, lambda: self.progress_bar.set(progress))
            self.master.after(0, lambda: self.progress_bar_label.configure(text=f"{LANGUAGES[self.master.current_language]['page1']['Processing']} {percent}%", font=self.master.FONT_BODY))
        else:
            self.master.after(100, lambda: self.progress_bar.set(progress))
            self.master.after(0, lambda: self.progress_bar_label.configure(text=LANGUAGES[self.master.current_language]['page1']['Processing_completed'], font=self.master.FONT_BODY))

    def download_video(self):
        """開始下載影片，使用 threading 執行下載任務"""
        self.download_button.configure(state="disabled")
        resolution = self.resolution_combobox.get()
        file_format = self.format_var.get()
        download_subtitles = self.download_sub_var.get()
        subtitle_lang = self.subtitle_combobox.get()

        def download_task():
            try:
                # 呼叫 yt_dlp 的 Python API 進行下載
                output_file = download_video_audio(
                    self.video_url, resolution, self.download_path,
                    file_format, download_subtitles,
                    subtitle_lang, self.master.cookies_path, self.update_progress
                )
                logger.info(f"Download Completed: {output_file}")
                self.master.after(0, lambda: messagebox.showinfo(
                    LANGUAGES[self.master.current_language]['page1']["download_complete_title"],
                    LANGUAGES[self.master.current_language]['page1']["download_complete_message"].format(output_file)
                ))
            except Exception as e:
                log_and_show_error(f"Download failed: {e}", self.master)
            finally:
                # 回到主執行緒後重新啟用下載按鈕
                self.master.after(0, lambda: self.download_button.configure(state="normal"))

        # 建立並啟動下載執行緒
        download_thread = threading.Thread(target=download_task)
        download_thread.start()

    def change_download_path(self):
        """變更下載位置"""
        lang = self.master.current_language
        path = filedialog.askdirectory()
        if path:
            self.download_path = path
            self.master.config["download_path"] = path
            save_config(self.master.config)

            self.download_path_textbox.configure(state="normal")
            self.download_path_textbox.delete("0.0", "end")
            self.download_path_textbox.insert("0.0", f"{LANGUAGES[self.master.current_language]['page1']['download_path_label']} {self.download_path}")
            self.download_path_textbox.configure(state="disabled")

    def update_bg_image(self):
        bg_image_path = self.master.bg_image_path 
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                img = Image.open(bg_image_path)
                img = ImageOps.fit(img, (1280, 720), Image.LANCZOS) # 圖片比例不變但填滿
                self.bg_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.bg_label.configure(text="", image=self.bg_image)
            except Exception as e:
                log_and_show_error(f"Background image load failed: {e}", self.master)
                self.bg_label.configure(text="", image=None)

    def update_ad_area(self):
        ad_image_path = self.master.config.get("ad_image", "")
        if ad_image_path and os.path.exists(ad_image_path):
            try:
                img = Image.open(ad_image_path)
                img = ImageOps.contain(img, (640, 480))
                self.ad_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.ad_label.configure(image=self.ad_image, text="")
            except Exception as e:
                log_and_show_error(f"AD image load failed: {e}", self.master)
                self.ad_label.configure(text=LANGUAGES[self.master.current_language]["page1"]["no_ad_label"], image=None, font=self.master.FONT_BODY)
        else:
            self.ad_label.configure(text=LANGUAGES[self.master.current_language]["page1"]["no_ad_label"], image=None, font=self.master.FONT_BODY)

    def update_text(self):
        lang = self.master.current_language
        text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"

        self.back_icon.configure(text_color=text_color, font=self.master.FONT_BUTTON)
        self.page_title.configure(text=LANGUAGES[lang]["page1"]["page1_title"], text_color=text_color, font=self.master.FONT_TITLE)
        self.edit_icon.configure(text_color=text_color, font=self.master.FONT_BUTTON)

        self.url_entry.configure(placeholder_text=LANGUAGES[lang]["page1"]["url_entry_placeholder"], font=self.master.FONT_BODY)
        self.submit_button.configure(text=LANGUAGES[lang]["page1"]["submit_button"], font=self.master.FONT_BUTTON)
        self.download_path_textbox.configure(state="normal", font=self.master.FONT_BODY)
        self.download_path_textbox.delete("0.0", "end")
        self.download_path_textbox.insert("0.0", f"{LANGUAGES[lang]['page1']['download_path_label']} {self.download_path}")
        self.download_path_textbox.configure(state="disabled")
        self.change_path_button.configure(text=LANGUAGES[lang]["page1"]["browse_button"], font=self.master.FONT_BUTTON)
        self.download_sub_checkbox.configure(text=LANGUAGES[lang]["page1"]["download_sub_checkbox"], font=self.master.FONT_BODY)

        self.resolution_combobox.configure(font=self.master.FONT_BODY)
        self.subtitle_combobox.configure(font=self.master.FONT_BODY)

        self.progress_bar_label.configure(text=LANGUAGES[lang]["page1"]["progress_ready"], font=self.master.FONT_BODY)
        self.download_button.configure(text=LANGUAGES[lang]["page1"]["download_button"], font=self.master.FONT_BUTTON)

    def update_frame_tranparency(self):
        # 根據主題設定物件透明度
        # 注意! 若frame已經消除透明度，內部物件則不需要再設定
        background_color = "#FFFFFF" if self.master.config.get("theme", "Dark") == "Light" else "#000001"
        pywinstyles.set_opacity(self.frame_top, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_left, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_first_right, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_second_right, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_bottom, value=self.master.transparency, color=background_color)

    def update_all_objects(self):
        self.update_bg_image()
        self.update_text()
        self.update_ad_area()
        self.update_frame_tranparency()

class Page2(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.download_path = self.master.config.get("download_path") or os.getcwd()
        self.video_url = ""
        
        # 播放清單資料，內部儲存，每筆為 dict
        self.playlist_items = []

        # 設定 Grid 權重
        self.grid_columnconfigure(0, weight=7)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1) 
        self.grid_rowconfigure(1, weight=2) 
        self.grid_rowconfigure(2, weight=5)
        self.grid_rowconfigure(3, weight=2)

        text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"

        # Background image initialization
        self.bg_label = ctk.CTkLabel(self, text="")
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        # ====== 頁面頂部 ======
        self.frame_top = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"),
        )
        self.frame_top.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.frame_top.grid_rowconfigure(0, weight=1)
        self.frame_top.grid_columnconfigure((0,2), weight=1)
        self.frame_top.grid_columnconfigure(1, weight=8)
        
        self.back_btn = ctk.CTkButton(
            self.frame_top, 
            text="\U00002190", 
            width=50, 
            command=lambda: master.show_frame(HomePage),
            fg_color="transparent",
            text_color=text_color, 
            border_width=0, 
            border_spacing=0, 
            corner_radius=2
        )
        self.back_btn.grid(row=0, column=0, padx=5, sticky="w")
        
        self.page_title = ctk.CTkLabel(self.frame_top)
        self.page_title.grid(row=0, column=1, padx=5, sticky="ew")
        
        self.edit_btn = ctk.CTkButton(
            self.frame_top, 
            text="\U00002699", 
            width=50, 
            command=master.open_Setting,
            fg_color="transparent", 
            text_color=text_color,
            border_width=0, 
            border_spacing=0, 
            corner_radius=2
        )
        self.edit_btn.grid(row=0, column=2, padx=5, sticky="e")
        
        # ====== 左側 Frame（播放清單表格）======
        self.frame_left_first = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_left_first.grid(row=1, column=0, sticky="nsew", rowspan=2, padx=5, pady=5)
        self.frame_left_first.grid_rowconfigure(0, weight=9)
        self.frame_left_first.grid_rowconfigure(1, weight=1)
        self.frame_left_first.grid_columnconfigure((0,1,2,3,4,5), weight=1)

        self.table_scroll = ctk.CTkScrollableFrame(
            self.frame_left_first,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"),

        )
        self.table_scroll.grid(row=0, column=0, columnspan=6, sticky="nsew")

        initial_data = [["Video Title", "Resolution", "Format", "URL"]]
        self.table = CTkTable(
            master=self.table_scroll,
            row=len(initial_data),
            column=len(initial_data[0]),
            values=initial_data,
            hover_color="skyblue",
            corner_radius=0,
            font=self.master.FONT_BODY,
            command=self.on_cell_click
        )
        self.table.pack(padx=10, pady=10, fill="both", expand=True)


        self.select_all_btn = ctk.CTkButton(self.frame_left_first, command=self.select_all_rows)
        self.select_all_btn.grid(row=1, column=0, sticky="w", padx=5, pady=1)

        self.delete_btn = ctk.CTkButton(self.frame_left_first, command=self.delete_selected_rows)
        self.delete_btn.grid(row=1, column=1, sticky="w", padx=5, pady=1)
        
        self.total_label = ctk.CTkLabel(self.frame_left_first)
        self.total_label.grid(row=1, column=5, sticky="e", padx=5, pady=1)
        
        # ====== 右上 Frame（URL 輸入 & 下載位置）======
        self.frame_first_right = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"),
        )
        self.frame_first_right.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        self.frame_first_right.grid_rowconfigure((0,1,2,3), weight=1)
        self.frame_first_right.grid_columnconfigure(0, weight=8)
        self.frame_first_right.grid_columnconfigure(1, weight=2)

        self.url_entry = ctk.CTkEntry(self.frame_first_right)
        self.url_entry.grid(row=0, column=0, padx=10, pady=2, sticky="ew")

        self.paste_button = ctk.CTkButton(self.frame_first_right, text="📋", width=30, font=("Arial", 16), command=self.paste_url) 
        self.paste_button.grid(row=0, column=1, padx=0, pady=2, sticky="w")

        self.submit_btn = ctk.CTkButton(self.frame_first_right, command=self.add_playlist_item)
        self.submit_btn.grid(row=0, column=1, padx=10, pady=2, sticky="e")

        self.download_path_textbox = ctk.CTkTextbox(self.frame_first_right, height=30, activate_scrollbars=False)
        self.download_path_textbox.configure(state="disabled")
        self.download_path_textbox.grid(row=1, column=0, padx=10, pady=2, sticky="ew")
        
        self.change_path_button = ctk.CTkButton(self.frame_first_right, command=self.change_download_path)
        self.change_path_button.grid(row=1, column=1, padx=10, pady=2, sticky="e")
        
        self.resolution_combobox = ctk.CTkComboBox(self.frame_first_right)
        self.resolution_combobox.grid(row=2, column=0, padx=10, pady=2, sticky="ew")
        
        self.format_var = ctk.StringVar(value="mp4")
        self.mp4_radio = ctk.CTkRadioButton(self.frame_first_right, text="MP4", variable=self.format_var, value="mp4")
        self.mp3_radio = ctk.CTkRadioButton(self.frame_first_right, text="MP3", variable=self.format_var, value="mp3")
        self.mp4_radio.grid(row=2, column=1, padx=10, pady=2)
        self.mp3_radio.grid(row=3, column=1, padx=10, pady=2)

        # 監聽 self.format_var 的變化，當格式改變時自動更新 resolution_combobox 的選項
        self.format_var.trace_add('write', lambda *args: self.update_resolution_options())

        # ====== 右下 Frame（廣告區）======
        self.frame_second_right = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"),
        )
        self.frame_second_right.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
        self.frame_second_right.grid_propagate(False)  # 固定 frame 尺寸
        self.frame_second_right.grid_rowconfigure(0, weight=1)
        self.frame_second_right.grid_columnconfigure(0, weight=1)

        self.ad_label = ctk.CTkLabel(self.frame_second_right, text='')
        self.ad_label.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        
        # ====== 底部 Frame ======
        self.frame_bottom = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_bottom.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.frame_bottom.grid_rowconfigure(0, weight=4)
        self.frame_bottom.grid_rowconfigure(1, weight=6)
        self.frame_bottom.grid_columnconfigure(0, weight=7)
        self.frame_bottom.grid_columnconfigure(1, weight=3)

        self.progress_bar_label = ctk.CTkLabel(self.frame_bottom)
        self.progress_bar_label.grid(row=0, column=0, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self.frame_bottom)
        self.progress_bar.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.progress_bar.set(0)

        self.download_button = ctk.CTkButton(self.frame_bottom, command=self.download_playlist)
        self.download_button.grid(row=1, column=1, pady=5)

        self.update_all_objects()  # 初始化所有物件的文字與樣式
        
    def set_fixed_column_widths(self, widths):
        """
        設定表格每一欄的寬度。
        widths: 一個列表，每個元素代表對應欄的固定寬度（像素）。
        例如: [300, 100, 80, 400]
        (但沒啥用)
        """
        for (row, col), cell in self.table.frame.items():
            if col in widths:
                # 設定固定寬度，同時設定 wraplength 避免文字溢出
                cell.configure(width=widths[col])

    def update_table_header(self):
        lang = self.master.current_language
        # 根據語系設定表頭
        header = [
            LANGUAGES[lang]["page2"]["video_title"],
            LANGUAGES[lang]["page2"]["resolution"],
            LANGUAGES[lang]["page2"]["format"],
            LANGUAGES[lang]["page2"]["url"]
        ]
        # 根據主題決定表頭背景色，這裡以 Light 主題用淺灰、Dark 主題用深灰為例
        header_color = "gray90" if self.master.config.get("theme", "Dark") == "Light" else "gray25"
        # 更新表頭每個 cell 的文字與背景色
        for col in range(len(header)):
            if (0, col) in self.table.frame:
                self.table.frame[(0, col)].configure(text=header[col], fg_color=header_color)
        self.set_fixed_column_widths([400, 200, 100, 100])

    def update_total_label(self):
        total = len(self.playlist_items)
        self.total_label.configure(text=f"{LANGUAGES[self.master.current_language]['page2']['totle']} {total} {LANGUAGES[self.master.current_language]['page2']['items']}", font=self.master.FONT_BODY)


    def paste_url(self):
        try:
            clipboard = self.clipboard_get()
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, clipboard)
        except Exception as e:
            log_and_show_error(f"Failed to paste URL from clipboard: {e}", self.master)

    def update_resolution_options(self):
        if self.format_var.get() == "mp3":
            # 當選擇 mp3 時，提供預設的音訊品質選項
            new_options = ["320kbps", "256kbps", "192kbps", "128kbps", "64kbps"]
            self.resolution_combobox.configure(values=new_options)
            # 例如預設使用最高品質
            self.resolution_combobox.set("320kbps")
        else:
            # 當選擇 mp4 時，使用影片解析度選項
            new_options = ["4320p","2160p", "1440p", "1080p", "720p", "480p", "360p", "240p", "144p"]
            self.resolution_combobox.configure(values=new_options)
            self.resolution_combobox.set("1080p")

    def add_playlist_item(self):
        """解析播放清單 URL，並將解析到的影片資料插入表格與內部清單中，使用線程執行並禁用提交按鈕"""
        self.submit_btn.configure(state="disabled")
        url = self.url_entry.get().strip()
        if not url:
            self.submit_btn.configure(state="normal")
            return
        # 檢查是否為播放清單 URL
        if "list=" not in url:
            log_and_show_error("Invalid playlist URL", self.master)
            self.submit_btn.configure(state="normal")
            return

        self.playlist_stop_event = threading.Event()
        self.playlist_thread = None

        def task():
            try:
                items = parse_playlist(url, self.resolution_combobox.get(), self.format_var.get(), self.master.cookies_path)
                if self.playlist_stop_event.is_set():
                    # 被用戶終止，不更新 UI
                    return
                if not items:
                    log_and_show_error("Failed to parse playlist or no videos found!", self.master)
                    # 回到主執行緒重新啟用提交按鈕
                    self.master.after(0, lambda: self.submit_btn.configure(state="normal"))
                    return
                # 將解析到的資料加入內部清單
                for item in items:
                    self.playlist_items.append(item)
                # 在主執行緒中更新 UI（因為 Tkinter 介面更新必須在主執行緒中進行）
                def update_ui():
                    for item in items:
                        self.table.add_row([item["title"], item["resolution"], item["format"], item["url"]])
                    self.update_total_label()
                    self.update_table_header()
                    self.submit_btn.configure(state="normal")
                self.master.after(0, update_ui)
            except Exception as e:
                if not self.playlist_stop_event.is_set():
                    log_and_show_error(f"Failed to parse playlist: {e}", self.master)
                self.master.after(0, lambda: self.submit_btn.configure(state="normal"))

        def ask_cancel():
            print(self.playlist_thread)
            if self.playlist_thread.is_alive():
                result = messagebox.askyesno(
                    LANGUAGES[self.master.current_language]["page2"]["timeout_title"],
                    LANGUAGES[self.master.current_language]["page2"]["timeout_message"]
                )
                if result:
                    self.playlist_stop_event.set()
                    self.submit_btn.configure(state="normal")

        self.playlist_thread = threading.Thread(target=task)
        self.playlist_thread.start()
        # 10秒後詢問是否終止
        self.master.after(10000, ask_cancel)


    def get_selected_rows(self):
        selected_rows = []
        for i in range(self.table.rows):
            # 依照你的設計，這邊只要檢查每行的第 1 欄是否 == hover_color 即可
            if self.table.frame[i, 1].cget("fg_color") == self.table.hover_color:
                selected_rows.append(i)
        return selected_rows

    def select_all_rows(self):
        """ 遍歷表格中除第一列（表頭）以外的所有列，呼叫 select_row() 使其選取 """
        selected_rows = self.get_selected_rows()
        if len(selected_rows) == self.table.rows - 1:
            # 如果所有列都已選取，則取消選取所有列
            for row in range(1, self.table.rows):
                self.table.deselect_row(row)
        else:
            # 假設 self.table.rows 回傳表格總列數，且第0列為表頭
            for row in range(1, self.table.rows):
                self.table.select_row(row)

    def delete_selected_rows(self):
        """刪除表格中選取的列，並從內部清單中移除"""
        selected = self.get_selected_rows()  # 假設此方法回傳選取列索引列表
        for index in sorted(selected, reverse=True):
            self.table.delete_row(index)
            del self.playlist_items[index-1]
        self.update_total_label()

    def on_cell_click(self, cell_data):
        """
        當使用者點擊儲存格時呼叫。
        cell_data 格式：
        {
            "row": <列號>,
            "column": <欄號>,
            "value": <該儲存格文字內容>,
            "args": <其他設定參數>
        }
        """
        row_index = cell_data["row"]
        if row_index == 0:
            # 如果點擊的是表頭，則不做任何事
            return
        # 檢查該 row 是否已經被選取：只要比對第二欄的 fg_color 是不是 hover_color
        is_selected = (self.table.frame[row_index, 1].cget("fg_color") == self.table.hover_color)

        if is_selected:
            # 如果已選取，則切換成「取消」
            self.table.deselect_row(row_index)
        else:
            # 如果還沒被選取，就「選取」
            self.table.select_row(row_index)

    def change_download_path(self):
        """變更下載位置"""
        path = filedialog.askdirectory()
        if path:
            self.download_path = path
            self.master.config["download_path"] = path
            save_config(self.master.config)

            self.download_path_textbox.configure(state="normal")
            self.download_path_textbox.delete("0.0", "end")
            self.download_path_textbox.insert("0.0", f"{LANGUAGES[self.master.current_language]['page2']['download_path_label']} {self.download_path}")
            self.download_path_textbox.configure(state="disabled")

    def update_progress(self, progress):
        lang = self.master.current_language
        if progress != -1:
            percent = int(progress * 100)
            # 使用 after 確保 GUI 更新在主執行緒中執行
            self.master.after(0, lambda: self.progress_bar.set(progress))
            self.master.after(0, lambda: self.progress_bar_label.configure(text=f"{LANGUAGES[self.master.current_language]['page2']['Processing']} {percent}%", font=self.master.FONT_BODY))
        else:
            self.master.after(100, lambda: self.progress_bar.set(progress))
            self.master.after(0, lambda: self.progress_bar_label.configure(text=LANGUAGES[self.master.current_language]['page2']['Processing_completed'], font=self.master.FONT_BODY))
    
    def download_playlist(self):
        """
        使用 ThreadPoolExecutor 控制同時線程數來多線程下載播放清單中所有影片，
        並根據已完成影片數更新進度條。整個流程放入獨立線程中以免阻塞主線程。
        """
        self.download_button.configure(state="disabled")
        total = len(self.playlist_items)
        if total == 0:
            self.download_button.configure(state="normal")
            return

        def download_item(item, idx):
            output_file = download_video_audio_playlist_with_retry(
                item["url"],
                item["resolution"],
                self.master.download_path,
                item["format"],
                self.master.cookies_path
            )
            return idx, output_file

        def thread_func():
            completed = 0
            max_threads = 4  # 同時最多執行 4 個下載任務
            self.master.after(0, lambda: self.update_progress(0))
            with ThreadPoolExecutor(max_workers=max_threads) as executor:
                futures = [executor.submit(download_item, item, idx)
                        for idx, item in enumerate(self.playlist_items)]
                for future in as_completed(futures):
                    try:
                        idx, output_file = future.result()
                    except Exception as e:
                        log_and_show_error(f"Download failed: {e}", self.master)
                        continue
                    completed += 1
                    progress = completed / total
                    # 更新進度條必須在主線程中執行
                    self.master.after(0, lambda: self.update_progress(progress))
                    logger.info(f"Video {idx} downloaded: {output_file}")
            # 所有任務完成後，回到主線程中重新啟用按鈕與設定進度條
            self.master.after(0, lambda: self.download_button.configure(state="normal"))
            self.master.after(0, lambda: self.update_progress(-1))
            self.master.after(0, lambda: messagebox.showinfo(
                LANGUAGES[self.master.current_language]['page2']["completed"],  # 標題
                f"{LANGUAGES[self.master.current_language]['page2']['completed']}: {completed}\n"
                f"{LANGUAGES[self.master.current_language]['page2']['failed']}: {total - completed}"
            ))
            logger.info("All videos downloaded")

        # 將整個 ThreadPoolExecutor 流程放到獨立線程中執行，避免阻塞主線程
        threading.Thread(target=thread_func).start()

    def update_bg_image(self):
        bg_image_path = self.master.bg_image_path 
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                img = Image.open(bg_image_path)
                img = ImageOps.fit(img, (1280, 720), Image.LANCZOS) # 圖片比例不變但填滿
                self.bg_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.bg_label.configure(text="", image=self.bg_image)
            except Exception as e:
                log_and_show_error(f"Background image load failed: {e}", self.master)
                self.bg_label.configure(text="", image=None)

    def update_ad_area(self):
        ad_image_path = self.master.config.get("ad_image", "")
        if ad_image_path and os.path.exists(ad_image_path):
            try:
                img = Image.open(ad_image_path)
                img = ImageOps.contain(img, (640, 480))
                self.ad_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.ad_label.configure(image=self.ad_image, text="")
            except Exception as e:
                log_and_show_error(f"AD image load failed: {e}", self.master)
                self.ad_label.configure(text=LANGUAGES[self.master.current_language]["page2"]["no_ad_label"], image=None, font=self.master.FONT_BODY)
        else:
            self.ad_label.configure(text=LANGUAGES[self.master.current_language]["page2"]["no_ad_label"], image=None, font=self.master.FONT_BODY)

    def update_text(self):
        lang = self.master.current_language
        text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"

        # 頂部
        self.back_btn.configure(text_color=text_color, font=self.master.FONT_BUTTON)
        self.page_title.configure(text=LANGUAGES[lang]["page2"]["page2_title"], text_color=text_color, font=self.master.FONT_TITLE)
        self.edit_btn.configure(text_color=text_color, font=self.master.FONT_BUTTON)

        # 左側表格
        self.select_all_btn.configure(text=LANGUAGES[lang]["page2"]["select_all"], font=self.master.FONT_BUTTON)
        self.delete_btn.configure(text=LANGUAGES[lang]["page2"]["delete_selected"], font=self.master.FONT_BUTTON)

        # 右側 URL 輸入與下載位置
        self.url_entry.configure(placeholder_text=LANGUAGES[lang]["page2"]["playlist_url_label"], font=self.master.FONT_BODY)
        self.submit_btn.configure(text=LANGUAGES[lang]["page2"]["submit_button"], font=self.master.FONT_BUTTON)
        self.resolution_combobox.configure(font=self.master.FONT_BODY)
        self.download_path_textbox.configure(state="normal", font=self.master.FONT_BODY)
        self.download_path_textbox.delete("0.0", "end")
        self.download_path_textbox.insert("0.0", f"{LANGUAGES[lang]['page2']['download_path_label']} {self.download_path}")
        self.download_path_textbox.configure(state="disabled")
        self.change_path_button.configure(text=LANGUAGES[lang]["page2"]["browse_button"], font=self.master.FONT_BUTTON)
        
        # 底部
        self.progress_bar_label.configure(text=LANGUAGES[lang]["page2"]["progress_ready"], font=self.master.FONT_BODY)
        self.download_button.configure(text=LANGUAGES[lang]["page2"]["download_button"], font=self.master.FONT_BUTTON)

    def update_frame_tranparency(self):
        # 根據主題設定物件透明度
        # 注意! 若frame已經消除透明度，內部物件則不需要再設定
        background_color = "#FFFFFF" if self.master.config.get("theme", "Dark") == "Light" else "#000001"
        pywinstyles.set_opacity(self.frame_top, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_left_first, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_first_right, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_second_right, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_bottom, value=self.master.transparency, color=background_color)

    def update_all_objects(self):
        self.update_bg_image()  # 更新背景圖片
        self.update_text() # 初始化文字
        self.update_table_header() # 更新表格表頭
        self.update_total_label() # 更新總筆數標籤
        self.update_resolution_options() # 初始化解析度選項
        self.update_ad_area() # 初始化廣告區
        self.update_frame_tranparency() # 初始化透明度

class Page3(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        
        # 設定區域佈局
        self.grid_rowconfigure(0, weight=1)  
        self.grid_rowconfigure(1, weight=8)   
        self.grid_rowconfigure(2, weight=1)   
        self.grid_columnconfigure(0, weight=6)
        self.grid_columnconfigure(1, weight=4)

        text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"
        
        # Background image initialization
        self.bg_label = ctk.CTkLabel(self, text="")
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        # ---------- 頂部工具列 ----------
        self.frame_top = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_top.grid(row=0, column=0, sticky="nsew", columnspan=2, padx=5)
        self.frame_top.grid_rowconfigure(0, weight=1)
        self.frame_top.grid_columnconfigure((0,2), weight=1)
        self.frame_top.grid_columnconfigure(1, weight=8)

        self.back_icon = ctk.CTkButton(
            self.frame_top, 
            text="\U00002190", 
            width=50, 
            command=lambda: master.show_frame(HomePage),
            fg_color="transparent",
            text_color=text_color, 
            border_width=0, 
            border_spacing=0, 
            corner_radius=2
        )
        self.back_icon.grid(row=0, column=0, padx=5, sticky="w")

        self.page_title = ctk.CTkLabel(self.frame_top)
        self.page_title.grid(row=0, column=1, padx=5, sticky="ew")

        self.edit_icon = ctk.CTkButton(
            self.frame_top, 
            text="\U00002699", 
            width=50, 
            command=master.open_Setting,
            fg_color="transparent", 
            text_color=text_color,
            border_width=0, 
            border_spacing=0, 
            corner_radius=2
        )
        self.edit_icon.grid(row=0, column=2, padx=5, sticky="e")

        # ---------- 主區域 ----------
        self.frame_main = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),  
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_main.grid(row=1, column=0, sticky="nsew", padx=10, pady=10, columnspan=2)
        self.frame_main.grid_rowconfigure((0,1), weight=1)
        self.frame_main.grid_columnconfigure(0, weight=4)
        self.frame_main.grid_columnconfigure(1, weight=6)

        # ---------- 左上：檔案選擇與顯示 ----------
        self.frame_left = ctk.CTkFrame(
            self.frame_main,
            bg_color=("#FFFFFF", "#000001"),  
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_left.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.frame_left.grid_columnconfigure(0, weight=3)
        self.frame_left.grid_columnconfigure(1, weight=3)
        self.frame_left.grid_columnconfigure(2, weight=4)

        self.file_label = ctk.CTkLabel(self.frame_left)
        self.file_label.grid(row=0, column=0, padx=5, pady=5, columnspan=2, sticky="w")

        self.selected_file = ctk.StringVar(value="")
        self.file_display = ctk.CTkEntry(self.frame_left, textvariable=self.selected_file, state="disabled")
        self.file_display.grid(row=1, column=0, padx=5, pady=5, columnspan=2, sticky="ew")

        self.file_button = ctk.CTkButton(self.frame_left, text="Browse", command=self.browse_file)
        self.file_button.grid(row=1, column=2, padx=5, pady=5, sticky="w")

        self.converted_file_label = ctk.CTkLabel(self.frame_left)
        self.converted_file_label.grid(row=2, column=0, padx=5, pady=5, columnspan=2, sticky="w")

        self.converted_file_display = ctk.CTkEntry(self.frame_left, state="disabled")
        self.converted_file_display.grid(row=3, column=0, padx=5, pady=5, columnspan=2, sticky="ew")

        self.converted_file_button = ctk.CTkButton(self.frame_left, text="Open", command=self.open_converted_file)
        self.converted_file_button.grid(row=3, column=2, padx=5, pady=5, columnspan=2, sticky="w")

        # 新增起始與結束時間輸入
        self.start_time_var = ctk.StringVar(value="00:00:00")
        self.end_time_var = ctk.StringVar(value="")  # 待檔案選擇後更新

        self.start_time_label = ctk.CTkLabel(self.frame_left)
        self.start_time_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        self.start_time_entry = ctk.CTkEntry(self.frame_left, textvariable=self.start_time_var)
        self.start_time_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        self.end_time_label = ctk.CTkLabel(self.frame_left)
        self.end_time_label.grid(row=5, column=0, padx=5, pady=5, sticky="w")
        
        self.end_time_entry = ctk.CTkEntry(self.frame_left, textvariable=self.end_time_var)
        self.end_time_entry.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

        # ---------- 左下：轉換參數設定區 ----------
        self.frame_left_second = ctk.CTkFrame(
            self.frame_main,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"),
        )
        self.frame_left_second.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.frame_left_second.grid_columnconfigure(0, weight=1)
        self.frame_left_second.grid_columnconfigure(1, weight=1)

        # 轉換器類型選擇
        self.converter_type = ctk.StringVar(value="video")
        self.video_radio = ctk.CTkRadioButton(self.frame_left_second, variable=self.converter_type, value="video", command=self.update_parameters)
        self.audio_radio = ctk.CTkRadioButton(self.frame_left_second, variable=self.converter_type, value="audio", command=self.update_parameters)
        self.video_radio.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.audio_radio.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # 參數設定：根據類型顯示不同的參數
        self.param_label = ctk.CTkLabel(self.frame_left_second)
        self.param_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.param_combobox = ctk.CTkComboBox(self.frame_left_second, values=[])
        self.param_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # 目標格式設定
        self.target_format_label = ctk.CTkLabel(self.frame_left_second)
        self.target_format_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.target_format_combobox = ctk.CTkComboBox(self.frame_left_second, values=[])
        self.target_format_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        # 綁定目標格式選擇事件，用於更新 audio 的 bitrate 選項
        self.target_format_combobox.bind("<<ComboboxSelected>>", self.on_audio_format_change)
        
        # 新增轉碼器選項（僅在 video 模式顯示）
        self.video_transcoder_label = ctk.CTkLabel(self.frame_left_second)
        self.video_transcoder_combobox = ctk.CTkComboBox(self.frame_left_second, values=["Default", "libx264", "libx265"])
        self.audio_transcoder_label = ctk.CTkLabel(self.frame_left_second)
        self.audio_transcoder_combobox = ctk.CTkComboBox(self.frame_left_second, values=["Default", "aac", "mp3"])

        # ---------- 右側: 廣告區 ----------
        self.frame_right = ctk.CTkFrame(
            self.frame_main,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"),
        )
        self.frame_right.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=5, pady=5)
        self.frame_right.grid_propagate(False)  # 固定 frame 尺寸
        self.frame_right.grid_rowconfigure(0, weight=1)
        self.frame_right.grid_columnconfigure(0, weight=1)

        self.ad_label = ctk.CTkLabel(self.frame_right, text='')
        self.ad_label.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)

        # ---------- 底部進度區 ----------
        self.frame_bottom = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_bottom.grid_rowconfigure(0, weight=4)
        self.frame_bottom.grid_rowconfigure(1, weight=6)
        self.frame_bottom.grid_columnconfigure(0, weight=7)
        self.frame_bottom.grid_columnconfigure(1, weight=3)

        self.progress_label = ctk.CTkLabel(self.frame_bottom)
        self.progress_label.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self.frame_bottom)
        self.progress_bar.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.progress_bar.set(0)

        self.frame_bottom.grid(row=2, column=0, sticky="nsew", padx=10, pady=10, columnspan=2)
        self.convert_button = ctk.CTkButton(self.frame_bottom, command=self.start_conversion)
        self.convert_button.grid(row=1, column=1, padx=5, pady=5)

        self.update_all_objects()

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Video Files", "*.mp4 *.webm *.mkv *.mov"), ("Audio Files", "*.mp3 *.wav *.flac *.ogg")]
        )
        if file_path:
            self.selected_file.set(file_path)
            self.end_time_var.set(get_media_duration(file_path))
    
    def open_converted_file(self):
        file_path = self.converted_file_display.get()
        if file_path:
            # 開啟轉換檔案所在的資料夾
            os.startfile(os.path.dirname(file_path))

    def update_parameters(self):
        """根據 converter_type 更新參數與目標格式設定"""
        lang = self.master.current_language
        if self.converter_type.get() == "video":
            self.param_label.configure(text=LANGUAGES[lang]["page3"]["resolution_label"], font=self.master.FONT_BODY)
            self.param_combobox.configure(values=["Original resolution", "7680x4320", "3840x2160", "2560x1440", "1920x1080", "1280x720", "854x480", "640x360", "426x240", "256x144"], font=self.master.FONT_BODY)
            self.param_combobox.set("Original resolution")
            self.target_format_label.configure(text=LANGUAGES[lang]["page3"]["target_format_label"], font=self.master.FONT_BODY)
            self.target_format_combobox.configure(values=["mp4", "webm", "mkv", "mov"], font=self.master.FONT_BODY)
            self.target_format_combobox.set("mp4")
            # 設定 target_format_combobox 的 callback，依據選擇動態更新轉碼器選項
            self.target_format_combobox.configure(command=self.on_video_format_change)
            # 顯示轉碼器選項並預設更新
            self.video_transcoder_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
            self.video_transcoder_combobox.grid(row=3, column=1, padx=5, pady=5, sticky="w")
            self.audio_transcoder_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")
            self.audio_transcoder_combobox.grid(row=4, column=1, padx=5, pady=5, sticky="w")
            self.on_video_format_change(None)
        else:
            self.param_label.configure(text=LANGUAGES[lang]["page3"]["bit_rate_label"], font=self.master.FONT_BODY)
            self.target_format_label.configure(text=LANGUAGES[lang]["page3"]["target_format_label"], font=self.master.FONT_BODY)
            self.target_format_combobox.configure(values=["mp3", "wav", "flac", "ogg"], font=self.master.FONT_BODY)
            self.target_format_combobox.set("mp3")
            # 設定 target_format_combobox 的 callback，更新 audio bitrate 選項
            self.target_format_combobox.configure(command=self.on_audio_format_change)
            self.on_audio_format_change(None)
            # 隱藏轉碼器選項
            self.video_transcoder_label.grid_remove()
            self.video_transcoder_combobox.grid_remove()
            self.audio_transcoder_label.grid_remove()
            self.audio_transcoder_combobox.grid_remove()

    def on_video_format_change(self, value):
        """根據 video 目標格式動態更新視訊與音訊轉碼器選項"""
        target_format = self.target_format_combobox.get().lower()
        video_opts = {
            "mp4": ["Default", "libx264", "libx265"],
            "webm": ["Default", "libvpx", "libvpx-vp9"],
            "mkv": ["Default", "libx264", "libx265"],
            "mov": ["Default", "prores", "libx264"]
        }
        audio_opts = {
            "mp4": ["Default", "aac", "mp3"],
            "webm": ["Default", "opus", "libvorbis"],
            "mkv": ["Default", "aac", "mp3"],
            "mov": ["Default", "aac", "mp3"]
        }
        v_options = video_opts.get(target_format, ["Default"])
        a_options = audio_opts.get(target_format, ["Default"])
        self.video_transcoder_combobox.configure(values=v_options)
        self.audio_transcoder_combobox.configure(values=a_options)
        self.video_transcoder_combobox.set(v_options[0])
        self.audio_transcoder_combobox.set(a_options[0])

    def on_audio_format_change(self, value):
        """依據音訊目標格式更新 bitrate 參數選項"""
        target_format = self.target_format_combobox.get().lower()
        if target_format == "wav":
            self.param_combobox.configure(values=["20kHz","44.1kHz", "48kHz", "96kHz"])
            self.param_combobox.set("44.1kHz")
        elif target_format == "flac":
            self.param_combobox.configure(values=["320kbps"])
            self.param_combobox.set("320kbps")
        else:
            self.param_combobox.configure(values=["64kbps", "128kbps", "192kbps", "320kbps"])
            self.param_combobox.set("128kbps")

    def update_progress(self, progress):
        lang= self.master.current_language 
        if progress != -1:
            percent = int(progress * 100)
            # 使用 after 確保 GUI 更新在主執行緒中執行
            self.master.after(0, lambda: self.progress_bar.set(progress))
            self.master.after(0, lambda: self.progress_label.configure(text=f"{LANGUAGES[lang]['page3']['converting']}: {percent}%", font=self.master.FONT_BODY))
        else:
            self.master.after(100, lambda: self.progress_bar.set(progress))
            self.master.after(0, lambda: self.progress_label.configure(text=f"{LANGUAGES[lang]['page3']['converting_completed']}", font=self.master.FONT_BODY))

    def start_conversion(self):
        lang= self.master.current_language 
        file_path = self.selected_file.get()
        if not file_path:
            log_and_show_error("No file selected!", self.master)
            return
        conv_type = self.converter_type.get()
        param = self.param_combobox.get()
        target_format = self.target_format_combobox.get()
        start_time = self.start_time_var.get()
        end_time = self.end_time_var.get()
        self.convert_button.configure(state="disabled")
        self.progress_label.configure(text=LANGUAGES[self.master.current_language]["page3"]["converting"], font=self.master.FONT_BODY)

        def conversion_task():
            if end_time and end_time.strip():
                    start_sec = time_to_seconds(start_time) if start_time and start_time != "00:00:00" else 0
                    end_sec = time_to_seconds(end_time)
                    conversion_duration = end_sec - start_sec if end_sec > start_sec else 0
            else:
                duration_str = get_media_duration(file_path)
                conversion_duration = time_to_seconds(duration_str) if duration_str else 0

            if conv_type == "video":
                video_transcoder = self.video_transcoder_combobox.get()
                audio_transcoder = self.audio_transcoder_combobox.get()
                output = convert_video(
                file_path, param, target_format, start_time, conversion_duration,
                video_transcoder, audio_transcoder, self.update_progress
            )
            else:
                output = convert_audio(file_path, param, target_format, start_time, conversion_duration, self.update_progress)
           
            # 使用 after 確保 GUI 更新在主執行緒中執行 
            self.master.after(0, lambda: self.converted_file_display.configure(state="normal"))
            self.master.after(0, lambda: self.converted_file_display.delete(0, "end"))
            self.master.after(0, lambda: self.converted_file_display.insert(0, output))
            self.master.after(0, lambda: self.converted_file_display.configure(state="disabled"))
            self.master.after(0, lambda: self.convert_button.configure(state="normal"))
            self.master.after(0, lambda: self.progress_label.configure(text=LANGUAGES[self.master.current_language]["page3"]["converting_completed"]))
            if output:
                self.master.after(0, lambda: messagebox.showinfo(
                    LANGUAGES[lang]["page3"]["convert_success_title"],
                    LANGUAGES[lang]["page3"]["convert_success_message"].format(output)
                ))
        threading.Thread(target=conversion_task).start()

        # 重製進度條
        self.progress_bar.set(0.0)

    def update_bg_image(self):
        bg_image_path = self.master.bg_image_path 
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                img = Image.open(bg_image_path)
                img = ImageOps.fit(img, (1280, 720), Image.LANCZOS) # 圖片比例不變但填滿
                self.bg_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.bg_label.configure(text="", image=self.bg_image)
            except Exception as e:
                log_and_show_error(f"Background image load failed: {e}", self.master)
                self.bg_label.configure(text="", image=None)

    def update_ad_area(self):
        ad_image_path = self.master.config.get("ad_image", "")
        if ad_image_path and os.path.exists(ad_image_path):
            try:
                img = Image.open(ad_image_path)
                img = ImageOps.contain(img, (640, 480))
                self.ad_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.ad_label.configure(image=self.ad_image, text="")
            except Exception as e:
                log_and_show_error(f"AD image load failed: {e}", self.master)
                self.ad_label.configure(text=LANGUAGES[self.master.current_language]["page3"]["no_ad_label"], image=None, font=self.master.FONT_BODY)
        else:
            self.ad_label.configure(text=LANGUAGES[self.master.current_language]["page3"]["no_ad_label"], image=None, font=self.master.FONT_BODY)

    def update_text(self):
        lang = self.master.current_language
        text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"
        self.back_icon.configure(text_color=text_color, font=self.master.FONT_BUTTON)
        self.page_title.configure(text=LANGUAGES[lang]["page3"]["page3_title"], text_color=text_color, font=self.master.FONT_TITLE)
        self.edit_icon.configure(text_color=text_color, font=self.master.FONT_BUTTON)
        self.file_label.configure(text=LANGUAGES[lang]["page3"]["select_files_label"], font=self.master.FONT_BODY)
        self.file_button.configure(text=LANGUAGES[lang]["page3"]["browse_button"], font=self.master.FONT_BUTTON)
        self.converted_file_label.configure(text=LANGUAGES[lang]["page3"]["converted_files_label"], font=self.master.FONT_BODY)
        self.converted_file_button.configure(text=LANGUAGES[lang]["page3"]["open_button"], font=self.master.FONT_BUTTON)
        self.start_time_label.configure(text=LANGUAGES[lang]["page3"]["start_time_label"], font=self.master.FONT_BODY)
        self.end_time_label.configure(text=LANGUAGES[lang]["page3"]["end_time_label"], font=self.master.FONT_BODY)
        self.video_radio.configure(text=LANGUAGES[lang]["page3"]["video_radio"], font=self.master.FONT_BODY)
        self.audio_radio.configure(text=LANGUAGES[lang]["page3"]["audio_radio"], font=self.master.FONT_BODY)
        self.video_transcoder_label.configure(text=LANGUAGES[lang]["page3"]["video_transcoder_label"], font=self.master.FONT_BODY)
        self.audio_transcoder_label.configure(text=LANGUAGES[lang]["page3"]["audio_transcoder_label"], font=self.master.FONT_BODY)
        self.convert_button.configure(text=LANGUAGES[lang]["page3"]["convert_button"], font=self.master.FONT_BUTTON)
        self.progress_label.configure(text=LANGUAGES[lang]["page3"]["progress_ready"], font=self.master.FONT_BODY)

        self.param_combobox.configure(font=self.master.FONT_BODY)
        self.target_format_combobox.configure(font=self.master.FONT_BODY)
        self.video_transcoder_combobox.configure(font=self.master.FONT_BODY)
        self.audio_transcoder_combobox.configure(font=self.master.FONT_BODY)
        self.update_parameters()  # 初始化參數設定

    def update_frame_tranparency(self):
        # 根據主題設定物件透明度
        # 注意! 若frame已經消除透明度，內部物件則不需要再設定
        background_color = "#FFFFFF" if self.master.config.get("theme", "Dark") == "Light" else "#000001"
        pywinstyles.set_opacity(self.frame_top, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_main, value=self.master.transparency, color=background_color)

        pywinstyles.set_opacity(self.frame_bottom, value=self.master.transparency, color=background_color)

    def update_all_objects(self):
        self.update_bg_image()
        self.update_text()
        self.update_parameters()  # 更新參數設定
        self.update_ad_area()  # 更新廣告區
        self.update_frame_tranparency()  # 更新透明度


class Page4(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.download_path = self.master.config.get("download_path") or os.getcwd()

        # 設定 Grid 權重
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1) 
        self.grid_rowconfigure(1, weight=4) 
        self.grid_rowconfigure(2, weight=1)

        text_color = "black" if self.master.config.get("theme", "Dark") == "Light" else "white"

        # Background image initialization
        self.bg_label = ctk.CTkLabel(self, text="")
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        # ====== 頁面頂部 ======
        self.frame_top = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_top.grid(row=0, column=0, sticky="nsew", columnspan=2, padx=5)
        self.frame_top.grid_rowconfigure(0, weight=1)
        self.frame_top.grid_columnconfigure((0,2), weight=1)
        self.frame_top.grid_columnconfigure(1, weight=8)

        self.back_icon = ctk.CTkButton(
            self.frame_top, 
            text="\U00002190", 
            width=50, 
            command=lambda: master.show_frame(HomePage),
            bg_color=("#FFFFFF", "#000001"), 
            fg_color="transparent",
            text_color=text_color, 
            border_width=0, 
        )
        self.back_icon.grid(row=0, column=0, padx=5, sticky="w")

        self.page_title = ctk.CTkLabel(
            self.frame_top,
            bg_color=("#FFFFFF", "#000001"),  
        )
        self.page_title.grid(row=0, column=1, padx=5, sticky="ew")

        self.edit_icon = ctk.CTkButton(
            self.frame_top, 
            text="\U00002699", 
            width=50, 
            command=master.open_Setting,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color="transparent", 
            text_color=text_color,
            border_width=0, 
        )
        self.edit_icon.grid(row=0, column=2, padx=5, sticky="e")

        # ====== 左半 ======
        self.frame_left = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_left.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.frame_left.grid_rowconfigure(0, weight=1)
        self.frame_left.grid_columnconfigure(0, weight=1)

        # 文字輸入框
        self.text_label = ctk.CTkTextbox(
            self.frame_left, 
            corner_radius=20,
            activate_scrollbars=True,
        )
        self.text_label.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # ====== 右半 ======
        self.frame_right = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"), 
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_right.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        self.frame_right.grid_rowconfigure((0,1,2,3,4,5,6,7), weight=1)
        self.frame_right.grid_columnconfigure((0,1,2,3), weight=1)

        # 下載位置
        self.download_path_textbox = ctk.CTkTextbox(self.frame_right, height=30, activate_scrollbars=False)
        self.download_path_textbox.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        self.change_path_button = ctk.CTkButton(self.frame_right, command=self.change_download_path)
        self.change_path_button.grid(row=0, column=3, padx=10, pady=10, sticky="e")

        # 語言對應表
        self.language_map = {
            "zh-TW": "中文 (TW)",
            "zh-CN": "中文 (CN)",
            "en-US": "English (US)",
            "ja-JP": "日本語",
            "ko-KR": "한국어",
            "es-ES": "Español (ES)",
            # 可依需求擴充
        }
        self.language_keys = list(self.language_map.keys())
        self.language_values = list(self.language_map.values())

        # 語言選擇
        self.language_label = ctk.CTkLabel(self.frame_right)
        self.language_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.language_combobox = ctk.CTkComboBox(self.frame_right, values=self.language_values, command=self.on_language_change)
        self.language_combobox.grid(row=1, column=3, padx=10, pady=10, sticky="ew")
        self.language_combobox.set(self.language_values[0])

        # 語音選擇
        self.voice_label = ctk.CTkLabel(self.frame_right)
        self.voice_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.voice_combobox = ctk.CTkComboBox(self.frame_right)
        self.voice_combobox.grid(row=2, column=3, padx=10, pady=10, sticky="ew")

        # 下載檔案格式選擇
        self.format_label = ctk.CTkLabel(self.frame_right)
        self.format_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.format_combobox = ctk.CTkComboBox(self.frame_right, values=["mp3", "wav", "ogg", "flac"])
        self.format_combobox.grid(row=3, column=3, padx=10, pady=10, sticky="ew")

        # 語速設定
        self.speed_label = ctk.CTkLabel(self.frame_right)
        self.speed_label.grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.speed_combobox = ctk.CTkComboBox(
            self.frame_right,
            values=["-50%","-20%", "-10%", "+0%", "+10%", "+20%", "+30%", "+40%", "+50%", "+100%"]
        )
        self.speed_combobox.grid(row=4, column=3, padx=10, pady=10, sticky="ew")
        self.speed_combobox.set("+0%")  # 預設語速

        # 音量設定
        self.volume_label = ctk.CTkLabel(self.frame_right)
        self.volume_label.grid(row=5, column=0, padx=10, pady=10, sticky="w")
        self.volume_combobox = ctk.CTkComboBox(
            self.frame_right,
            values=["-200%","-150%","-100%","-50%","-25%","+0%", "+25%", "+50%", "+75%", "+100%", "+125%", "+150%", "+175%", "+200%"]
        )
        self.volume_combobox.grid(row=5, column=3, padx=10, pady=10, sticky="ew")
        self.volume_combobox.set("+0%")  # 預設音量

        # 音高設定
        self.pitch_label = ctk.CTkLabel(self.frame_right)
        self.pitch_label.grid(row=6, column=0, padx=10, pady=10, sticky="w")
        self.pitch_combobox = ctk.CTkComboBox(
            self.frame_right,
            values=["-200Hz", "-100Hz", "-50Hz", "-25Hz", "+0Hz", "+25Hz", "+50Hz", "+75Hz", "+100Hz", "+200Hz"]
        )
        self.pitch_combobox.grid(row=6, column=3, padx=10, pady=10, sticky="ew")
        self.pitch_combobox.set("+0Hz")  # 預設音高   

        # ====== 底部 Frame ======
        self.frame_bottom = ctk.CTkFrame(
            self,
            bg_color=("#FFFFFF", "#000001"),
            fg_color=("#FFFFFF", "#000001"), 
        )
        self.frame_bottom.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        self.frame_bottom.grid_rowconfigure(0, weight=4)
        self.frame_bottom.grid_rowconfigure(1, weight=6)
        self.frame_bottom.grid_columnconfigure(0, weight=7)
        self.frame_bottom.grid_columnconfigure(1, weight=3)

        self.progress_bar_label = ctk.CTkLabel(self.frame_bottom)
        self.progress_bar_label.grid(row=0, column=0, sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(self.frame_bottom)
        self.progress_bar.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.progress_bar.set(0)

        self.convert_button = ctk.CTkButton(self.frame_bottom, command=self.start_conversion)
        self.convert_button.grid(row=1, column=1, pady=5)

        self.update_all_objects()

    def on_language_change(self, selected_value):
        lang_key = None
        for k, v in self.language_map.items():
            if v == selected_value:
                lang_key = k
                break
        if lang_key:
            self.update_voice_options(lang_key)

    def update_voice_options(self, lang_key=None):
        """根據語言 key 更新語音選項"""
        try:
            if lang_key is None:
                # 預設用第一個語言
                lang_key = self.language_keys[0]
            # 取得所有 voices
            voices = asyncio.run(fetch_voice_names())
            # 過濾出符合語言的 voices
            filtered = [v for v in voices if v.startswith(lang_key)]
            if not filtered:
                filtered = ["Default"]
            self.voice_combobox.configure(values=filtered)
            self.voice_combobox.set(filtered[0])
        except Exception as e:
            log_and_show_error(f"Failed to fetch voices: {e}", self.master)
            self.voice_combobox.configure(values=["Default"])
            self.voice_combobox.set("Default")

    def change_download_path(self):
        """變更下載位置"""
        path = filedialog.askdirectory()
        if path:
            self.download_path = path
            self.master.config["download_path"] = path
            save_config(self.master.config)

            self.download_path_textbox.configure(state="normal")
            self.download_path_textbox.delete("0.0", "end")
            self.download_path_textbox.insert("0.0", f"{LANGUAGES[self.master.current_language]['page4']['download_path_label']} {self.download_path}")
            self.download_path_textbox.configure(state="disabled")

    def update_progress(self, progress, status_text=None):
        # 騙用戶而已，由於 edge-tts 的 Communicate.save() 是一次性寫檔，無法直接回報進度
        self.progress_bar.set(progress)
        if status_text is not None:
            self.progress_bar_label.configure(text=status_text, font=self.master.FONT_BODY)

    def start_conversion(self):
        def run_conversion():
            try:
                lang = self.master.current_language
                voice = self.voice_combobox.get()
                format_ = self.format_combobox.get().lower()
                download_path = self.download_path or os.getcwd()
                speed = self.speed_combobox.get()
                volume = self.volume_combobox.get()
                pitch = self.pitch_combobox.get()
                text = self.text_label.get("0.0", "end").strip()
                if not text:
                    log_and_show_error(LANGUAGES[lang]["page4"]["no_text_error"], self.master)
                    return

                # 開始進度
                self.master.after(0, lambda: self.update_progress(0.0, LANGUAGES[lang]["page4"]["converting"]))
                # 禁用轉換按鈕
                self.convert_button.configure(state="disabled")

                # 語音合成
                result = asyncio.run(convert_text_to_speech(
                    text=text,
                    voice=voice,
                    format=format_,
                    download_path=download_path,
                    speed=speed,
                    volume=volume,
                    pitch=pitch,
                ))

                # 完成進度
                self.master.after(0, lambda: self.update_progress(1.0, LANGUAGES[lang]["page4"]["converting_completed"]))
                if result:
                    self.master.after(0, lambda: messagebox.showinfo(
                        LANGUAGES[lang]["page4"]["convert_success_title"],
                        LANGUAGES[lang]["page4"]["convert_success_message"].format(result)
                    ))
                    
                # reset
                self.master.after(0, lambda: self.convert_button.configure(state="normal"))
                self.progress_bar_label.configure(text=LANGUAGES[lang]["page4"]["progress_ready"], font=self.master.FONT_BODY)
                self.master.after(0, lambda: self.update_progress(0.0, LANGUAGES[lang]["page4"]["converting"]))
            except Exception as e:
                log_and_show_error(f"Conversion failed: {e}", self.master)
                self.master.after(0, lambda: self.update_progress(0.0, LANGUAGES[self.master.current_language]["page4"]["progress_failed"]))
                self.master.after(0, lambda: self.convert_button.configure(state="normal"))
                
        threading.Thread(target=run_conversion).start()

    def update_bg_image(self):
        bg_image_path = self.master.bg_image_path 
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                img = Image.open(bg_image_path)
                img = ImageOps.fit(img, (1280, 720), Image.LANCZOS) # 圖片比例不變但填滿
                self.bg_image = CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
                self.bg_label.configure(text="", image=self.bg_image)
            except Exception as e:
                log_and_show_error(f"Background image load failed: {e}", self.master)
                self.bg_label.configure(text="", image=None)

    def update_frame_tranparency(self):
        # 根據主題設定物件透明度
        # 注意! 若frame已經消除透明度，內部物件則不需要再設定
        background_color = "#FFFFFF" if self.master.config.get("theme", "Dark") == "Light" else "#000001"
        pywinstyles.set_opacity(self.frame_top, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_left, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_right, value=self.master.transparency, color=background_color)
        pywinstyles.set_opacity(self.frame_bottom, value=self.master.transparency, color=background_color)

    def update_text(self):
        lang = self.master.current_language
        self.page_title.configure(text=LANGUAGES[lang]["page4"]["page4_title"], font=self.master.FONT_TITLE)
        self.language_label.configure(text=LANGUAGES[lang]["page4"]["language_label"], font=self.master.FONT_BODY)
        self.voice_label.configure(text=LANGUAGES[lang]["page4"]["voice_label"], font=self.master.FONT_BODY)
        self.format_label.configure(text=LANGUAGES[lang]["page4"]["format_label"], font=self.master.FONT_BODY)

        self.download_path_textbox.configure(state="normal", font=self.master.FONT_BODY)
        self.download_path_textbox.delete("0.0", "end")
        self.download_path_textbox.insert("0.0", f"{LANGUAGES[lang]['page4']['download_path_label']} {self.download_path}")
        self.download_path_textbox.configure(state="disabled")

        self.change_path_button.configure(text=LANGUAGES[lang]["page4"]["browse_button"], font=self.master.FONT_BUTTON)
        self.speed_label.configure(text=LANGUAGES[lang]["page4"]["speed_label"], font=self.master.FONT_BODY)
        self.volume_label.configure(text=LANGUAGES[lang]["page4"]["volume_label"], font=self.master.FONT_BODY)
        self.pitch_label.configure(text=LANGUAGES[lang]["page4"]["pitch_label"], font=self.master.FONT_BODY)
        self.progress_bar_label.configure(text=LANGUAGES[lang]["page4"]["progress_ready"], font=self.master.FONT_BODY)
        self.convert_button.configure(text=LANGUAGES[lang]["page4"]["convert_button"], font=self.master.FONT_BUTTON)

        self.language_combobox.configure(font=self.master.FONT_BODY)
        self.format_combobox.configure(font=self.master.FONT_BODY)
        self.speed_combobox.configure(font=self.master.FONT_BODY)
        self.volume_combobox.configure(font=self.master.FONT_BODY)
        self.pitch_combobox.configure(font=self.master.FONT_BODY)

    def update_all_objects(self):
        self.update_bg_image()
        self.update_frame_tranparency()
        self.update_text()
        self.update_voice_options(self.language_keys[0])  # 初始化語音選項
        

if __name__ == "__main__":
    config = load_config()
    ctk.set_default_color_theme(config["theme_color"]) # Themes: "blue" (standard), "green", "dark-blue"
    app = MainApp()
    app.mainloop()
