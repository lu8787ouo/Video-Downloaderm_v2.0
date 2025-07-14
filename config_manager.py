import os
import json

CONFIG_FILE = "config.json"
DEFAULT_THEME_COLOR = os.path.join(os.path.dirname(__file__), "assets/themes/SakuraPink.json")
DEFAULT_BG_IMAGE = os.path.join(os.path.dirname(__file__), "assets/background/sakura_background.png")
DEFAULT_CONFIG = {
    "theme": "Light",
    "theme_color": DEFAULT_THEME_COLOR, 
    "language": "zh-TW",
    "resolution": "1280x720",
    "download_path": os.getcwd(),
    "ad_image": "",
    "bg_image": DEFAULT_BG_IMAGE,
    "transparency": "1",
    "cookies": ""
}

def load_config():
    """讀取設定檔，若不存在則建立預設設定檔並回傳預設值"""
    if not os.path.exists(CONFIG_FILE):
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_config(config):
    """儲存設定檔"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
