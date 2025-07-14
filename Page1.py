import os
import re
import time
import functools
import subprocess
from logging_config import setup_logger, log_and_show_error
import yt_dlp

# ------------------------------
# TODO: 
# 1. 年齡限制的影片，可能需要額外處理 cookies 或登入資訊 (已完成)
# 2. async await 下載功能
# ------------------------------

# 初始化 Logger
logger = setup_logger(__name__)

def timeit(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        result = func(*args, **kwargs)
        end_time = time.perf_counter()
        elapsed = end_time - start_time
        logger.info(f"{func.__name__} executed in {elapsed:.4f} seconds")
        return result
    return wrapper

def _sanitize_filename(filename):
    """
    將檔案名稱中 Windows 不允許的字元替換為底線，
    並移除控制字元或非可見字元。
    """
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    filename = re.sub(r'[\x00-\x1f\x80-\x9f]', '', filename)
    return filename

def _generate_new_filename(download_path, filename):
    """
    檢查 download_path 中是否已存在相同檔名，若存在則在檔名後方加上 (1), (2) 等標記。
    """
    filename = _sanitize_filename(filename)
    base, ext = os.path.splitext(filename)
    new_filename = filename
    counter = 1
    while os.path.exists(os.path.join(download_path, new_filename)):
        new_filename = f"{base} ({counter}){ext}"
        counter += 1
    return new_filename

def _resolution_sort_key(res):
    try:
        width, height = map(int, res.split('x'))
        return width * height  # 根據解析度大小排序
    except ValueError:
        return 0  # 若解析度格式異常，則視為最小

@timeit
def get_video_info(url, file_format="mp4", cookiefile=''):
    """取得影片資訊，包括標題、可用畫質、封面圖 URL、可用字幕"""
    title = "Unknown Title"  # 預設值
    subtitles = ["No subtitle"]  # 預設值
    logger.info("Starting to fetch video info from URL: %s", url)
    ydl_opts = {
        'quiet': True,
        'no_warnings': True,
        'skip_download': True,
        # 讓 yt_dlp 下載字幕資訊
        'writesubtitles': True,
        'writeautomaticsub': True,
        'subtitlesformat': 'srt',
        'noplaylist': True,
    }
    # 若有指定cookies檔案，則加入 cookies 選項
    if cookiefile != '':
        # 使用 cookies 來處理年齡限制或地區限制的影片
        # cookiesfrombrowser無法使用, ERROR: _parse_browser_specification() takes from 1 to 4 positional arguments but 6 were given
        # cookies會過期
        ydl_opts['cookiefile'] = cookiefile # 只有這能用，需先匯出cookies.txt
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(url, download=False)
    title = info.get('title', 'Unknown Title')
    thumbnail_url = info.get('thumbnail')
    if file_format == "mp4":
        resolutions = list(set([
            stream['resolution'] for stream in info.get('formats', []) if stream.get('resolution')
        ]))
        # 排序前先移除 "audio only"
        resolutions = [res for res in resolutions if res.lower() != "audio only"]
        resolutions.sort(key=lambda res: _resolution_sort_key(res), reverse=True)
    else:
        resolutions = ["64kbps","128kbps", "192kbps", "256kbps", "320kbps"]
        resolutions.sort(key=lambda s: int(s.replace("kbps", "")) if s and s.replace("kbps", "").isdigit() else 0, reverse=True)

    # 檢查影片是否有字幕資訊
    if info.get("subtitles"):
        available_subs = list(info["subtitles"].keys())
        subtitles = ["No subtitle"] + available_subs
    elif info.get("automatic_captions"):
        available_subs = list(info["automatic_captions"].keys())
        subtitles = ["No subtitle"] + available_subs

    return title, thumbnail_url, resolutions, subtitles

@timeit
def download_video_audio(url, resolution, download_path, file_format, download_subtitles, subtitle_lang, cookiefile='', progress_callback=None):
    logger.info("Starting to download video/audio from URL: %s", url)
    final_filepath = None
    if file_format == 'mp4':
        # 從解析度字串中取得寬高，例如 "1920x1080"
        try:
            width_str, height_str = resolution.split('x')
            width = int(width_str)
            height = int(height_str)
            # 當高度為 720，且寬度接近 1280，則修正為 1280
            if height == 720 and abs(width - 1280) <= 20:
                width = 1280
        except Exception as e:
            log_and_show_error("解析解析度失敗，請檢查格式是否正確(例如 '1920x1080')", master=None)
            raise ValueError("解析解析度失敗，請檢查格式是否正確(例如 '1920x1080')") from e
        # 將下載檔案暫存為固定名稱，例如 temp_download.mp4
        temp_template = os.path.join(download_path, "temp_download.%(ext)s")
        ydl_opts = {
            # 第一段先嘗試 webm，再嘗試 mp4，最後 fallback 到 best
            'format': f'bestvideo[height={height}]+bestaudio/best/bestvideo+bestaudio/best',
            'outtmpl': temp_template,
            'noplaylist': True,
            'merge_output_format': 'mp4',
            'postprocessor_args': ['-c:a', 'aac'],  # 強制使用 aac 音訊編碼
        }
    elif file_format == 'mp3':
        temp_template = os.path.join(download_path, "temp_download.%(ext)s")
        # 嘗試解析用戶選擇的位元率
        selected_bitrate = None
        try:
            selected_bitrate = int(resolution.replace("kbps", "").strip())
        except Exception:
            pass

        # 若未選擇特定位元率，則預設為 192 kbps
        preferred_quality = str(selected_bitrate) if selected_bitrate is not None else "192"
        if selected_bitrate is not None:
            format_str = f"bestaudio[abr={selected_bitrate}]/bestaudio/best"
            preferred_quality = str(selected_bitrate)
        else:
            format_str = "bestaudio/best"
            preferred_quality = str(selected_bitrate)
        ydl_opts = {
            'format': format_str,
            'outtmpl': temp_template,
            'noplaylist': True,
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': preferred_quality,
            }],
        }
    # 若勾選下載字幕且選擇了特定語言，加入 yt_dlp 下載字幕的選項
    if download_subtitles and subtitle_lang != "No subtitle":
        ydl_opts["subtitlesformat"] = 'srt'
        ydl_opts["writesubtitles"] = True
        ydl_opts["writeautomaticsub"] = True
        ydl_opts["subtitleslangs"] = [subtitle_lang]

    # 若有指定cookies檔案，則加入 cookies 選項
    if cookiefile != '':
        # 使用 cookies 來處理年齡限制或地區限制的影片
        # cookiesfrombrowser無法使用, ERROR: _parse_browser_specification() takes from 1 to 4 positional arguments but 6 were given
        # cookies會過期
        ydl_opts['cookiefile'] = cookiefile # 只有這能用，需先匯出cookies.txt
        
    try:
        def progress_hook(d):
            if d['status'] == 'downloading':
                current = d.get('downloaded_bytes', 0)
                total = d.get('total_bytes_estimate') or d.get('total_bytes') or 1
                fraction = current / total if total else 0
                if progress_callback:
                    progress_callback(fraction * 0.6)
            elif d['status'] == 'finished':
                if progress_callback:
                    progress_callback(0.99)
        ydl_opts['progress_hooks'] = [progress_hook]
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
        if file_format == 'mp4':
            output_ext = 'mp4'
        else:
            output_ext = 'mp3'
        # 取得 yt_dlp 回傳的影片標題
        raw_title = info['title']
        # 利用自訂函式先清理標題，再產生唯一檔案名稱
        safe_title = _sanitize_filename(raw_title)
        filename = safe_title + f".{output_ext}"
        unique_filename = _generate_new_filename(download_path, filename)
        final_filepath = os.path.join(download_path, unique_filename)
        # 取得暫存檔案的完整路徑
        temp_filepath = os.path.join(download_path, f"temp_download.{output_ext}")
        # 重新命名暫存檔案
        os.rename(temp_filepath, final_filepath)
    except Exception as e:
        log_and_show_error(f"下載失敗: {e}", master=None)
        raise e
    finally:
        # 處理完成
        if progress_callback: progress_callback(-1)
    return final_filepath