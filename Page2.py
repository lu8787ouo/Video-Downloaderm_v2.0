import os
import re
import time
import functools
import subprocess
from logging_config import setup_logger, log_and_show_error
import yt_dlp
import uuid

# ------------------------------
# 初始化 Logger
# ------------------------------
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

@timeit
def parse_playlist(url, resolution, file_format="mp4", cookiefile=''):
    """
    解析播放清單 URL，若不是播放清單則印出錯誤並回傳空列表；
    否則回傳列表，每筆為影片資料字典，包含 "title", "resolution", "format", "url"。
    """
    if "list=" not in url:
        return []
    
    playlist = []
    
    try:
        logger.info("Parsing playlist from URL: %s", url)
        ydl_opts = {
            'quiet': True,
            'extract_flat': False,  # 取得完整資訊(必須為 False)
            'skip_download': True,
            'noplaylist': False,    # 強制解析播放清單
        }
        # 若有指定cookies檔案，則加入 cookies 選項
        if cookiefile != '':
        # 使用 cookies 來處理年齡限制或地區限制的影片
        # cookiesfrombrowser無法使用, ERROR: _parse_browser_specification() takes from 1 to 4 positional arguments but 6 were given
        # cookies會過期
            ydl_opts['cookiefile'] = cookiefile # 只有這能用，需先匯出cookies.txt

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=False)
        if "entries" not in info:
            log_and_show_error("No playlist entries found!")
            return []
        for entry in info['entries']:
            video_url = f"https://www.youtube.com/watch?v={entry['id']}"
            playlist.append({
                "title": entry.get("title", "Unknown"),
                "resolution": resolution,
                "format": file_format,
                "url": video_url
            })
        return playlist
    except Exception as e:
        log_and_show_error(f"Error parsing playlist: {e}")
        return []

@timeit
def download_video_audio_playlist_with_retry(url, resolution, download_path, file_format, cookiefile='', max_retries=3):
    for attempt in range(max_retries):
        logger.info(f"Attempt {attempt + 1} to download: {url}")
        result = download_video_audio_playlist(url, resolution, download_path, file_format, cookiefile)
        if result is not None and os.path.exists(result) and os.path.getsize(result) > 0:
            return result
        logger.info("Retrying in 2 seconds...")
        time.sleep(2)  # 等待2秒再重試
    log_and_show_error(f"多次嘗試仍失敗: {url}")
    return None

def download_video_audio_playlist(url, resolution, download_path, file_format, cookiefile=''):
    temp_id = uuid.uuid4().hex
    final_filepath = None
    if file_format == 'mp4':
        try:
            # 解析如 "1080p", "720p" 這種格式，只保留數字部分作為 height
            height_str = resolution.lower().replace('p', '').strip()
            height = int(height_str)
        except Exception as e:
            log_and_show_error("解析解析度失敗，請檢查格式是否正確(例如 '1080p')")
            raise ValueError("解析解析度失敗，請檢查格式是否正確(例如 '1080p')") from e
        
        temp_template = os.path.join(download_path, f"temp_download_{temp_id}.%(ext)s")
        ydl_opts = {
            # 下載時只指定 height
            'format': f'bestvideo[height={height}]+bestaudio/best/bestvideo+bestaudio/best',
            'outtmpl': temp_template,
            'noplaylist': True,
            'merge_output_format': 'mp4',
            'postprocessor_args': ['-c:a', 'aac'],  # 強制使用 aac 音訊編碼
        }
    elif file_format == 'mp3':
        temp_template = os.path.join(download_path, f"temp_download_{temp_id}.%(ext)s")
        # 嘗試解析用戶選擇的位元率
        selected_bitrate = None
        try:
            selected_bitrate = int(resolution.replace("kbps", "").strip())
        except Exception:
            pass
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

    try:
        # 若有指定cookies檔案，則加入 cookies 選項
        if cookiefile != '':
        # 使用 cookies 來處理年齡限制或地區限制的影片
        # cookiesfrombrowser無法使用, ERROR: _parse_browser_specification() takes from 1 to 4 positional arguments but 6 were given
        # cookies會過期
            ydl_opts['cookiefile'] = cookiefile # 只有這能用，需先匯出cookies.txt

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
        temp_filepath = os.path.join(download_path, f"temp_download_{temp_id}.{output_ext}")
        # 重新命名暫存檔案
        os.rename(temp_filepath, final_filepath)
        
        return final_filepath
    except Exception as e:
        logger.error(f"Error downloading {url}: {e}")
        # log_and_show_error(f"Error downloading {url} : {e}") # 不需要顯示視窗
        return None
