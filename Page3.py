import os
import re
import time
import functools
import subprocess
from logging_config import setup_logger, log_and_show_error

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

def get_media_duration(file_path):
        try:
            # 使用 ffprobe 取得影片長度（以秒為單位）
            ffprobe_path = os.path.join(os.path.dirname(__file__), 'ffmpeg', 'bin', 'ffprobe.exe')
            result = subprocess.run(
                [ffprobe_path, "-v", "error", "-select_streams", "v:0", "-show_entries", "format=duration", "-of", "default=noprint_wrappers=1:nokey=1", file_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True
            )
            duration_str = result.stdout.strip()
            duration_float = float(duration_str)
            # 轉換成 HH:MM:SS 格式
            hours = int(duration_float // 3600)
            minutes = int((duration_float % 3600) // 60)
            seconds = int(duration_float % 60)
            return f"{hours:02}:{minutes:02}:{seconds:02}"
        except Exception as e:
            log_and_show_error(f"Failed to get media duration: {e}")
            return ""
        
def time_to_seconds(time_str):
    """
    將 "HH:MM:SS" 格式的字串轉換成秒數（可接受秒的小數形式）
    """
    try:
        parts = time_str.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
        elif len(parts) == 2:
            return int(parts[0]) * 60 + float(parts[1])
        else:
            return float(time_str)
    except Exception as e:
        return 0.0

def _get_unique_filename(path):
    """
    若檔案存在，則自動加上 (1)、(2) … 的後綴
    """
    base, ext = os.path.splitext(path)
    counter = 1
    new_path = path
    while os.path.exists(new_path):
        new_path = f"{base} ({counter}){ext}"
        counter += 1
    return new_path

@timeit
def convert_video(input_path, resolution, target_format, start_time, duration, video_transcoder="Default", audio_transcoder="Default", progress_callback=None):
    """
    input_path: 輸入檔案路徑
    resolution: 若為 "Original resolution" 則不進行縮放
    target_format: 輸出格式，例如 mp4、webm 等
    start_time: 剪輯起始時間 (格式 "HH:MM:SS")
    duration: 剪輯持續時間，單位秒（已由 main.py 計算好）
    video_transcoder / audio_transcoder: 若非 "Default" 則加入對應 ffmpeg 參數
    progress_callback: 回呼函式，傳入 0~1 之間的進度值
    """
    base_output = os.path.splitext(input_path)[0] + f"_converted.{target_format}"
    output_path = _get_unique_filename(base_output)

    ffmpeg_path = os.path.join(os.path.dirname(__file__), 'ffmpeg', 'bin', 'ffmpeg.exe')
    command = [ffmpeg_path]
    if start_time and start_time != "00:00:00":
        command.extend(["-ss", start_time])
    command.extend(["-i", input_path])
    if duration > 0:
        command.extend(["-t", str(duration)])
    if video_transcoder != "Default":
        command.extend(["-c:v", video_transcoder])
    if audio_transcoder != "Default":
        command.extend(["-c:a", audio_transcoder])
    if resolution.lower() != "original resolution":
        command.extend(["-vf", f"scale={resolution}"])
    # 加入 -progress 選項，將進度資訊輸出到 stdout
    command.extend(["-progress", "pipe:1"])
    command.append(output_path)

    # 記錄 wall-clock 起始時間
    start_clock = time.time()
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True, encoding="utf-8")
    
    conversion_speed = None  # 預設：已處理媒體時間 / wall-clock 時間
    estimated_total_wall = None

    while True:
        line = process.stdout.readline()
        if not line:
            if process.poll() is not None:
                break
            continue
        line = line.strip()
        # 如果讀到已處理時間，單位是毫秒
        if line.startswith("out_time_ms="):
            try:
                out_time_ms = int(line.split("=")[1])
                processed_media_time = out_time_ms / 1e6  # 換算成秒
                elapsed_wall = time.time() - start_clock  # 已經過的 wall-clock 時間

                # 如果已處理時間大於0，估算編碼速度（媒體時間 / wall-clock）
                if processed_media_time > 0:
                    conversion_speed = processed_media_time / elapsed_wall
                    # 估算總共需要多少 wall-clock 時間
                    estimated_total_wall = elapsed_wall + (duration - processed_media_time) / conversion_speed

                # 若估算出總 wall-clock 時間，則使用 wall-clock 進度
                if estimated_total_wall and progress_callback:
                    progress = elapsed_wall / estimated_total_wall
                    progress_callback(min(progress, 1.0))
            except Exception:
                pass
        elif line.startswith("progress="):
            # 當 ffmpeg 輸出 progress=end 時，代表轉換完成
            if line.split("=")[1] == "end":
                if progress_callback:
                    progress_callback(1.0)
                break
    process.wait()
    return output_path


@timeit
def convert_audio(input_path, bitrate, target_format, start_time, duration, progress_callback=None):
    """
    input_path: 輸入檔案路徑
    bitrate: 使用者指定的位元率（例如 "128kbps"）
    target_format: 輸出格式，例如 mp3、wav、flac、ogg
    start_time: 剪輯起始時間（格式 "HH:MM:SS"）
    duration: 剪輯持續時間（以秒計），可由 main.py 計算得出
    progress_callback: 回呼函式，傳入 0~1 之間的進度數值
    """
    # 產生初步 output 路徑，再檢查是否衝突
    base_output = os.path.splitext(input_path)[0] + f"_converted.{target_format}"
    output_path = _get_unique_filename(base_output)
    
    ffmpeg_path = os.path.join(os.path.dirname(__file__), 'ffmpeg', 'bin', 'ffmpeg.exe')
    command = [ffmpeg_path]
    
    if start_time and start_time != "00:00:00":
        command.extend(["-ss", start_time])
    
    command.extend(["-i", input_path])

    # 加入 -vn 參數，關閉視頻流
    command.extend(["-vn"])
    
    if duration > 0:
        command.extend(["-t", str(duration)])
    
    # 根據目標格式處理 bitrate 參數
    if target_format.lower() == "wav":
        try:
            khz_value = float(bitrate.lower().replace("khz", "").strip())
            sample_rate = int(khz_value * 1000)
            command.extend(["-ar", str(sample_rate)])
        except Exception as e:
            pass
    elif target_format.lower() == "flac":
        # flac 使用預設壓縮參數，不設定 bitrate
        pass
    else:
        # 將 "128kbps" 轉成 "128k" 格式
        converted_bitrate = bitrate.lower().replace("kbps", "k")
        command.extend(["-b:a", converted_bitrate])
    
    # 加入 -progress 選項，讓 ffmpeg 將進度資訊輸出到 stdout
    command.extend(["-progress", "pipe:1"])
    command.append(output_path)
    
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True, encoding="utf-8")
    
    # 解析 ffmpeg 進度資訊，更新 progress_callback
    while True:
        line = process.stdout.readline()
        if not line:
            if process.poll() is not None:
                break
            continue
        line = line.strip()
        processed_media_time = None
        if line.startswith("out_time_ms="):
            try:
                # 注意： out_time_ms 的單位是 microseconds
                out_time_ms = int(line.split("=")[1])
                processed_media_time = out_time_ms / 1000000.0
            except Exception:
                pass
        elif line.startswith("out_time="):
            try:
                out_time_str = line.split("=")[1].strip()
                processed_media_time = time_to_seconds(out_time_str)
            except Exception:
                pass
        
        if processed_media_time is not None and duration > 0 and progress_callback:
            progress = processed_media_time / duration
            progress_callback(min(progress, 1.0))
        
        if line.startswith("progress=") and line.split("=")[1] == "end":
            if progress_callback:
                progress_callback(1.0)
            break
    process.wait()
    return output_path