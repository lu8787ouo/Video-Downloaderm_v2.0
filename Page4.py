'''
Because ssml is not supported in the edge-tts, so the paragraphs cannot be separated by <break time="Xs"/>.
'''
from edge_tts import Communicate, list_voices
import functools
import time
from logging_config import setup_logger, log_and_show_error
import os

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

async def fetch_voice_names(proxy: str = None):
    # 回傳所有 ShortName 的字串清單
    voices = await list_voices(connector=None, proxy=proxy)
    return [v["ShortName"] for v in voices]

async def convert_text_to_speech(
    text: str,
    voice: str,
    format: str,
    download_path: str,
    speed: str = "0%",
    volume: str = "0%",
    pitch: str = "0%",
    proxy: str = None
):
    """
    將文字轉換為語音，並將音訊存檔於 download_path，回傳檔案路徑。
    """
    try:
        communicate = Communicate(text, voice=voice, rate=speed, volume=volume, pitch=pitch, connector=None, proxy=proxy)
        filename = f"tts_{int(time.time())}.{format}"
        output_path = os.path.join(download_path, filename)
        await communicate.save(output_path)
        logger.info(f"Audio saved to {output_path}")
        return output_path
    except Exception as e:
        logger.error(f"Error converting text to speech: {e}")
        return None