# logging_config.py
import logging
from tkinter import messagebox
import inspect

def setup_logger(name: str, log_file: str = "app.log", level: int = logging.DEBUG) -> logging.Logger:
    """
    初始化 logger。
    如果 log_file 不存在，FileHandler 會自動建立新檔案。
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    # 如果你需要輸出到終端，也可以啟用 StreamHandler
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    return logger

# 全域 logger 供 logging_config.py 內部使用（若需要）
logger = setup_logger(__name__)

def log_and_show_error(message: str, master=None):
    """
    記錄錯誤訊息並顯示錯誤視窗。
    此函式會動態取得呼叫者的模組名稱，
    以便記錄正確的模組名稱，而非固定的 logging_config。
    """
    # 取得呼叫此函式的 caller frame
    frame = inspect.currentframe()
    if frame is not None and frame.f_back is not None:
        caller_frame = frame.f_back
        module = inspect.getmodule(caller_frame)
        caller_name = module.__name__ if module else __name__
    else:
        caller_name = __name__
    
    # 根據呼叫者的模組名稱取得 logger
    caller_logger = logging.getLogger(caller_name)
    # 使用 stacklevel=2 可讓 log 記錄正確的呼叫資訊（Python 3.8 以上支援）
    caller_logger.error(message, stacklevel=2)
    
    if master:
        master.after(0, lambda: messagebox.showerror("Error", message))
    else:
        messagebox.showerror("Error", message)