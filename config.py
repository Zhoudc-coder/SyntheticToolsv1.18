import os
import sys
import json
from pathlib import Path

CONFIG_FILENAME = "config.json"
DEFAULT_RETENTION_DAYS = 15

def _get_default_log_dir() -> Path:
    """返回默认日志目录（统一为程序所在目录下的 merge_log）"""
    env_dir = os.environ.get('MERGE_LOG_DIR')
    if env_dir:
        return Path(env_dir)

    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
    return base_dir / 'merge_log'

def _find_config_file() -> Path | None:
    """查找配置文件，返回存在的配置文件路径，否则返回 None"""
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
    else:
        exe_dir = Path(__file__).parent
    config_path = exe_dir / CONFIG_FILENAME
    if config_path.exists():
        return config_path

    home_config = Path.home() / CONFIG_FILENAME
    if home_config.exists():
        return home_config

    return None

def _normalize_unc_path(path_str: str) -> str:
    """
    规范化 Windows UNC 路径。
    用户可能在 config.json 中写出不规范的 UNC 路径，例如：
        \192.168.12.123\a
    或  \\192.168.12.123\a
    该函数确保最终字符串以两个反斜杠开头（即正确的 UNC 格式）。
    """
    path_str = path_str.strip()
    # 如果已经是 UNC 路径（以两个反斜杠开头），直接返回
    if path_str.startswith('\\\\'):
        return path_str
    # 如果以单个反斜杠开头且不是 UNC，补充一个反斜杠使其成为 UNC
    if path_str.startswith('\\') and not path_str.startswith('\\\\'):
        return '\\' + path_str
    # 其他情况（如盘符路径、相对路径）保持不变
    return path_str

def _load_config() -> dict:
    """从配置文件读取全部配置项，若读取失败返回空字典"""
    config_path = _find_config_file()
    if not config_path:
        return {}
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def _load_config_log_dir() -> Path | None:
    """从配置文件读取 log_dir，如果有效返回 Path，否则返回 None"""
    config = _load_config()
    log_dir_str = config.get('log_dir')
    if log_dir_str:
        # 对字符串进行 UNC 路径规范化
        normalized = _normalize_unc_path(log_dir_str)
        return Path(normalized)
    return None

def _load_config_retention_days() -> int:
    """从配置文件读取 retention_days，若不存在或无效则返回默认值15"""
    config = _load_config()
    retention = config.get('retention_days', DEFAULT_RETENTION_DAYS)
    try:
        days = int(retention)
        if days > 0:
            return days
    except (ValueError, TypeError):
        pass
    return DEFAULT_RETENTION_DAYS

def _ensure_config_file(default_log_dir: Path):
    """
    如果配置文件不存在，则创建一个包含默认路径和默认保留天数的配置文件；
    若存在但缺少 retention_days 字段，则补充该字段并保存。
    """
    config_path = _find_config_file()
    if config_path is None:
        # 配置文件不存在，创建新的
        if getattr(sys, 'frozen', False):
            config_dir = Path(sys.executable).parent
        else:
            config_dir = Path(__file__).parent
        config_path = config_dir / CONFIG_FILENAME
        config_data = {
            "log_dir": str(default_log_dir),
            "retention_days": DEFAULT_RETENTION_DAYS
        }
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=4)
        except Exception:
            # 尝试写入用户主目录
            home_config = Path.home() / CONFIG_FILENAME
            try:
                with open(home_config, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=4)
            except Exception:
                pass
    else:
        # 配置文件已存在，检查是否缺少 retention_days，若缺少则添加
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            changed = False
            if 'retention_days' not in config_data:
                config_data['retention_days'] = DEFAULT_RETENTION_DAYS
                changed = True
            if changed:
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=4)
        except Exception:
            pass

def _get_log_dir() -> Path:
    """获取日志目录，优先级：环境变量 > 配置文件 > 默认路径"""
    env_dir = os.environ.get('MERGE_LOG_DIR')
    if env_dir:
        return Path(env_dir)

    config_dir = _load_config_log_dir()
    if config_dir:
        return config_dir

    return _get_default_log_dir()

LOG_DIR = _get_log_dir()
DB_PATH = LOG_DIR / 'merge_log.db'
RETENTION_DAYS = _load_config_retention_days()
XLSX_SUFFIX = '.xlsx'
CSV_SUFFIX = '.csv'

def ensure_log_dir() -> Path:
    """
    确保日志目录存在，若不存在则尝试创建。
    如果创建失败，抛出 RuntimeError 并附带详细说明。
    同时尝试创建或更新配置文件，方便用户修改路径和保留天数。
    """
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        raise RuntimeError(
            f"无法创建日志目录 {LOG_DIR}，请检查：\n"
            f"1. 路径是否正确（网络共享盘是否已连接？）\n"
            f"2. 是否有写入权限\n"
            f"3. 可修改配置文件 config.json 中的 log_dir 项，或设置环境变量 MERGE_LOG_DIR\n"
            f"原始错误：{e}"
        )

    _ensure_config_file(LOG_DIR)

    return LOG_DIR
