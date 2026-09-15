"""配置加载模块：全局单例

此前 config.json 在 server.py / cad_controller.py / nlp_processor.py 三处被重复读取，
现统一收敛到本模块，整个进程只加载一次配置。
"""

import json
import logging
import os
from functools import lru_cache

logger = logging.getLogger('cad_mcp_config')

# 配置文件与本模块位于同一目录
_CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'config.json')

# 默认配置（配置文件缺失或损坏时的兜底，与旧版 __init__.py 行为一致）
_DEFAULT_CONFIG = {
    "server": {
        "name": "CAD MCP Server",
        "version": "2.0.0",
        "transport": "stdio",
        "host": "127.0.0.1",
        "port": 8000
    },
    "cad": {
        "type": "AUTOCAD",
        "startup_wait_time": 20,
        "command_delay": 0.5
    },
    "output": {
        "directory": "./output",
        "default_filename": "cad_drawing.dwg"
    }
}


@lru_cache(maxsize=1)
def get_config() -> dict:
    """加载并缓存配置文件（进程内单例）"""
    try:
        with open(_CONFIG_PATH, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info("配置文件加载成功")
        return config
    except Exception as e:
        logger.error(f"加载配置文件失败: {str(e)}，使用默认配置")
        return _DEFAULT_CONFIG
