"""
CAD MCP 服务包
"""

import os
import json
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger('cad_mcp')

# 加载配置
def load_config():
    """加载配置文件"""
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info("配置文件加载成功")
        return config
    except Exception as e:
        logger.error(f"加载配置文件失败: {str(e)}")
        # 返回默认配置
        return {
            "server": {
                "name": "CAD MCP Server",
                "version": "1.0.0",
                "host": "0.0.0.0",
                "port": 5000,
                "debug": True
            },
            "cad": {
                "type": "AUTOCAD",
                "startup_wait_time": 20,
                "command_delay": 0.5
            },
            "output": {
                "directory": "output",
                "default_filename": "cad_drawing.dwg"
            }
        }

# 导出配置
config = load_config()

__all__ = [
    'config'
]
