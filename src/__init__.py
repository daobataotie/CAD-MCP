"""
CAD MCP 服务包
"""

import logging

# 注意：此处绝不能调用 logging.basicConfig——
# 包方式运行（cad-mcp入口导入src.server）时本模块先于server.py执行，
# basicConfig会抢先给root logger装上默认StreamHandler，
# 导致server.py中的日志配置（含FileHandler）被静默忽略
# （basicConfig在root已有handler时不生效），cad_mcp.log将永远为空。
# 日志配置统一由 server.py 入口完成。

logger = logging.getLogger('cad_mcp')

# 统一使用 config.py 的配置单例（含加载失败时的默认配置兜底）
try:
    from .config import get_config
    config = get_config()
except ImportError:
    from config import get_config
    config = get_config()

__all__ = [
    'config'
]
