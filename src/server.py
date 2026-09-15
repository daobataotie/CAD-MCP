import os
import sys
import json
import base64
import logging
import argparse
from typing import Annotated, Any, Dict, List, Optional, Union

from pydantic import Field

sys.dont_write_bytecode = True

# 兼容两种运行方式：
# 1. 作为包导入（pip install 后，模块名为 src.server）
# 2. 直接运行 src/server.py（保持旧的 Claude/Cursor 配置可用）
if __package__:
    from .cad_controller import CADController
    from .config import get_config
    from .models import (ToolResult, ListEntitiesResult,
                         ListLayersResult, EntityPropertiesResult, ListBlocksResult)
    from .nlp_processor import NLPProcessor
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from cad_controller import CADController
    from config import get_config
    from models import (ToolResult, ListEntitiesResult,
                        ListLayersResult, EntityPropertiesResult, ListBlocksResult)
    from nlp_processor import NLPProcessor

from mcp.server.fastmcp import FastMCP, Context
from mcp.types import ToolAnnotations, ImageContent, TextContent
from logging.handlers import RotatingFileHandler

# 配置Windows环境下的UTF-8编码
if sys.platform == "win32" and os.environ.get('PYTHONIOENCODING') is None:
    sys.stdin.reconfigure(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

config = get_config()

def _build_log_handlers() -> list:
    """构建日志handler：控制台 + 滚动文件

    日志文件优先写源码目录（开发期就近查看、路径不随启动目录漂移）；
    pip安装到site-packages等无写权限的位置时回退到用户目录，
    避免FileHandler构造即抛PermissionError导致服务器启动失败。
    使用滚动文件（单文件5MB、保留2个备份）防止长期运行日志无限增长。
    StreamHandler默认输出到stderr，stdio传输模式下不会污染协议数据。
    """
    handlers = [logging.StreamHandler()]
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cad_mcp.log'),
        os.path.join(os.path.expanduser('~'), '.cad-mcp', 'cad_mcp.log'),
    ]
    for path in candidates:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            handlers.append(RotatingFileHandler(path, maxBytes=5 * 1024 * 1024,
                                                backupCount=2, encoding='utf-8'))
            break
        except OSError:
            continue
    return handlers


# 配置日志（必须在模块导入早期完成，且全进程只能有这一处basicConfig：
# root已有handler时basicConfig会被静默忽略，详见src/__init__.py的说明）
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=_build_log_handlers()
)

logger = logging.getLogger('mcp_cad_server')
logger.info("启动 CAD MCP 服务器")

SERVER_NAME = config["server"]["name"]
SERVER_VERSION = config["server"].get("version", "2.0.0")

PROMPT_TEMPLATE = """
助手的目标是演示如何使用CAD MCP服务。通过这个服务，您可以直接在对话中控制CAD软件，创建和修改图形。

<cad-mcp>

提示：
这个服务器提供了一个名为"cad-assistant"的预设提示，帮助用户通过自然语言指令控制CAD。用户可以要求绘制各种形状、修改图形元素或保存图纸。

资源：
此服务器提供"drawing://current"资源，表示当前的CAD图纸状态（会话记录 + 来自CAD文档的实时统计）。

工具：

此服务器提供以下几类工具：

基本绘图工具：
"draw_line": 在CAD中绘制直线
"draw_circle": 在CAD中绘制圆
"draw_arc": 在CAD中绘制弧
"draw_ellipse": 在CAD中绘制椭圆
"draw_polyline": 在CAD中绘制多段线
"draw_rectangle": 在CAD中绘制矩形
"draw_text": 在CAD中添加文本
"draw_hatch": 在CAD中绘制填充
"add_dimension": 在CAD中添加线性标注

查询工具（了解图纸实际内容）：
"list_layers": 列出所有图层
"list_entities": 列出模型空间中的实体（返回可编辑的句柄）
"get_entity_properties": 按句柄查询实体详细属性
"screenshot": 截取CAD窗口画面确认绘图效果

编辑工具（通过句柄操作已有实体）：
"erase_entity": 删除实体
"move_entity": 移动实体
"rotate_entity": 旋转实体
"scale_entity": 缩放实体
"copy_entity": 复制实体
"mirror_entity": 镜像实体
"offset_entity": 偏移实体
"array_linear_entity" / "array_polar_entity": 阵列
"undo" / "redo": 撤销/重做

图层与图纸管理：
"create_layer" / "set_current_layer": 图层管理
"new_drawing" / "open_drawing" / "close_drawing" / "save_drawing": 图纸管理
"create_block" / "insert_block" / "list_blocks": 图块操作
"send_command": 直通CAD命令行执行任意命令

遗留接口：
"process_command": 处理自然语言命令并转换为CAD操作（推荐优先使用上述结构化工具）

</cad-mcp>

请以友好的方式开始演示，例如："嗨！今天我将向您展示如何使用CAD MCP服务。通过这个服务，您可以直接在我们的对话中控制CAD软件，无需手动操作界面。让我们开始吧！"
"""

# 服务器级使用说明（随 initialize 响应下发给客户端）
SERVER_INSTRUCTIONS = """CAD MCP 服务器：通过 MCP 工具控制 Windows 上的 AutoCAD / 浩辰CAD / 中望CAD。

使用要点：
- 首次调用工具时若 CAD 未运行会自动启动（可能需要约20秒），期间会有进度提示
- 绘图/编辑工具返回 entity_handles（实体句柄），可用于后续编辑与查询
- 颜色参数支持 ACI 索引(1-256)或中英文颜色名称，不指定则保持随层(ByLayer)
- 绘图后建议用 screenshot 截图确认效果，用 list_entities 查询图纸实际内容
- send_command 可直通 CAD 命令行，执行任何未封装为工具的 CAD 命令
"""

# ==================== 参数类型（保留原版 schema 的坐标约束） ====================

# 坐标点：2~3个数字 [x, y] 或 [x, y, z]
Point = Annotated[List[float], Field(min_length=2, max_length=3, description="坐标 [x, y] 或 [x, y, z]")]
# 点集：至少2个坐标点
Points = Annotated[List[Point], Field(min_length=2, description="点集 [[x1,y1], [x2,y2], ...]（每个点2~3个分量）")]
# 填充边界点集：至少3个坐标点
HatchPoints = Annotated[List[Point], Field(min_length=3, description="填充边界点集（至少3个点）")]
# 正数参数（半径/字高/缩放系数等：CAD对0和负值会直接报错，在schema层提前拦截）
PosFloat = Annotated[float, Field(gt=0, description="正数")]
# 阵列数量（含原实体）：下限2（<2没有阵列意义），上限防止误操作生成海量实体卡死图纸
ArrayCount = Annotated[int, Field(ge=2, le=10000, description="阵列总数（含原实体，2~10000）")]


class CADService:
    def __init__(self):
        """初始化CAD服务"""
        self.controller = CADController()
        self.nlp_processor = NLPProcessor()
        self.drawing_state = {
            "entities": [],
            "current_layer": "0",
            "last_command": "",
            "last_result": ""
        }
        logger.info("CAD服务已初始化")

    @staticmethod
    def _handle(entity) -> Optional[str]:
        """安全读取实体句柄"""
        return CADController._safe_attr(entity, "Handle")

    def start_cad(self):
        """启动CAD"""
        return self.controller.start_cad()

    def draw_line(self, start_point, end_point, layer=None, color=None, lineweight=None):
        """绘制直线"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.draw_line(start_point, end_point, current_layer, color, lineweight)
        if result:
            self.drawing_state["entities"].append({
                "type": "line",
                "handle": self._handle(result),
                "start": start_point,
                "end": end_point,
                "layer": current_layer,
                "color": color,
                "lineweight": lineweight
            })
            self.drawing_state["last_command"] = f"绘制直线从{start_point}到{end_point}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def draw_circle(self, center, radius, layer=None, color=None, lineweight=None):
        """绘制圆"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.draw_circle(center, radius, current_layer, color, lineweight)
        if result:
            self.drawing_state["entities"].append({
                "type": "circle",
                "handle": self._handle(result),
                "center": center,
                "radius": radius,
                "layer": current_layer,
                "color": color,
                "lineweight": lineweight
            })
            self.drawing_state["last_command"] = f"绘制圆，中心点{center}，半径{radius}，图层{current_layer}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def draw_arc(self, center, radius, start_angle, end_angle, layer=None, color=None, lineweight=None):
        """绘制弧"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.draw_arc(center, radius, start_angle, end_angle, current_layer, color, lineweight)

        if result:
            self.drawing_state["entities"].append({
                "type": "arc",
                "handle": self._handle(result),
                "center": center,
                "radius": radius,
                "start_angle": start_angle,
                "end_angle": end_angle,
                "layer": current_layer,
                "color": color,
                "lineweight": lineweight
            })
            self.drawing_state["last_command"] = f"绘制弧，中心点{center}，半径{radius}，起始角度{start_angle}，结束角度{end_angle}，图层{current_layer}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def draw_ellipse(self, center, major_axis, minor_axis, rotation=0, layer=None, color=None, lineweight=None):
        """绘制椭圆"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.draw_ellipse(center, major_axis, minor_axis, rotation, current_layer, color, lineweight)

        if result:
            self.drawing_state["entities"].append({
                "type": "ellipse",
                "handle": self._handle(result),
                "center": center,
                "major_axis": major_axis,
                "minor_axis": minor_axis,
                "rotation": rotation,
                "layer": current_layer,
                "color": color,
                "lineweight": lineweight
            })
            self.drawing_state["last_command"] = f"绘制椭圆，中心点{center}，长轴{major_axis}，短轴{minor_axis}，旋转角度{rotation}，图层{current_layer}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def draw_polyline(self, points, closed=False, layer=None, color=None, lineweight=None):
        """绘制多段线"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.draw_polyline(points, closed, current_layer, color, lineweight)
        if result:
            self.drawing_state["entities"].append({
                "type": "polyline",
                "handle": self._handle(result),
                "points": points,
                "closed": closed,
                "layer": current_layer,
                "color": color,
                "lineweight": lineweight
            })
            self.drawing_state["last_command"] = f"绘制多段线，点集{points}，{'闭合' if closed else '不闭合'}，图层{current_layer}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def draw_rectangle(self, corner1, corner2, layer=None, color=None, lineweight=None):
        """绘制矩形"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.draw_rectangle(corner1, corner2, current_layer, color, lineweight)
        if result:
            self.drawing_state["entities"].append({
                "type": "rectangle",
                "handle": self._handle(result),
                "corner1": corner1,
                "corner2": corner2,
                "layer": current_layer,
                "color": color,
                "lineweight": lineweight
            })
            self.drawing_state["last_command"] = f"绘制矩形，对角点{corner1}和{corner2}，图层{current_layer}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def draw_text(self, position, text, height=2.5, rotation=0, layer=None, color=None):
        """添加文本"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.draw_text(position, text, height, rotation, current_layer, color)
        if result:
            self.drawing_state["entities"].append({
                "type": "text",
                "handle": self._handle(result),
                "position": position,
                "text": text,
                "height": height,
                "rotation": rotation,
                "layer": current_layer,
                "color": color
            })
            self.drawing_state["last_command"] = f"添加文本'{text}'，位置{position}，高度{height}，旋转{rotation}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def draw_hatch(self, points, pattern_name="SOLID", scale=1.0, layer=None, color=None):
        """绘制填充"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.draw_hatch(points, pattern_name, scale, current_layer, color)
        if result:
            self.drawing_state["entities"].append({
                "type": "hatch",
                "handle": self._handle(result),
                "points": points,
                "pattern_name": pattern_name,
                "scale": scale,
                "layer": current_layer,
                "color": color
            })
            self.drawing_state["last_command"] = f"绘制填充，点集{points}，图案{pattern_name}，比例{scale}，图层{current_layer}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def add_dimension(self, start_point, end_point, text_position=None, textheight=5, layer=None, color=None):
        """添加线性标注"""
        if not self.controller.is_running():
            self.start_cad()

        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]

        result = self.controller.add_dimension(start_point, end_point, text_position, textheight, current_layer, color)
        if result:
            self.drawing_state["entities"].append({
                "type": "dimension",
                "handle": self._handle(result),
                "start": start_point,
                "end": end_point,
                "text_position": text_position,
                "textheight": textheight,
                "layer": current_layer,
                "color": color
            })
            self.drawing_state["last_command"] = f"添加标注从{start_point}到{end_point}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def save_drawing(self, file_path):
        """保存图纸"""
        if not self.controller.is_running():
            return False

        result = self.controller.save_drawing(file_path)
        if result:
            self.drawing_state["last_command"] = f"保存图纸到{file_path}"
            self.drawing_state["last_result"] = "成功"
        else:
            self.drawing_state["last_result"] = "失败"

        return result

    def process_command(self, command: str) -> Dict[str, Any]:
        """处理自然语言命令（遗留接口，推荐客户端直接调用结构化绘图工具）"""
        if not self.controller.is_running():
            self.start_cad()
        # 使用NLP处理器解析命令
        parsed_command = self.nlp_processor.process_command(command)
        command_type = parsed_command.get("type")
        try:
            # 未识别的命令明确返回失败（此前会静默返回None）
            if command_type in (None, "unknown", "error"):
                return {
                    "success": False,
                    "message": parsed_command.get("error") or "无法识别的命令",
                    "original_command": command
                }

            # 统一解析颜色：命令文本中提取的颜色优先于解析结果中的颜色参数
            color = parsed_command.get("color")
            color_rgb = self.nlp_processor.extract_color_from_command(command)
            if color_rgb is not None:
                color = color_rgb
            # 获取线宽参数
            lineweight = parsed_command.get("lineweight")

            if command_type == "draw_line":
                result = self.draw_line(parsed_command.get("start_point"), parsed_command.get("end_point"), None, color, lineweight)
                return {
                    "success": result is not None,
                    "message": "直线已绘制" if result else "绘制直线失败",
                    "entity_id": self._handle(result)
                }

            elif command_type == "draw_circle":
                result = self.draw_circle(parsed_command.get("center"), parsed_command.get("radius"), None, color, lineweight)
                return {
                    "success": result is not None,
                    "message": "圆已绘制" if result else "绘制圆失败",
                    "entity_id": self._handle(result)
                }

            elif command_type == "draw_arc":
                center = parsed_command.get("center")
                radius = parsed_command.get("radius")
                start_angle = parsed_command.get("start_angle")
                end_angle = parsed_command.get("end_angle")
                # 确保所有必要参数都存在且有效
                if center is None or radius is None or start_angle is None or end_angle is None:
                    return {
                        "success": False,
                        "message": "绘制圆弧失败：缺少必要参数",
                        "error": "缺少必要参数：中心点、半径、起始角度或结束角度"
                    }
                result = self.draw_arc(center, radius, start_angle, end_angle, None, color, lineweight)
                return {
                    "success": result is not None,
                    "message": "圆弧已绘制" if result else "绘制圆弧失败",
                    "entity_id": self._handle(result)
                }

            elif command_type == "draw_ellipse":
                center = parsed_command.get("center")
                major_axis = parsed_command.get("major_axis")
                minor_axis = parsed_command.get("minor_axis")
                rotation = parsed_command.get("rotation", 0)  # 默认旋转角度为0
                # 确保所有必要参数都存在且有效
                if center is None or major_axis is None or minor_axis is None:
                    return {
                        "success": False,
                        "message": "绘制椭圆失败：缺少必要参数",
                        "error": "缺少必要参数：中心点、长轴或短轴"
                    }
                result = self.draw_ellipse(center, major_axis, minor_axis, rotation, None, color, lineweight)
                return {
                    "success": result is not None,
                    "message": "椭圆已绘制" if result else "绘制椭圆失败",
                    "entity_id": self._handle(result)
                }

            elif command_type == "draw_rectangle":
                result = self.draw_rectangle(parsed_command.get("corner1"), parsed_command.get("corner2"), None, color, lineweight)
                return {
                    "success": result is not None,
                    "message": "矩形已绘制" if result else "绘制矩形失败",
                    "entity_id": self._handle(result)
                }

            elif command_type == "draw_text":
                height = parsed_command.get("height", 2.5)
                rotation = parsed_command.get("rotation", 0)
                result = self.draw_text(parsed_command.get("position"), parsed_command.get("text"), height, rotation, None, color)
                return {
                    "success": result is not None,
                    "message": "文本已添加" if result else "添加文本失败",
                    "entity_id": self._handle(result)
                }

            elif command_type == "draw_hatch":
                points = parsed_command.get("points")
                pattern_name = parsed_command.get("pattern_name", "SOLID")
                scale = parsed_command.get("scale", 1.0)
                # 确保所有必要参数都存在且有效
                if points is None or len(points) < 3:
                    return {
                        "success": False,
                        "message": "绘制填充失败：缺少必要参数或点数不足",
                        "error": "填充边界至少需要3个点"
                    }
                result = self.draw_hatch(points, pattern_name, scale, None, color)
                return {
                    "success": result is not None,
                    "message": "填充已绘制" if result else "绘制填充失败",
                    "entity_id": self._handle(result)
                }

            elif command_type == "add_dimension":
                # 统一经由服务层（同时更新会话状态）
                result = self.add_dimension(parsed_command.get("start_point"), parsed_command.get("end_point"), parsed_command.get("text_position"))
                return {
                    "success": result is not None,
                    "message": "标注已添加" if result else "添加标注失败",
                    "entity_id": self._handle(result)
                }

            elif command_type == "save":
                file_path = parsed_command.get("file_path")
                result = self.save_drawing(file_path)
                return {
                    "success": result,
                    "message": f"图纸已保存到 {file_path}" if result else f"保存图纸到 {file_path} 失败"
                }

            # 处理图层操作
            elif command_type == "create_layer":
                layer_name = parsed_command.get("layer_name")
                if not layer_name:
                    return {
                        "success": False,
                        "message": "创建图层失败：未指定图层名称"
                    }
                result = self.controller.create_layer(layer_name)  # , color
                if result:
                    self.drawing_state["current_layer"] = layer_name
                return {
                    "success": result,
                    "message": f"图层 {layer_name} 已创建" if result else f"创建图层 {layer_name} 失败"
                }

            # 其他已知但暂不支持的命令类型
            return {
                "success": False,
                "message": f"暂不支持的命令类型: {command_type}"
            }

        except Exception as e:
            return {
                "success": False,
                "message": f"处理命令时出错: {str(e)}"
            }


# ==================== 服务实例与通用工具函数 ====================

cad_service = CADService()


async def com(fn, *args, **kwargs):
    """在专用COM线程上执行调用（避免阻塞事件循环，且线程安全）"""
    return await cad_service.controller.executor.run(fn, *args, **kwargs)


def _resolve_color(color) -> Optional[int]:
    """解析颜色参数为ACI索引：None=随层(ByLayer)；支持索引和中英文颜色名称"""
    return cad_service.nlp_processor.color_to_index(color)


async def _entity_result(entity, ok_message: str, fail_message: str) -> ToolResult:
    """将单个COM实体包装为统一的工具结果

    句柄必须在COM线程上读取——实体由COM线程创建，
    主线程直接访问未封送的COM接口指针属于跨套间调用，会失败甚至崩溃。
    """
    if entity is None:
        return ToolResult(success=False, message=fail_message, error="CAD操作失败，详情见 cad_mcp.log")
    handle = await com(CADController._safe_attr, entity, "Handle")
    return ToolResult(success=True, message=ok_message,
                      entity_handles=[str(handle)] if handle else [])


def _extract_handles_sync(entities: List[Any]) -> List[str]:
    """在COM线程上批量提取实体句柄"""
    return [str(h) for h in (CADController._safe_attr(e, "Handle") for e in entities) if h]


async def _entities_result(entities: List[Any], ok_message: str, fail_message: str) -> ToolResult:
    """将COM实体列表包装为统一的工具结果（句柄在COM线程上读取）"""
    if not entities:
        return ToolResult(success=False, message=fail_message)
    handles = await com(_extract_handles_sync, entities)
    return ToolResult(success=True, message=ok_message, entity_handles=handles)


async def _ensure_running(ctx: Context) -> bool:
    """确保CAD已运行：未运行时先给进度提示再自动启动（首次启动可能较慢）

    注意必须真正发起启动——此前只发提示不启动，导致screenshot等
    不带自动启动逻辑的工具在CAD未运行时直接抛异常。
    """
    if cad_service.controller.is_running():
        return True
    await ctx.info("CAD未运行，正在启动CAD（首次启动可能需要约20秒）...")
    if await com(cad_service.start_cad):
        return True
    await ctx.info("CAD自动启动失败，请检查CAD软件是否已安装并能正常启动")
    return False


# ==================== MCP 服务器与工具注册 ====================

mcp = FastMCP(name=SERVER_NAME, instructions=SERVER_INSTRUCTIONS)


# ---------- 基本绘图工具 ----------

@mcp.tool()
async def draw_line(ctx: Context, start_point: Point, end_point: Point,
                    layer: Optional[str] = None, color: Union[int, str, None] = None,
                    lineweight: Optional[float] = None) -> ToolResult:
    """在CAD中绘制直线

    color 支持CAD颜色索引(1-256)或中英文颜色名称（如"红色"、"red"），不指定则保持随层颜色。
    """
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.draw_line, start_point, end_point, layer, _resolve_color(color), lineweight)
    return await _entity_result(entity, "直线已绘制", "绘制直线失败")


@mcp.tool()
async def draw_circle(ctx: Context, center: Point, radius: PosFloat,
                      layer: Optional[str] = None, color: Union[int, str, None] = None,
                      lineweight: Optional[float] = None) -> ToolResult:
    """在CAD中绘制圆

    color 支持CAD颜色索引(1-256)或中英文颜色名称，不指定则保持随层颜色。
    """
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.draw_circle, center, radius, layer, _resolve_color(color), lineweight)
    return await _entity_result(entity, "圆已绘制", "绘制圆失败")


@mcp.tool()
async def draw_arc(ctx: Context, center: Point, radius: PosFloat,
                   start_angle: float, end_angle: float,
                   layer: Optional[str] = None, color: Union[int, str, None] = None,
                   lineweight: Optional[float] = None) -> ToolResult:
    """在CAD中绘制弧（角度单位：度）"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.draw_arc, center, radius, start_angle, end_angle,
                       layer, _resolve_color(color), lineweight)
    return await _entity_result(entity, "圆弧已绘制", "绘制圆弧失败")


@mcp.tool()
async def draw_ellipse(ctx: Context, center: Point, major_axis: PosFloat, minor_axis: PosFloat,
                       rotation: float = 0, layer: Optional[str] = None,
                       color: Union[int, str, None] = None,
                       lineweight: Optional[float] = None) -> ToolResult:
    """在CAD中绘制椭圆（rotation 为旋转角度，单位：度）"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.draw_ellipse, center, major_axis, minor_axis, rotation,
                       layer, _resolve_color(color), lineweight)
    return await _entity_result(entity, "椭圆已绘制", "绘制椭圆失败")


@mcp.tool()
async def draw_polyline(ctx: Context, points: Points, closed: bool = False,
                        layer: Optional[str] = None, color: Union[int, str, None] = None,
                        lineweight: Optional[float] = None) -> ToolResult:
    """在CAD中绘制多段线（closed=True 时闭合）"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.draw_polyline, points, closed, layer, _resolve_color(color), lineweight)
    return await _entity_result(entity, "多段线已绘制", "绘制多段线失败")


@mcp.tool()
async def draw_rectangle(ctx: Context, corner1: Point, corner2: Point,
                         layer: Optional[str] = None, color: Union[int, str, None] = None,
                         lineweight: Optional[float] = None) -> ToolResult:
    """在CAD中绘制矩形（由两个对角点确定）"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.draw_rectangle, corner1, corner2, layer, _resolve_color(color), lineweight)
    return await _entity_result(entity, "矩形已绘制", "绘制矩形失败")


@mcp.tool()
async def draw_text(ctx: Context, position: Point, text: str, height: PosFloat = 2.5,
                    rotation: float = 0, layer: Optional[str] = None,
                    color: Union[int, str, None] = None) -> ToolResult:
    """在CAD中添加单行文本"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.draw_text, position, text, height, rotation, layer, _resolve_color(color))
    return await _entity_result(entity, "文本已添加", "添加文本失败")


@mcp.tool()
async def draw_hatch(ctx: Context, points: HatchPoints, pattern_name: str = "SOLID",
                     scale: PosFloat = 1.0, layer: Optional[str] = None,
                     color: Union[int, str, None] = None) -> ToolResult:
    """在CAD中绘制填充（points 为闭合边界点集，至少3个点）

    pattern_name 为填充图案名称，默认 SOLID（实体填充）。
    """
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.draw_hatch, points, pattern_name, scale, layer, _resolve_color(color))
    return await _entity_result(entity, "填充已绘制", "绘制填充失败")


@mcp.tool()
async def add_dimension(ctx: Context, start_point: Point, end_point: Point,
                        text_position: Optional[Point] = None, textheight: PosFloat = 5,
                        layer: Optional[str] = None, color: Union[int, str, None] = None) -> ToolResult:
    """在CAD中添加线性标注（text_position 不指定时自动放在中点上方）"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    entity = await com(cad_service.add_dimension, start_point, end_point, text_position, textheight,
                       layer, _resolve_color(color))
    return await _entity_result(entity, "标注已添加", "添加标注失败")


# ---------- 查询工具（闭合AI反馈回路：AI能看到图纸实际内容） ----------

@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
async def list_layers(ctx: Context) -> ListLayersResult:
    """列出当前图纸的所有图层（名称、颜色、冻结/锁定/开关状态、是否当前层）"""
    if not cad_service.controller.is_running():
        return ListLayersResult(success=False, message="CAD未运行")
    layers = await com(cad_service.controller.list_layers)
    return ListLayersResult(success=True, message=f"共{len(layers)}个图层", layers=layers)


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
async def list_entities(ctx: Context, entity_type: Optional[str] = None,
                        limit: int = 50) -> ListEntitiesResult:
    """列出模型空间中的实体（实时查询，与CAD文档保持同步）

    返回实体的句柄(handle)，可用于编辑类工具与 get_entity_properties。
    entity_type 可按类型过滤（如 line/circle/arc/text）；limit 控制返回数量。
    """
    if not cad_service.controller.is_running():
        return ListEntitiesResult(success=False, message="CAD未运行")
    result = await com(cad_service.controller.list_entities, entity_type, limit)
    return ListEntitiesResult(success=True,
                              message=f"模型空间共{result['total']}个实体，返回{len(result['entities'])}个",
                              total=result["total"],
                              returned=len(result["entities"]),
                              entities=result["entities"])


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
async def get_entity_properties(ctx: Context, handle: str) -> EntityPropertiesResult:
    """按句柄查询实体详细属性（几何信息随实体类型而异：坐标/半径/文本内容等）"""
    if not cad_service.controller.is_running():
        return EntityPropertiesResult(success=False, message="CAD未运行")
    try:
        props = await com(cad_service.controller.get_entity_properties, handle)
        return EntityPropertiesResult(success=True, message="查询成功", **props)
    except Exception as e:
        return EntityPropertiesResult(success=False, message=f"查询句柄 {handle} 失败", error=str(e))


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
async def screenshot(ctx: Context):
    """截取CAD主窗口画面并返回PNG预览图（用于确认绘图效果）"""
    if not await _ensure_running(ctx):
        return [TextContent(type="text", text="CAD未运行且自动启动失败，无法截图")]
    png_bytes = await com(cad_service.controller.screenshot)
    # 编码为base64返回图片内容块
    b64 = base64.b64encode(png_bytes).decode("utf-8")
    return [
        ImageContent(type="image", data=b64, mimeType="image/png"),
        TextContent(type="text", text="CAD窗口截图（PNG）"),
    ]


# ---------- 编辑工具（通过句柄操作已有实体） ----------

@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
async def erase_entity(ctx: Context, handle: str) -> ToolResult:
    """删除指定句柄的实体"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    ok = await com(cad_service.controller.erase_entity, handle)
    return ToolResult(success=ok, message="实体已删除" if ok else f"删除句柄 {handle} 失败")


@mcp.tool()
async def move_entity(ctx: Context, handle: str, displacement: Point) -> ToolResult:
    """移动实体（displacement 为位移向量 [dx, dy, dz]）"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    ok = await com(cad_service.controller.move_entity, handle, displacement)
    return ToolResult(success=ok, message="实体已移动" if ok else f"移动句柄 {handle} 失败")


@mcp.tool()
async def rotate_entity(ctx: Context, handle: str, base_point: Point, angle_degrees: float) -> ToolResult:
    """绕基点旋转实体（角度单位：度，逆时针为正）"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    ok = await com(cad_service.controller.rotate_entity, handle, base_point, angle_degrees)
    return ToolResult(success=ok, message="实体已旋转" if ok else f"旋转句柄 {handle} 失败")


@mcp.tool()
async def scale_entity(ctx: Context, handle: str, base_point: Point, scale_factor: PosFloat) -> ToolResult:
    """绕基点缩放实体（scale_factor>1 放大，<1 缩小）"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    ok = await com(cad_service.controller.scale_entity, handle, base_point, scale_factor)
    return ToolResult(success=ok, message="实体已缩放" if ok else f"缩放句柄 {handle} 失败")


@mcp.tool()
async def copy_entity(ctx: Context, handle: str, displacement: Point) -> ToolResult:
    """复制实体并平移指定位移，返回新实体句柄"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    copy = await com(cad_service.controller.copy_entity, handle, displacement)
    return await _entity_result(copy, "实体已复制", f"复制句柄 {handle} 失败")


@mcp.tool()
async def mirror_entity(ctx: Context, handle: str, point1: Point, point2: Point) -> ToolResult:
    """沿两点确定的镜像轴镜像实体（保留原实体），返回新实体句柄"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    mirrored = await com(cad_service.controller.mirror_entity, handle, point1, point2)
    return await _entity_result(mirrored, "实体已镜像", f"镜像句柄 {handle} 失败")


@mcp.tool()
async def offset_entity(ctx: Context, handle: str, distance: float) -> ToolResult:
    """偏移实体（正负值决定偏移方向），返回新实体句柄"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    entities = await com(cad_service.controller.offset_entity, handle, distance)
    return await _entities_result(entities, "实体已偏移", f"偏移句柄 {handle} 失败")


@mcp.tool()
async def array_linear_entity(ctx: Context, handle: str, count: ArrayCount, displacement: Point) -> ToolResult:
    """矩形阵列：沿位移向量方向复制 count-1 份（含原实体共 count 个）"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    entities = await com(cad_service.controller.array_linear_entity, handle, count, displacement)
    return await _entities_result(entities, f"已完成{count}个的线性阵列", f"阵列句柄 {handle} 失败")


@mcp.tool()
async def array_polar_entity(ctx: Context, handle: str, count: ArrayCount, center: Point,
                             total_angle_degrees: float = 360.0) -> ToolResult:
    """环形阵列：绕中心点在 total_angle 角度范围内复制 count-1 份"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    entities = await com(cad_service.controller.array_polar_entity, handle, count, center, total_angle_degrees)
    return await _entities_result(entities, f"已完成{count}个的环形阵列", f"阵列句柄 {handle} 失败")


@mcp.tool()
async def undo(ctx: Context) -> ToolResult:
    """撤销上一步操作"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    ok = await com(cad_service.controller.undo)
    return ToolResult(success=ok, message="已撤销" if ok else "撤销失败")


@mcp.tool()
async def redo(ctx: Context) -> ToolResult:
    """重做被撤销的操作"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    ok = await com(cad_service.controller.redo)
    return ToolResult(success=ok, message="已重做" if ok else "重做失败")


# ---------- 图层管理 ----------

@mcp.tool()
async def create_layer(ctx: Context, layer_name: str) -> ToolResult:
    """创建新图层并设为当前图层（若已存在则直接激活）"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    ok = await com(cad_service.controller.create_layer, layer_name)
    if ok:
        cad_service.drawing_state["current_layer"] = layer_name
    return ToolResult(success=ok, message=f"图层 {layer_name} 已创建并设为当前层" if ok else f"创建图层 {layer_name} 失败")


@mcp.tool()
async def set_current_layer(ctx: Context, layer_name: str) -> ToolResult:
    """切换当前图层（图层必须已存在）"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    ok = await com(cad_service.controller.set_current_layer, layer_name)
    if ok:
        cad_service.drawing_state["current_layer"] = layer_name
    return ToolResult(success=ok, message=f"当前图层已切换为 {layer_name}" if ok else f"图层 {layer_name} 不存在或切换失败")


# ---------- 图纸管理 ----------

@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
async def new_drawing(ctx: Context) -> ToolResult:
    """新建图纸文档"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    ok = await com(cad_service.controller.new_drawing)
    if ok:
        # 会话绘图记录归属于旧文档，切换文档后清空
        cad_service.drawing_state["entities"] = []
    return ToolResult(success=ok, message="新文档已创建" if ok else "新建文档失败")


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
async def open_drawing(ctx: Context, file_path: str) -> ToolResult:
    """打开已有图纸文件（DWG/DXF等）"""
    if cad_service.controller.app is None and not await com(cad_service.controller.start_cad):
        return ToolResult(success=False, message="CAD未运行且启动失败")
    try:
        ok = await com(cad_service.controller.open_drawing, file_path)
    except FileNotFoundError as e:
        return ToolResult(success=False, message=str(e))
    if ok:
        cad_service.drawing_state["entities"] = []
    return ToolResult(success=ok, message=f"文档已打开: {file_path}" if ok else f"打开文档失败: {file_path}")


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
async def close_drawing(ctx: Context, save: bool = False) -> ToolResult:
    """关闭当前文档（save=False 时不保存直接关闭）"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    ok = await com(cad_service.controller.close_drawing, save)
    if ok:
        cad_service.drawing_state["entities"] = []
    return ToolResult(success=ok, message="文档已关闭" if ok else "关闭文档失败")


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
async def save_drawing(ctx: Context, file_path: str) -> ToolResult:
    """保存当前图纸（按扩展名自动选择 DWG/DXF 格式；已存在的文件将被覆盖）"""
    if not cad_service.controller.is_running():
        return ToolResult(success=False, message="CAD未运行")
    await ctx.info(f"正在保存图纸到 {file_path} ...")
    try:
        ok = await com(cad_service.save_drawing, file_path)
    except ValueError as e:
        # 目标路径被其他打开的图纸占用：提前快速失败并给出可操作的提示
        return ToolResult(success=False, message=f"保存失败: {e}", error=str(e))
    return ToolResult(success=ok, message=f"图纸已保存到 {file_path}" if ok else f"保存图纸到 {file_path} 失败")


@mcp.tool(annotations=ToolAnnotations(destructiveHint=True))
async def send_command(ctx: Context, command: str) -> ToolResult:
    """直通CAD命令行执行任意命令（功能逃生舱）

    示例："_.ZOOM E"（缩放全图）、"_.ERASE" 配合选择集等。
    建议命令以 "_." 前缀强制英文命令名以兼容中文版CAD。
    """
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    ok = await com(cad_service.controller.send_command, command)
    return ToolResult(success=ok, message=f"命令已提交: {command}" if ok else "命令执行失败")


# ---------- 图块操作 ----------

@mcp.tool()
async def create_block(ctx: Context, block_name: str, base_point: Optional[Point] = None) -> ToolResult:
    """创建图块定义（若同名图块已存在则直接返回）

    注意：图块内部实体可通过 send_command 执行 BEDIT 等命令编辑。
    """
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    block = await com(cad_service.controller.create_block, block_name, base_point)
    return await _entity_result(block, f"图块 {block_name} 已就绪", f"创建图块 {block_name} 失败")


@mcp.tool()
async def insert_block(ctx: Context, block_name: str, position: Point,
                       x_scale: float = 1.0, y_scale: float = 1.0, z_scale: float = 1.0,
                       rotation_degrees: float = 0.0) -> ToolResult:
    """在模型空间插入图块引用（返回图块引用实体的句柄）"""
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败")
    block_ref = await com(cad_service.controller.insert_block, block_name, position,
                          x_scale, y_scale, z_scale, rotation_degrees)
    return await _entity_result(block_ref, f"图块 {block_name} 已插入", f"插入图块 {block_name} 失败（请确认图块已定义）")


@mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))
async def list_blocks(ctx: Context) -> ListBlocksResult:
    """列出当前图纸中的用户图块定义"""
    if not cad_service.controller.is_running():
        return ListBlocksResult(success=False, message="CAD未运行")
    blocks = await com(cad_service.controller.list_blocks)
    return ListBlocksResult(success=True, message=f"共{len(blocks)}个图块", blocks=blocks)


# ---------- 遗留接口 ----------

@mcp.tool()
async def process_command(ctx: Context, command: str) -> ToolResult:
    """处理自然语言命令并转换为CAD操作（遗留接口）

    推荐优先直接调用结构化绘图工具（draw_line/draw_circle等），
    仅在用户明确要求自然语言入口时使用本工具。
    """
    if not await _ensure_running(ctx):
        return ToolResult(success=False, message="CAD未运行且自动启动失败", error="CAD未运行且自动启动失败")
    result = await com(cad_service.process_command, command)
    # 兼容旧返回结构（entity_id）并映射到统一结果
    handles = []
    if result.get("entity_id"):
        handles = [str(result["entity_id"])]
    return ToolResult(
        success=bool(result.get("success")),
        message=result.get("message", ""),
        entity_handles=handles,
        error=result.get("error") if not result.get("success") else None
    )


# ---------- 资源与提示词 ----------

@mcp.resource("drawing://current", name="当前CAD图纸",
              description="当前CAD图纸的状态（会话绘图记录 + 来自CAD文档的实时统计）",
              mime_type="application/json")
async def current_drawing() -> str:
    """当前CAD图纸状态资源"""
    # 会话记录 + 实时统计，让客户端既看到本会话画了什么，也看到文档真实状态
    stats = await com(cad_service.controller.get_drawing_stats)
    payload = {
        "session": cad_service.drawing_state,
        "live": stats,
    }
    return json.dumps(payload, ensure_ascii=False, default=str)


@mcp.prompt(name="cad-assistant", description="一个用于通过自然语言控制CAD的助手")
def cad_assistant() -> str:
    """CAD助手提示模板"""
    return PROMPT_TEMPLATE.strip()


# ==================== 主入口 ====================

def main():
    """主入口：支持 stdio（默认）与 streamable-http 两种传输方式"""
    parser = argparse.ArgumentParser(description="CAD MCP 服务器")
    parser.add_argument("--transport", choices=["stdio", "streamable-http"], default=None,
                        help="传输方式（默认读取config.json，未配置则为stdio）")
    parser.add_argument("--host", default=None, help="Streamable HTTP 监听地址")
    parser.add_argument("--port", type=int, default=None, help="Streamable HTTP 监听端口")
    args = parser.parse_args()

    server_cfg = config.get("server", {})
    transport = args.transport or server_cfg.get("transport", "stdio")
    host = args.host or server_cfg.get("host", "127.0.0.1")
    port = args.port or int(server_cfg.get("port", 8000))

    if transport == "streamable-http":
        # Streamable HTTP 传输：支持远程/多客户端共享同一CAD实例
        mcp.settings.host = host
        mcp.settings.port = port
        logger.info(f"服务器以 Streamable HTTP 传输运行: http://{host}:{port}/mcp")
        mcp.run(transport="streamable-http")
    else:
        logger.info("服务器以 stdio 传输运行")
        mcp.run()  # 默认 stdio


if __name__ == "__main__":
    main()
