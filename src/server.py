import os
import sys
import json
import logging
from pathlib import Path
from mcp.server.models import InitializationOptions
import mcp.types as types
from mcp.server import NotificationOptions, Server
import mcp.server.stdio
from pydantic import AnyUrl
from typing import Any, Dict, List
import sys

sys.dont_write_bytecode = True

# 使用绝对导入

from nlp_processor import NLPProcessor

from cad_controller import CADController


# 配置Windows环境下的UTF-8编码
if sys.platform == "win32" and os.environ.get('PYTHONIOENCODING') is None:
    sys.stdin.reconfigure(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('cad_mcp.log', encoding='utf-8')
    ]
)


logger = logging.getLogger('mcp_cad_server')
logger.info("启动 CAD MCP 服务器")


PROMPT_TEMPLATE = """
助手的目标是演示如何使用CAD MCP服务。通过这个服务，您可以直接在对话中控制CAD软件，创建和修改图形。

<cad-mcp>

提示：
这个服务器提供了一个名为"cad-assistant"的预设提示，帮助用户通过自然语言指令控制CAD。用户可以要求绘制各种形状、修改图形元素或保存图纸。

资源：
此服务器提供"drawing://current"资源，表示当前的CAD图纸状态。

工具：

此服务器提供以下基本绘图工具：
"draw_line": 在CAD中绘制直线
"draw_circle": 在CAD中绘制圆
"draw_arc": 在CAD中绘制弧
"draw_ellipse": 在CAD中绘制椭圆
"draw_polyline": 在CAD中绘制多段线
"draw_rectangle": 在CAD中绘制矩形
"draw_text": 在CAD中添加文本
"draw_hatch": 在CAD中绘制填充
"save_drawing": 保存当前图纸
"add_dimension": 在CAD中添加线性标注
"process_command": 处理自然语言命令并转换为CAD操作

</cad-mcp>

请以友好的方式开始演示，例如："嗨！今天我将向您展示如何使用CAD MCP服务。通过这个服务，您可以直接在我们的对话中控制CAD软件，无需手动操作界面。让我们开始吧！"
"""

# 后期再做
# "erase_entity": 删除指定的实体
# "move_entity": 移动指定的实体
# "rotate_entity": 旋转指定的实体
# "scale_entity": 缩放指定的实体

class Config:

    def __init__(self):
        # 直接读取config.json文件
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        logger.info("配置文件加载成功")

    @property
    def server_name(self) -> str:
        return self.config['server']['name']

    @property
    def server_version(self) -> str:
        return self.config['server']['version']

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
    
    def add_dimension(self, start_point, end_point, text_position=None,textheight=5, layer=None, color=None):
        """添加线性标注"""
        if not self.controller.is_running():
            self.start_cad()
        
        # 使用当前图层或指定图层
        current_layer = layer or self.drawing_state["current_layer"]
        
        result = self.controller.add_dimension(start_point, end_point, text_position, textheight,current_layer, color)
        if result:
            self.drawing_state["entities"].append({
                "type": "dimension",
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
        """处理自然语言命令"""
        if not self.controller.is_running():
            self.start_cad()
        # 使用NLP处理器解析命令
        parsed_command = self.nlp_processor.process_command(command)
        command_type = parsed_command.get("type")
        try:
            # 基本绘图命令处理
            if command_type == "draw_line":
                start_point = parsed_command.get("start_point")
                end_point = parsed_command.get("end_point")
                # 获取颜色参数
                color = parsed_command.get("color")
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = self.nlp_processor.extract_color_from_command(command)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                # 获取线宽参数
                lineweight = parsed_command.get("lineweight")
                result = self.draw_line(start_point, end_point, None, color, lineweight)
                return {
                    "success": result is not None,
                    "message": "直线已绘制" if result else "绘制直线失败",
                    "entity_id": result.Handle if result else None
                }

            elif command_type == "draw_circle":
                center = parsed_command.get("center")
                radius = parsed_command.get("radius")
                # 获取颜色参数
                color = parsed_command.get("color")
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = self.nlp_processor.extract_color_from_command(command)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                # 获取线宽参数
                lineweight = parsed_command.get("lineweight")
                result = self.draw_circle(center, radius, None, color, lineweight)
                return {
                    "success": result is not None,
                    "message": "圆已绘制" if result else "绘制圆失败",
                    "entity_id": result.Handle if result else None
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
                # 获取颜色参数
                color = parsed_command.get("color")
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = self.nlp_processor.extract_color_from_command(command)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                # 获取线宽参数
                lineweight = parsed_command.get("lineweight")
                result = self.draw_arc(center, radius, start_angle, end_angle, None, color, lineweight)

                return {
                    "success": result is not None,
                    "message": "圆弧已绘制" if result else "绘制圆弧失败",
                    "entity_id": result.Handle if result else None
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
                # 获取颜色参数
                color = parsed_command.get("color")
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = self.nlp_processor.extract_color_from_command(command)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                # 获取线宽参数
                lineweight = parsed_command.get("lineweight")
                result = self.draw_ellipse(center, major_axis, minor_axis, rotation, None, color, lineweight)

                return {
                    "success": result is not None,
                    "message": "椭圆已绘制" if result else "绘制椭圆失败",
                    "entity_id": result.Handle if result else None
                }

            elif command_type == "draw_rectangle":
                corner1 = parsed_command.get("corner1")
                corner2 = parsed_command.get("corner2")
                # 获取颜色参数
                color = parsed_command.get("color")
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = self.nlp_processor.extract_color_from_command(command)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                # 获取线宽参数
                lineweight = parsed_command.get("lineweight")
                result = self.draw_rectangle(corner1, corner2, None, color, lineweight)

                return {
                    "success": result is not None,
                    "message": "矩形已绘制" if result else "绘制矩形失败",
                    "entity_id": result.Handle if result else None
                }

            elif command_type == "draw_text":
                position = parsed_command.get("position")
                text = parsed_command.get("text")
                height = parsed_command.get("height", 2.5)
                rotation = parsed_command.get("rotation", 0)
                # 获取颜色参数
                color = parsed_command.get("color")
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = self.nlp_processor.extract_color_from_command(command)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                result = self.draw_text(position, text, height, rotation, None, color)

                return {
                    "success": result is not None,
                    "message": "文本已添加" if result else "添加文本失败",
                    "entity_id": result.Handle if result else None
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
                # 获取颜色参数，可能是索引或RGB值
                color = parsed_command.get("color")  # 获取颜色参数
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = self.nlp_processor.extract_color_from_command(command)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                result = self.draw_hatch(points, pattern_name, scale, None, color)

                return {
                    "success": result is not None,
                    "message": "填充已绘制" if result else "绘制填充失败",
                    "entity_id": result.Handle if result else None
                }

            # 处理标注
            elif command_type == "add_dimension":
                start_point = parsed_command.get("start_point")
                end_point = parsed_command.get("end_point")
                text_position = parsed_command.get("text_position")
                result = self.controller.add_dimension(start_point, end_point, text_position)
                return {
                    "success": result is not None,
                    "message": "标注已添加" if result else "添加标注失败",
                    "entity_id": result.Handle if result else None
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
                color = parsed_command.get("color", 7)  # 默认白色
                result = self.controller.create_layer(layer_name)  # , color
                return {
                    "success": result,
                    "message": f"图层 {layer_name} 已创建" if result else f"创建图层 {layer_name} 失败"
                }
            
        except Exception as e:
            return {
                "success": False,
                "message": f"处理命令时出错: {str(e)}"
            }

async def main():
    """主入口函数"""
    logger.info("启动 CAD MCP 服务器")

    # 加载配置
    config = Config()
    cad_service = CADService()
    server = Server(config.server_name)

    # 注册处理程序
    logger.debug("注册处理程序")

    @server.list_resources()
    async def handle_list_resources() -> list[types.Resource]:
        logger.debug("处理 list_resources 请求")
        return [
            types.Resource(
                uri=AnyUrl("drawing://current"),
                name="当前CAD图纸",
                description="当前CAD图纸的状态",
                mimeType="application/json",
            )
        ]

    @server.read_resource()
    async def handle_read_resource(uri: AnyUrl) -> str:
        logger.debug(f"处理 read_resource 请求，URI: {uri}")
        if uri.scheme != "drawing":
            logger.error(f"不支持的 URI 协议: {uri.scheme}")
            raise ValueError(f"不支持的 URI 协议: {uri.scheme}")

        path = str(uri).replace("drawing://", "")
        if not path or path != "current":
            logger.error(f"未知的资源路径: {path}")
            raise ValueError(f"未知的资源路径: {path}")

        return json.dumps(cad_service.drawing_state)


    @server.list_tools()
    async def handle_list_tools() -> list[types.Tool]:
        """列出可用工具"""
        return [
            types.Tool(
                name="draw_line",
                description="在CAD中绘制直线",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "start_point": {
                            "type": "array",
                            "description": "起点坐标 [x, y, z]",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3
                        },
                        "end_point": {
                            "type": "array",
                            "description": "终点坐标 [x, y, z]",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3
                        },
                        "layer": {"type": "string", "description": "图层名称（可选）"},
                        "color": {"type": "string", "description": "颜色名称（可选）"},
                        "lineweight": {"type": "number", "description": "线宽（可选）"}
                    },
                    "required": ["start_point", "end_point"],
                },
            ),

            types.Tool(
                name="draw_circle",
                description="在CAD中绘制圆",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "center": {
                            "type": "array",
                            "description": "圆心坐标 [x, y, z]",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3
                        },
                        "radius": {"type": "number", "description": "圆的半径"},
                        "layer": {"type": "string", "description": "图层名称（可选）"},
                        "color": {"type": "string", "description": "颜色名称（可选）"},
                        "lineweight": {"type": "number", "description": "线宽（可选）"}
                    },
                    "required": ["center", "radius"],
                },
            ),

            types.Tool(
                name="draw_arc",
                description="在CAD中绘制弧",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "center": {
                            "type": "array",
                            "description": "圆心坐标 [x, y, z]",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3
                        },
                        "radius": {"type": "number", "description": "弧的半径"},
                        "start_angle": {"type": "number", "description": "起始角度（度）"},
                        "end_angle": {"type": "number", "description": "结束角度（度）"},
                        "layer": {"type": "string", "description": "图层名称（可选）"},
                        "color": {"type": "string", "description": "颜色名称（可选）"},
                        "lineweight": {"type": "number", "description": "线宽（可选）"}
                    },
                    "required": ["center", "radius", "start_angle", "end_angle"],
                },
            ),

            types.Tool(
                name="draw_ellipse",
                description="在CAD中绘制椭圆",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "center": {
                            "type": "array",
                            "description": "椭圆中心坐标 [x, y, z]",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3
                        },
                        "major_axis": {"type": "number", "description": "长轴长度"},
                        "minor_axis": {"type": "number", "description": "短轴长度"},
                        "rotation": {"type": "number", "description": "旋转角度（度）（可选）"},
                        "layer": {"type": "string", "description": "图层名称（可选）"},
                        "color": {"type": "string", "description": "颜色名称（可选）"},
                        "lineweight": {"type": "number", "description": "线宽（可选）"}
                    },
                    "required": ["center", "major_axis", "minor_axis"],
                },
            ),

            types.Tool(
                name="draw_polyline",
                description="在CAD中绘制多段线",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "points": {
                            "type": "array",
                            "description": "点集 [[x1, y1, z1], [x2, y2, z2], ...]",
                            "items": {
                                "type": "array",
                                "items": {"type": "number"},
                                "minItems": 2,
                                "maxItems": 3
                            },
                            "minItems": 2
                        },
                        "closed": {"type": "boolean", "description": "是否闭合"},
                        "layer": {"type": "string", "description": "图层名称（可选）"},
                        "color": {"type": "string", "description": "颜色名称（可选）"},
                        "lineweight": {"type": "number", "description": "线宽（可选）"}
                    },
                    "required": ["points"],
                },
            ),

            types.Tool(
                name="draw_rectangle",
                description="在CAD中绘制矩形",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "corner1": {
                            "type": "array",
                            "description": "第一个角点坐标 [x, y, z]",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3
                        },
                        "corner2": {
                            "type": "array",
                            "description": "第二个角点坐标 [x, y, z]",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3
                        },
                        "layer": {"type": "string", "description": "图层名称（可选）"},
                        "color": {"type": "string", "description": "颜色名称（可选）"},
                        "lineweight": {"type": "number", "description": "线宽（可选）"}
                    },
                    "required": ["corner1", "corner2"],
                },
            ),

            types.Tool(
                name="draw_text",
                description="在CAD中添加文本",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "position": {
                            "type": "array",
                            "description": "插入点坐标 [x, y, z]",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3
                        },
                        "text": {"type": "string", "description": "文本内容"},
                        "height": {"type": "number", "description": "文本高度"},
                        "rotation": {"type": "number", "description": "旋转角度（度）"},
                        "layer": {"type": "string", "description": "图层名称（可选）"},
                        "color": {"type": "string", "description": "颜色名称（可选）"}
                    },
                    "required": ["position", "text"],
                },
            ),
           
            types.Tool(
                name="draw_hatch",
                description="在CAD中绘制填充",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "points": {
                            "type": "array",
                            "description": "填充边界点集 [[x1, y1, z1], [x2, y2, z2], ...]",
                            "items": {
                                "type": "array",
                                "items": {"type": "number"},
                                "minItems": 2,
                                "maxItems": 3
                            },
                            "minItems": 3
                        },
                        "pattern_name": {"type": "string", "description": "填充图案名称（可选，默认为SOLID）"},
                        "scale": {"type": "number", "description": "填充图案比例（可选，默认为1.0）"},
                        "layer": {"type": "string", "description": "图层名称（可选）"},
                        "color": {"type": "string", "description": "颜色名称（可选）"}
                    },
                    "required": ["points"],
                },
            ),

            types.Tool(
                name="add_dimension",
                description="在CAD中添加线性标注",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "start_point": {"type": "array", "description": "起点坐标 [x, y, z]", "items": {"type": "number"}, "minItems": 2, "maxItems": 3},
                        "end_point": {"type": "array", "description": "终点坐标 [x, y, z]", "items": {"type": "number"}, "minItems": 2, "maxItems": 3},
                        "text_position": {"type": "array", "description": "文本位置坐标 [x, y, z]，可选", "items": {"type": "number"}, "minItems": 2, "maxItems": 3},
                        "textheight": {"type": "number", "description": "标注文本高度，可选"},
                        "layer": {"type": "string", "description": "图层名称，可选"},
                        "color": {"type": "string", "description": "颜色名称，可选"}
                    },
                    "required": ["start_point", "end_point"],
                },
            ),

            types.Tool(
                name="save_drawing",
                description="保存当前图纸",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string", "description": "保存路径"}
                    },
                    "required": ["file_path"],
                },
            ),

            types.Tool(
                name="process_command",
                description="处理自然语言命令并转换为CAD操作",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "command": {"type": "string", "description": "自然语言命令"}
                    },
                    "required": ["command"],
                },
            ),
        ]

    @server.call_tool()
    async def handle_call_tool(
        name: str, arguments: dict[str, Any] | None
    ) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
        """处理工具执行请求"""
        try:
            if not arguments:
                raise ValueError("缺少参数")

            if name == "draw_line":
                start_point = arguments.get("start_point")
                end_point = arguments.get("end_point")
                layer = arguments.get("layer")
                color = arguments.get("color") 
                lineweight = arguments.get("lineweight")
                
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                
                if not start_point or not end_point:
                    raise ValueError("缺少起点或终点坐标")
                
                
                result = cad_service.draw_line(start_point, end_point, layer, color, lineweight)
                return [types.TextContent(type="text", text=str(result))]

            elif name == "draw_circle":
                center = arguments.get("center")
                radius = arguments.get("radius")
                layer = arguments.get("layer")
                color = arguments.get("color")
                lineweight = arguments.get("lineweight")
                
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                
                if not center or radius is None:
                    raise ValueError("缺少圆心坐标或半径")
                
                
                result = cad_service.draw_circle(center, radius, layer, color, lineweight)
                return [types.TextContent(type="text", text=str(result))]

            elif name == "draw_arc":
                center = arguments.get("center")
                radius = arguments.get("radius")
                start_angle = arguments.get("start_angle")
                end_angle = arguments.get("end_angle")
                layer = arguments.get("layer")
                color = arguments.get("color")
                lineweight = arguments.get("lineweight")
                
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                
                # 详细检查每个参数并提供具体的错误信息
                error_msgs = []
                if not center:
                    error_msgs.append("中心点坐标")
                if radius is None:
                    error_msgs.append("半径")
                if start_angle is None:
                    error_msgs.append("起始角度")
                if end_angle is None:
                    error_msgs.append("结束角度")
                
                if error_msgs:
                    raise ValueError(f"缺少必要参数: {', '.join(error_msgs)}")                
                
                result = cad_service.draw_arc(center, radius, start_angle, end_angle, layer, color, lineweight)
                return [types.TextContent(type="text", text=str(result))]

            elif name == "draw_ellipse":
                center = arguments.get("center")
                major_axis = arguments.get("major_axis")
                minor_axis = arguments.get("minor_axis")
                rotation = arguments.get("rotation")
                layer = arguments.get("layer")
                color = arguments.get("color")
                lineweight = arguments.get("lineweight")
                
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                
                # 详细检查每个参数并提供具体的错误信息
                error_msgs = []
                if not center:
                    error_msgs.append("中心点坐标")
                if major_axis is None:
                    error_msgs.append("长轴")
                if minor_axis is None:
                    error_msgs.append("短轴")
                
                if error_msgs:
                    raise ValueError(f"缺少必要参数: {', '.join(error_msgs)}")
                
                result = cad_service.draw_ellipse(center, major_axis, minor_axis, rotation, layer, color, lineweight)
                return [types.TextContent(type="text", text=str(result))]

            elif name == "draw_polyline":
                points = arguments.get("points")
                closed = arguments.get("closed", False)
                layer = arguments.get("layer")
                color = arguments.get("color")
                lineweight = arguments.get("lineweight")
                
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                
                if not points or len(points) < 2:
                    raise ValueError("缺少点集或点数不足")                
                
                result = cad_service.draw_polyline(points, closed, layer, color, lineweight)
                return [types.TextContent(type="text", text=str(result))]

            elif name == "draw_rectangle":
                corner1 = arguments.get("corner1")
                corner2 = arguments.get("corner2")
                layer = arguments.get("layer")
                color = arguments.get("color")
                lineweight = arguments.get("lineweight")
                
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                
                if not corner1 or not corner2:
                    raise ValueError("缺少角点坐标")                
                
                result = cad_service.draw_rectangle(corner1, corner2, layer, color, lineweight)
                return [types.TextContent(type="text", text=str(result))]


            elif name == "draw_text":
                position = arguments.get("position")
                text = arguments.get("text")
                height = arguments.get("height", 2.5)
                rotation = arguments.get("rotation", 0)
                layer = arguments.get("layer")
                color = arguments.get("color")
                
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                
                if not position or not text:
                    raise ValueError("缺少插入点坐标或文本内容")
                
                result = cad_service.draw_text(position, text, height, rotation, layer, color)
                return [types.TextContent(type="text", text=str(result))]


            elif name == "draw_hatch":
                points = arguments.get("points")
                pattern_name = arguments.get("pattern_name", "SOLID")
                scale = arguments.get("scale", 1.0)
                layer = arguments.get("layer")
                color = arguments.get("color")
                
                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色
                
                if not points or len(points) < 3:
                    raise ValueError("缺少点集或点数不足，填充边界至少需要3个点")
                
                result = cad_service.draw_hatch(points, pattern_name, scale, layer, color)
                return [types.TextContent(type="text", text=str(result))]

            elif name == "save_drawing":
                file_path = arguments.get("file_path")
                
                if not file_path:
                    raise ValueError("缺少保存路径")
                
                result = cad_service.save_drawing(file_path)
                return [types.TextContent(type="text", text=str(result))]

            elif name == "add_dimension":
                start_point = arguments.get("start_point")
                end_point = arguments.get("end_point")
                text_position = arguments.get("text_position")
                layer = arguments.get("layer")
                color = arguments.get("color")
                textheight = arguments.get("textheight")                

                # 尝试从命令中提取颜色名称并转换为RGB值
                color_rgb = cad_service.nlp_processor.extract_color_from_command(color)
                if color_rgb is not None:
                    color = color_rgb  # 优先使用从命令中提取的颜色                

                # 详细检查每个参数并提供具体的错误信息
                error_msgs = []
                if not start_point:
                    error_msgs.append("起点坐标")
                if not end_point:
                    error_msgs.append("终点坐标")                

                if error_msgs:
                    raise ValueError(f"缺少必要参数: {', '.join(error_msgs)}")                

                result = cad_service.add_dimension(start_point, end_point, text_position, textheight, layer, color)
                return [types.TextContent(type="text", text=str(result))]

            elif name == "process_command":
                command = arguments.get("command")                

                if not command:
                    raise ValueError("缺少命令")                

                result = cad_service.process_command(command)
                return [types.TextContent(type="text", text=str(result))]
            else:
                raise ValueError(f"未知工具: {name}")

        except Exception as e:
            return [types.TextContent(type="text", text=f"错误: {str(e)}")]

    @server.list_prompts()
    async def handle_list_prompts() -> list[types.Prompt]:
        logger.debug("处理 list_prompts 请求")
        return [
            types.Prompt(
                name="cad-assistant",
                description="一个用于通过自然语言控制CAD的助手",
                arguments=[],
            )
        ]

    @server.get_prompt()
    async def handle_get_prompt(name: str, arguments: dict[str, str] | None) -> types.GetPromptResult:
        logger.debug(f"处理 get_prompt 请求，名称: {name}，参数: {arguments}")
        if name != "cad-assistant":
            logger.error(f"未知的提示: {name}")
            raise ValueError(f"未知的提示: {name}")

        prompt = PROMPT_TEMPLATE
        logger.debug("生成提示模板")

        return types.GetPromptResult(
            description="CAD助手提示模板",
            messages=[
                types.PromptMessage(
                    role="user",
                    content=types.TextContent(type="text", text=prompt.strip()),
                )
            ],
        )

    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        logger.info("服务器正在使用 stdio 传输运行")
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name=config.server_name,
                server_version=config.server_version,
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())