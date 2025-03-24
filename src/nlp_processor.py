import logging
import re
import math
from typing import Any, Dict, List, Optional, Tuple
import json
import os
import os.path

# 直接读取config.json文件
config_path = os.path.join(os.path.dirname(__file__), 'config.json')
with open(config_path, 'r', encoding='utf-8') as f:
    config = json.load(f)

logger = logging.getLogger('nlp_processor')

class NLPProcessor:
    """自然语言处理器类，负责解析用户指令并转换为CAD操作"""
    
    def __init__(self):
        """
        初始化NLP处理器
        """
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("自然语言处理器已初始化")
        
        # 详细颜色名称到RGB值的映射
        self.color_rgb_map = {
            # 基本颜色
            "红色": 1,
            "黄色": 2,
            "绿色": 3,
            "青色": 4,
            "蓝色": 5,
            "洋红色": 6,
            "白色": 7,
            "灰色": 8,
            "浅灰色": 9,
            "黑色": 250,
            "棕色": 251,
            "橙色": 30,
            "紫色": 200,

            # 基本颜色 - 英文键
            "Red": 1,
            "Yellow": 2,
            "Green": 3,
            "Cyan": 4,
            "Blue": 5,
            "Magenta": 6,
            "White": 7,
            "Gray": 8,
            "Light Gray": 9,
            "Black": 250,
            "Brown": 251,
            "Orange": 30,
            "Purple": 200
        }
        
        # 扩展关键词映射
        self.shape_keywords = {
            # 基本形状
            "直线": "line", "线": "line",
            "圆": "circle", "圆形": "circle",
            "弧": "arc", "圆弧": "arc",
            "矩形": "rectangle", "方形": "rectangle", "正方形": "square",
            "多段线": "polyline", "折线": "polyline",
            "文本": "text", "文字": "text",
            "标注": "dimension", "尺寸标注": "dimension",
            
            # 建筑元素
            "墙": "wall", "墙体": "wall",
            "门": "door", "窗": "window",
            "楼梯": "stair", "柱子": "column",
            
            # 电气元素
            "插座": "outlet",
            "开关": "switch",
            "灯": "light", "灯具": "light",
            "配电箱": "distribution_box",
            
            # 机械元素
            "轴": "shaft",
            "齿轮": "gear",
            "轴承": "bearing",
            "法兰": "flange"
        }
        
        # 动作关键词映射
        self.action_keywords = {
            # 基本动作
            "画": "draw", "绘制": "draw", "创建": "draw", "添加": "draw",
            "修改": "modify", "调整": "modify", "改变": "modify",
            "移动": "move", "旋转": "rotate", "缩放": "scale",
            "放大": "scale_up", "缩小": "scale_down",
            "删除": "erase", "擦除": "erase", "移除": "erase",
            "保存": "save",
            
            # 专业操作
            "标注": "dimension",
            "填充": "hatch",
            "创建图层": "create_layer",
            "切换图层": "change_layer"
        }
        
        # # 专业领域映射
        # self.domain_keywords = {
        #     "建筑": "architecture",
        #     "电气": "electrical",
        #     "机械": "mechanical",
        #     "土木": "civil",
        #     "管道": "piping",
        #     "结构": "structural"
        # }
    
      
    def extract_color_from_command(self, command: str) -> Optional[int]:
       
        if command is None:
            return 7

        try:
            num = int(command)
            if num >= 1 and num <= 255:
                return num
        except:
            pass

        # 将命令转换为小写
        command = command.lower()
        
        # 尝试匹配颜色名称
        for color_name in self.color_rgb_map.keys():
            if color_name.lower() in command:
                return self.color_rgb_map[color_name]
        
        # 尝试匹配颜色描述（如"淡蓝色"）
        color_pattern = r'([深浅淡]?[a-zA-Z\u4e00-\u9fa5]+色)'
        color_matches = re.findall(color_pattern, command)
        
        for color_match in color_matches:
            # 检查是否是已知的颜色名称
            if color_match in self.color_rgb_map:
                return self.color_rgb_map[color_match]
                    
        # 如果找不到颜色信息，返回7 默认白色
        return 7
    
    def process_command(self, command: str) -> Dict[str, Any]:
        """处理自然语言命令并返回结果"""
        self.logger.info(f"处理命令: {command}")
        
        # 解析命令
        parsed_command = self.parse_command(command)
      
        # 返回解析结果
        return parsed_command
    
    def parse_command(self, command: str) -> Dict[str, Any]:
        """解析自然语言命令并转换为CAD操作参数"""
        self.logger.info(f"解析命令: {command}")
        
        # 将命令转换为小写并去除多余空格
        command = command.lower().strip()
        
        # 尝试识别命令类型
        command_type = self._identify_command_type(command)
        self.logger.debug(f"识别到的命令类型: {command_type}")
        
        # 根据命令类型分发到不同的处理函数
        if command_type == "draw_line":
            return self._parse_draw_line(command)
        elif command_type == "draw_circle":
            return self._parse_draw_circle(command)
        elif command_type == "draw_arc":
            return self._parse_draw_arc(command)
        elif command_type == "draw_rectangle":
            return self._parse_draw_rectangle(command)
        elif command_type == "draw_polyline":
            return self._parse_draw_polyline(command)
        elif command_type == "draw_text":
            return self._parse_draw_text(command)
        elif command_type == "draw_hatch":
            return self._parse_draw_hatch(command)
        # elif command_type == "erase":
        #     return self._parse_erase(command)
        # elif command_type == "move":
        #     return self._parse_move(command)
        # elif command_type == "rotate":
        #     return self._parse_rotate(command)
        # elif command_type == "scale":
        #     return self._parse_scale(command)
        elif command_type == "save":
            return self._parse_save(command)
        elif command_type == "draw_wall":
            return self._parse_draw_wall(command)
        elif command_type == "create_layer":
            return self._parse_create_layer(command)
        # elif command_type == "draw_electrical_symbol":
        #     return self._parse_draw_electrical_symbol(command)
        else:
            # 默认返回一个错误结果
            return {
                "type": "unknown",
                "error": "无法识别的命令类型",
                "original_command": command
            }
    
    def _identify_command_type(self, command: str) -> str:
        """识别命令类型"""
        # 首先检查专业领域
        domain = None
        for domain_keyword, domain_type in self.domain_keywords.items():
            if domain_keyword in command:
                domain = domain_type
                break
        
        # 检查操作类型
        for action, action_type in self.action_keywords.items():
            if action in command:
                # 建筑领域特殊处理
                # if domain == "architecture":
                #     for shape, shape_type in self.shape_keywords.items():
                #         if shape in command:
                #             if shape_type == "wall":
                #                 return "draw_wall"
                #             elif shape_type == "door":
                #                 return "draw_door"
                #             elif shape_type == "window":
                #                 return "draw_window"
                
                # # 电气领域特殊处理
                # elif domain == "electrical":
                #     for shape, shape_type in self.shape_keywords.items():
                #         if shape in command:
                #             if shape_type in ["outlet", "switch", "light"]:
                #                 return f"draw_electrical_{shape_type}"
                
                # 基本形状处理
                for shape, shape_type in self.shape_keywords.items():
                    if shape in command:
                        if action_type == "draw":
                            if shape_type == "line":
                                return "draw_line"
                            elif shape_type == "circle":
                                return "draw_circle"
                            elif shape_type == "arc":
                                return "draw_arc"
                            elif shape_type in ["rectangle", "square"]:
                                return "draw_rectangle"
                            elif shape_type == "polyline":
                                return "draw_polyline"
                            elif shape_type == "text":
                                return "draw_text"
                            elif shape_type == "dimension":
                                return "add_dimension"
        
        # 检查是否是创建图层命令
        if "图层" in command and any(action in command for action in ["创建", "新建", "添加"]):
            return "create_layer"
        
        # 检查是否是标注命令
        if "标注" in command:
            return "add_dimension"
        
        # # 检查其他通用命令
        # if any(word in command for word in ["删除", "擦除", "移除"]):
        #     return "erase"
        
        # if any(word in command for word in ["移动", "挪动"]):
        #     return "move"
        
        # if "旋转" in command:
        #     return "rotate"
        
        # if any(word in command for word in ["缩放", "放大", "缩小"]):
        #     return "scale"
        
        if "保存" in command:
            return "save"
        
        # 如果无法识别，返回未知类型
        return "unknown"
    
    def _extract_coordinates(self, text: str) -> List[Tuple[float, float, float]]:
        """从文本中提取坐标点"""
        # 匹配坐标格式: (x,y,z) 或 (x,y) 或 x,y,z 或 x,y
        pattern = r'\(?\s*(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)(?:\s*,\s*(-?\d+\.?\d*))?\s*\)?'
        matches = re.finditer(pattern, text)
        
        coordinates = []
        for match in matches:
            x = float(match.group(1))
            y = float(match.group(2))
            z = float(match.group(3)) if match.group(3) else 0.0
            coordinates.append((x, y, z))
        
        return coordinates
    
    def _extract_numbers(self, text: str) -> List[float]:
        """从文本中提取数字"""
        pattern = r'(-?\d+\.?\d*)'
        matches = re.findall(pattern, text)
        return [float(match) for match in matches]
    
    def _parse_draw_line(self, command: str) -> Dict[str, Any]:
        """解析绘制直线命令"""
        # 尝试提取坐标
        coordinates = self._extract_coordinates(command)
        
        if len(coordinates) >= 2:
            # 如果找到至少两个坐标点，使用前两个作为起点和终点
            return {
                "type": "draw_line",
                "start_point": coordinates[0],
                "end_point": coordinates[1]
            }
        else:
            # 如果没有找到足够的坐标点，尝试使用默认值
            # 这里可以根据需要设置默认的起点和终点
            return {
                "type": "draw_line",
                "start_point": (0, 0, 0),
                "end_point": (100, 100, 0),
                "note": "使用默认坐标，因为命令中未提供足够的坐标信息"
            }
    
    def _parse_draw_circle(self, command: str) -> Dict[str, Any]:
        """解析绘制圆命令"""
        # 尝试提取坐标和半径
        coordinates = self._extract_coordinates(command)
        numbers = self._extract_numbers(command)
        
        # 提取半径
        radius = None
        radius_pattern = r'(?:半径|r|radius)[^\d]*?(-?\d+\.?\d*)'
        radius_match = re.search(radius_pattern, command, re.IGNORECASE)
        if radius_match:
            radius = float(radius_match.group(1))
        elif len(numbers) > 0:
            # 如果没有明确指定半径，使用找到的第一个数字作为半径
            radius = numbers[0]
        else:
            # 默认半径
            radius = 50.0
        
        # 提取中心点
        center = None
        if len(coordinates) > 0:
            center = coordinates[0]
        else:
            # 默认中心点
            center = (0, 0, 0)
        
        return {
            "type": "draw_circle",
            "center": center,
            "radius": radius
        }
    
    def _parse_draw_arc(self, command: str) -> Dict[str, Any]:
        """解析绘制圆弧命令"""
        # 尝试提取坐标、半径和角度
        coordinates = self._extract_coordinates(command)
        numbers = self._extract_numbers(command)
        
        # 提取中心点
        center = None
        if len(coordinates) > 0:
            center = coordinates[0]
        else:
            # 默认中心点
            center = (0, 0, 0)
        
        # 提取半径
        radius = None
        radius_pattern = r'(?:半径|r|radius)[^\d]*?(-?\d+\.?\d*)'
        radius_match = re.search(radius_pattern, command, re.IGNORECASE)
        if radius_match:
            radius = float(radius_match.group(1))
        elif len(numbers) > 0:
            # 如果没有明确指定半径，使用找到的第一个数字作为半径
            radius = numbers[0]
        else:
            # 默认半径
            radius = 50.0
        
        # 提取起始角度和结束角度
        start_angle = 0.0
        end_angle = 90.0
        
        start_angle_pattern = r'(?:起始角度|start angle)[^\d]*?(-?\d+\.?\d*)'
        start_angle_match = re.search(start_angle_pattern, command, re.IGNORECASE)
        if start_angle_match:
            start_angle = float(start_angle_match.group(1))
        
        end_angle_pattern = r'(?:结束角度|end angle)[^\d]*?(-?\d+\.?\d*)'
        end_angle_match = re.search(end_angle_pattern, command, re.IGNORECASE)
        if end_angle_match:
            end_angle = float(end_angle_match.group(1))
        
        # 确保所有必要参数都有值
        if center is None:
            center = (0, 0, 0)
        if radius is None:
            radius = 50.0
        if start_angle is None:
            start_angle = 0.0
        if end_angle is None:
            end_angle = 90.0
            
        self.logger.debug(f"解析圆弧命令结果: 中心点={center}, 半径={radius}, 起始角度={start_angle}, 结束角度={end_angle}")
        
        return {
            "type": "draw_arc",
            "center": center,
            "radius": radius,
            "start_angle": start_angle,
            "end_angle": end_angle
        }
    
    def _parse_draw_rectangle(self, command: str) -> Dict[str, Any]:
        """解析绘制矩形命令"""
        # 尝试提取坐标
        coordinates = self._extract_coordinates(command)
        
        if len(coordinates) >= 2:
            # 如果找到至少两个坐标点，使用前两个作为对角点
            return {
                "type": "draw_rectangle",
                "corner1": coordinates[0],
                "corner2": coordinates[1]
            }
        else:
            # 如果没有找到足够的坐标点，尝试提取宽度和高度
            width = 100.0
            height = 100.0
            
            width_pattern = r'(?:宽度|width)[^\d]*?(-?\d+\.?\d*)'
            width_match = re.search(width_pattern, command, re.IGNORECASE)
            if width_match:
                width = float(width_match.group(1))
            
            height_pattern = r'(?:高度|height)[^\d]*?(-?\d+\.?\d*)'
            height_match = re.search(height_pattern, command, re.IGNORECASE)
            if height_match:
                height = float(height_match.group(1))
            
            # 如果找到一个坐标点，使用它作为起点
            if len(coordinates) == 1:
                corner1 = coordinates[0]
                corner2 = (corner1[0] + width, corner1[1] + height, corner1[2])
            else:
                # 默认起点和终点
                corner1 = (0, 0, 0)
                corner2 = (width, height, 0)
            
            return {
                "type": "draw_rectangle",
                "corner1": corner1,
                "corner2": corner2
            }
    
    def _parse_draw_polyline(self, command: str) -> Dict[str, Any]:
        """解析绘制多段线命令"""
        # 尝试提取坐标
        coordinates = self._extract_coordinates(command)
        
        # 检查是否需要闭合
        closed = "闭合" in command or "封闭" in command
        
        if len(coordinates) >= 2:
            # 如果找到至少两个坐标点，使用它们作为多段线的点
            return {
                "type": "draw_polyline",
                "points": coordinates,
                "closed": closed
            }
        else:
            # 如果没有找到足够的坐标点，返回错误
            return {
                "type": "error",
                "message": "绘制多段线需要至少两个坐标点"
            }
    
    def _parse_draw_text(self, command: str) -> Dict[str, Any]:
        """解析绘制文本命令"""
        # 尝试提取坐标
        coordinates = self._extract_coordinates(command)
        
        # 提取文本内容
        text_pattern = r'[文本内容|text|内容][：:]\s*[\"\'](.*?)[\"\']'
        text_match = re.search(text_pattern, command)
        
        text = ""
        if text_match:
            text = text_match.group(1)
        else:
            # 尝试提取引号中的内容作为文本
            quote_pattern = r'[\"\'](.*?)[\"\']'
            quote_match = re.search(quote_pattern, command)
            if quote_match:
                text = quote_match.group(1)
            else:
                # 如果没有找到引号中的内容，使用默认文本
                text = "示例文本"
        
        # 提取文本高度
        height = 2.5  # 默认高度
        height_pattern = r'(?:高度|height)[^\d]*?(-?\d+\.?\d*)'
        height_match = re.search(height_pattern, command, re.IGNORECASE)
        if height_match:
            height = float(height_match.group(1))
        
        # 提取旋转角度
        rotation = 0.0  # 默认角度
        rotation_pattern = r'(?:旋转|角度|rotation)[^\d]*?(-?\d+\.?\d*)'
        rotation_match = re.search(rotation_pattern, command, re.IGNORECASE)
        if rotation_match:
            rotation = float(rotation_match.group(1))
        
        # 提取插入点
        position = None
        if len(coordinates) > 0:
            position = coordinates[0]
        else:
            # 默认插入点
            position = (0, 0, 0)
        
        return {
            "type": "draw_text",
            "position": position,
            "text": text,
            "height": height,
            "rotation": rotation
        }
    
    def _parse_draw_hatch(self, command: str) -> Dict[str, Any]:
        """解析绘制填充命令"""
        # 尝试提取坐标点集
        coordinates = self._extract_coordinates(command)
        
        # 提取填充图案名称
        pattern_name = "SOLID"  # 默认为实体填充
        pattern_patterns = [
            r'(?:图案|pattern)[^\w]*?["\'](.*?)["\']\'',
            r'(?:图案|pattern)[^\w]*?(\w+)'
        ]
        
        for pattern in pattern_patterns:
            pattern_match = re.search(pattern, command, re.IGNORECASE)
            if pattern_match:
                pattern_name = pattern_match.group(1).upper()
                break
        
        # 提取填充比例
        scale = 1.0  # 默认比例
        scale_pattern = r'(?:比例|缩放|scale)[^\d]*?(\d+\.?\d*)'
        scale_match = re.search(scale_pattern, command, re.IGNORECASE)
        if scale_match:
            scale = float(scale_match.group(1))
        
        # 检查是否有足够的点来定义填充边界
        if len(coordinates) >= 3:
            self.logger.debug(f"解析填充命令结果: 点集={coordinates}, 图案={pattern_name}, 比例={scale}")
            return {
                "type": "draw_hatch",
                "points": coordinates,
                "pattern_name": pattern_name,
                "scale": scale
            }
        else:
            # 如果没有足够的点，返回错误信息
            self.logger.warning("填充命令解析失败: 需要至少3个点来定义填充边界")
            return {
                "type": "error",
                "message": "绘制填充需要至少3个点来定义边界"
            }

    # def _parse_erase(self, command: str) -> Dict[str, Any]:
    #     """解析删除命令"""
    #     # 尝试提取实体ID
    #     entity_id_pattern = r'(?:ID|编号|实体)[^\w]*?([a-fA-F0-9]+)'
    #     entity_id_match = re.search(entity_id_pattern, command, re.IGNORECASE)
        
    #     if entity_id_match:
    #         entity_id = entity_id_match.group(1)
    #         return {
    #             "type": "erase",
    #             "entity_id": entity_id
    #         }
    #     else:
    #         # 如果没有找到实体ID，返回错误
    #         return {
    #             "type": "error",
    #             "message": "删除命令需要指定实体ID"
    #         }
    
    # def _parse_move(self, command: str) -> Dict[str, Any]:
    #     """解析移动命令"""
    #     # 尝试提取实体ID
    #     entity_id_pattern = r'(?:ID|编号|实体)[^\w]*?([a-fA-F0-9]+)'
    #     entity_id_match = re.search(entity_id_pattern, command, re.IGNORECASE)
        
    #     # 尝试提取坐标
    #     coordinates = self._extract_coordinates(command)
        
    #     if entity_id_match and len(coordinates) >= 2:
    #         entity_id = entity_id_match.group(1)
    #         from_point = coordinates[0]
    #         to_point = coordinates[1]
            
    #         return {
    #             "type": "move",
    #             "entity_id": entity_id,
    #             "from_point": from_point,
    #             "to_point": to_point
    #         }
    #     else:
    #         # 如果没有找到足够的信息，返回错误
    #         return {
    #             "type": "error",
    #             "message": "移动命令需要指定实体ID和两个坐标点"
    #         }
    
    # def _parse_rotate(self, command: str) -> Dict[str, Any]:
    #     """解析旋转命令"""
    #     # 尝试提取实体ID
    #     entity_id_pattern = r'(?:ID|编号|实体)[^\w]*?([a-fA-F0-9]+)'
    #     entity_id_match = re.search(entity_id_pattern, command, re.IGNORECASE)
        
    #     # 尝试提取坐标
    #     coordinates = self._extract_coordinates(command)
        
    #     # 提取旋转角度
    #     rotation = 0.0  # 默认角度
    #     rotation_pattern = r'(?:旋转|角度|rotation)[^\d]*?(-?\d+\.?\d*)'
    #     rotation_match = re.search(rotation_pattern, command, re.IGNORECASE)
    #     if rotation_match:
    #         rotation = float(rotation_match.group(1))
        
    #     if entity_id_match and len(coordinates) >= 1:
    #         entity_id = entity_id_match.group(1)
    #         base_point = coordinates[0]
            
    #         return {
    #             "type": "rotate",
    #             "entity_id": entity_id,
    #             "base_point": base_point,
    #             "rotation_angle": rotation
    #         }
    #     else:
    #         # 如果没有找到足够的信息，返回错误
    #         return {
    #             "type": "error",
    #             "message": "旋转命令需要指定实体ID、基点和旋转角度"
    #         }
    
    # def _parse_scale(self, command: str) -> Dict[str, Any]:
        """解析缩放命令"""
        # 尝试提取实体ID
        entity_id_pattern = r'(?:ID|编号|实体)[^\w]*?([a-fA-F0-9]+)'
        entity_id_match = re.search(entity_id_pattern, command, re.IGNORECASE)
        
        # 尝试提取坐标
        coordinates = self._extract_coordinates(command)
        
        # 提取缩放比例
        scale_factor = 1.0  # 默认比例
        scale_pattern = r'(?:比例|缩放|scale)[^\d]*?(-?\d+\.?\d*)'
        scale_match = re.search(scale_pattern, command, re.IGNORECASE)
        if scale_match:
            scale_factor = float(scale_match.group(1))
        elif "放大" in command:
            scale_factor = 2.0
        elif "缩小" in command:
            scale_factor = 0.5
        
        if entity_id_match and len(coordinates) >= 1:
            entity_id = entity_id_match.group(1)
            base_point = coordinates[0]
            
            return {
                "type": "scale",
                "entity_id": entity_id,
                "base_point": base_point,
                "scale_factor": scale_factor
            }
        else:
            # 如果没有找到足够的信息，返回错误
            return {
                "type": "error",
                "message": "缩放命令需要指定实体ID、基点和缩放比例"
            }
    
    def _parse_save(self, command: str) -> Dict[str, Any]:
        """解析绘制墙体命令"""
        # 尝试提取坐标
        coordinates = self._extract_coordinates(command)
        
        # 提取墙体宽度
        width = 10.0  # 默认宽度
        width_pattern = r'(?:宽度|宽|厚|厚度)[^\d]*?(\d+\.?\d*)'
        width_match = re.search(width_pattern, command)
        if width_match:
            width = float(width_match.group(1))
        
        if len(coordinates) >= 2:
            start_point = coordinates[0]
            end_point = coordinates[1]
            
            return {
                "type": "draw_wall",
                "start_point": start_point,
                "end_point": end_point,
                "width": width
            }
        else:
            # 没有足够的坐标信息，返回错误
            return {
                "type": "error",
                "message": "绘制墙体需要提供起点和终点坐标"
            }
    
    def _parse_create_layer(self, command: str) -> Dict[str, Any]:
        """解析保存命令"""
        # 尝试提取文件路径
        path_pattern = r'(?:路径|保存到|path)[^\w]*?[\"\'](.*?)[\"\']'
        path_match = re.search(path_pattern, command, re.IGNORECASE)
        
        if path_match:
            file_path = path_match.group(1)
        else:
            # 默认文件路径
            file_path = os.path.join(config["output"]["directory"], config["output"]["default_filename"])
        
        return {
            "type": "save",
            "file_path": file_path
        }
    
    # def _parse_draw_wall(self, command: str) -> Dict[str, Any]:
    #     """解析创建图层命令"""
    #     # 尝试提取图层名称
    #     layer_name_pattern = r'(?:名称|名字|叫)[^\w]*?[\"\'](.*?)[\"\']'
    #     layer_name_match = re.search(layer_name_pattern, command)
        
    #     if not layer_name_match:
    #         # 尝试另一种模式
    #         layer_name_pattern = r'图层[^\w]*?[\"\'](.*?)[\"\']'
    #         layer_name_match = re.search(layer_name_pattern, command)
        
    #     # 提取颜色
    #     color = 7  # 默认白色
    #     # 基本颜色映射（CAD颜色索引）
    #     color_map = {
    #         "红": 1, "红色": 1,
    #         "黄": 2, "黄色": 2,
    #         "绿": 3, "绿色": 3,
    #         "青": 4, "青色": 4,
    #         "蓝": 5, "蓝色": 5,
    #         "洋红": 6, "洋红色": 6, "紫": 6, "紫色": 6,
    #         "白": 7, "白色": 7,
    #         "灰": 8, "灰色": 8
    #     }
        
    #     for color_name, color_code in color_map.items():
    #         if color_name in command:
    #             color = color_code
    #             break
        
    #     # 尝试提取RGB颜色值
    #     rgb_color = self.extract_color_from_command(command)
    #     if rgb_color is not None:
    #         # 如果找到了RGB颜色值，优先使用它
    #         color = rgb_color
        
    #     if layer_name_match:
    #         layer_name = layer_name_match.group(1)
    #         return {
    #             "type": "create_layer",
    #             "layer_name": layer_name,
    #             "color": color
    #         }
    #     else:
    #         # 如果没有找到图层名称，尝试使用默认名称
    #         return {
    #             "type": "create_layer",
    #             "layer_name": "新图层",
    #             "color": color,
    #             "note": "使用默认图层名称，因为命令中未提供图层名称"
    #         }
    
    # def _parse_draw_electrical_symbol(self, command: str) -> Dict[str, Any]:
    #     """解析绘制电气符号命令"""
    #     # 尝试提取坐标
    #     coordinates = self._extract_coordinates(command)
        
    #     # 提取缩放比例
    #     scale = 1.0  # 默认比例
    #     scale_pattern = r'(?:比例|缩放|scale)[^\d]*?(\d+\.?\d*)'
    #     scale_match = re.search(scale_pattern, command, re.IGNORECASE)
    #     if scale_match:
    #         scale = float(scale_match.group(1))
        
    #     # 提取旋转角度
    #     rotation = 0.0  # 默认角度
    #     rotation_pattern = r'(?:旋转|角度|rotation)[^\d]*?(-?\d+\.?\d*)'
    #     rotation_match = re.search(rotation_pattern, command, re.IGNORECASE)
    #     if rotation_match:
    #         rotation = float(rotation_match.group(1))
        
    #     if len(coordinates) >= 1:
    #         insertion_point = coordinates[0]
            
    #         return {
    #             "type": "draw_electrical_symbol",
    #             "insertion_point": insertion_point,
    #             "scale": scale,
    #             "rotation": rotation
    #         }
    #     else:
    #         # 没有坐标信息，返回错误
    #         return {
    #             "type": "error",
            # }
