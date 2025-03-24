import logging
import math
import time
import os
import json
from typing import Any, Dict, List, Optional, Tuple, Union

# 直接读取config.json文件
config_path = os.path.join(os.path.dirname(__file__), 'config.json')
with open(config_path, 'r', encoding='utf-8') as f:
    config = json.load(f)

try:
    import win32com.client
    # pythoncom是pywin32的一部分，不需要单独安装
    import pythoncom
except ImportError:
    logging.error("无法导入win32com.client或pythoncom，请确保已安装pywin32库")
    raise

logger = logging.getLogger('cad_controller')

class CADController:
    """CAD控制器类，负责与CAD应用程序交互"""
    
    def __init__(self):
        """初始化CAD控制器"""
        self.app = None
        self.doc = None
        self.entities = {}  # 存储已创建图形的实体引用，用于后续修改
        # 从配置文件加载参数
        self.startup_wait_time = config["cad"]["startup_wait_time"]
        self.command_delay = config["cad"]["command_delay"]
        # 获取CAD类型
        self.cad_type = config["cad"]["type"]
        # 有效的线宽值列表
        self.valid_lineweights = [0, 5, 9, 13, 15, 18, 20, 25, 30, 35, 40, 50, 53, 60, 70, 80, 90, 100, 106, 120, 140, 158, 200, 211]
        logger.info("CAD控制器已初始化")
    
    def start_cad(self) -> bool:
        """启动CAD并创建或打开一个文档"""
        try:
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 存储旧实例引用（如果有）以便后续清理
            old_app = None
            if self.app is not None:
                old_app = self.app
                self.app = None
                self.doc = None
            
            try:
                # 根据配置的CAD类型选择不同的应用程序标识符
                app_id = "AutoCAD.Application"
                app_name = "AutoCAD"
                
                if self.cad_type.lower() == "autocad":
                    app_id = "AutoCAD.Application"
                    app_name = "AutoCAD"
                elif self.cad_type.lower() == "gcad":
                    app_id = "GCAD.Application"
                    app_name = "浩辰CAD"
                elif self.cad_type.lower() == "gstarcad":
                    app_id = "GCAD.Application"
                    app_name = "浩辰CAD"
                elif self.cad_type.lower() == "zwcad":
                    app_id = "ZWCAD.Application"
                    app_name = "中望CAD"
                
                # 尝试连接到已运行的CAD实例
                logger.info(f"尝试连接现有{app_name}实例...")
                try:
                    self.app = win32com.client.GetActiveObject(app_id)
                    logger.info(f"成功连接到已运行的{app_name}实例")
                except Exception as e:
                    logger.info(f"未找到运行中的{app_name}实例，将尝试启动新实例: {str(e)}")  
                    raise

                # 已在上面的代码中处理
                
                # 如果当前没有文档，创建一个新文档
                try:
                    if self.app.Documents.Count == 0:
                        logger.info("创建新文档...")
                        self.doc = self.app.Documents.Add()
                    else:
                        logger.info("获取活动文档...")
                        self.doc = self.app.ActiveDocument
                except Exception as doc_ex:
                    # 如果获取文档失败，强制创建新文档
                    logger.warning(f"获取文档失败，尝试创建新文档: {str(doc_ex)}")
                    try:
                        # 关闭所有打开的文档
                        for i in range(self.app.Documents.Count):
                            try:
                                self.app.Documents.Item(0).Close(False)  # 不保存
                            except:
                                pass
                        
                        # 创建新文档
                        self.doc = self.app.Documents.Add()
                    except Exception as new_doc_ex:
                        logger.error(f"创建新文档失败: {str(new_doc_ex)}")
                        raise
                    
            except Exception as app_ex:
                # 如果连接失败，启动一个新实例
                logger.info(f"连接失败，正在启动新的CAD实例: {str(app_ex)}")
                try:
                    # 根据配置的CAD类型启动相应的应用程序
                    app_id = "AutoCAD.Application"
                    app_name = "AutoCAD"
                    
                    if self.cad_type.lower() == "autocad":
                        app_id = "AutoCAD.Application"
                        app_name = "AutoCAD"
                    elif self.cad_type.lower() == "gcad":
                        app_id = "GCAD.Application"
                        app_name = "浩辰CAD"
                    elif self.cad_type.lower() == "gstarcad":
                        app_id = "GCAD.Application"
                        app_name = "浩辰CAD"
                    elif self.cad_type.lower() == "zwcad":
                        app_id = "ZWCAD.Application"
                        app_name = "中望CAD"
                    
                    logger.info(f"正在启动{app_name}实例...")
                    self.app = win32com.client.Dispatch(app_id)
                    self.app.Visible = True
                    
                    # 等待CAD启动
                    time.sleep(self.startup_wait_time)  # 使用配置的等待时间
                    
                    # 创建新文档
                    logger.info("尝试创建新文档...")
                    # self.doc = self.app.Documents.Add()
                    self.doc = self.app.ActiveDocument
                except Exception as new_app_ex:
                    logger.error(f"启动新CAD实例失败: {str(new_app_ex)}")
                    raise
            
            # 额外安全检查和等待
            time.sleep(2)  # 给CAD更多时间处理文档创建
            
            if self.doc is None:
                raise Exception("无法获取有效的Document对象")
            
            # 尝试读取文档属性以验证其有效性
            try:
                name = self.doc.Name
                logger.info(f"文档名称: {name}")
            except Exception as name_ex:
                logger.error(f"无法读取文档名称: {str(name_ex)}")
                raise Exception("文档对象无效")
            
            logger.info("CAD已成功启动和准备")
            return True
            
        except Exception as e:
            logger.error(f"启动CAD失败: {str(e)}")
            return False
        finally:
            # 清理旧实例
            if old_app is not None:
                try:
                    del old_app
                except:
                    pass
    
        
    def is_running(self) -> bool:
        """检查CAD是否正在运行"""
        return self.app is not None and self.doc is not None
    
    def save_drawing(self, file_path: str) -> bool:
        """保存当前图纸到指定路径"""
        if not self.is_running():
            logger.error("CAD未运行，无法保存图纸")
            return False
            
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            # 保存文件
            self.doc.SaveAs(file_path)
            logger.info(f"图纸已保存到: {file_path}")
            return True
        except Exception as e:
            logger.error(f"保存图纸失败: {str(e)}")
            return False
    
    # def clear_entities(self, shape_name: str) -> None:
    #     """清除特定形状的所有实体"""
    #     if shape_name in self.entities:
    #         for entity in self.entities[shape_name]:
    #             try:
    #                 entity.Delete()
    #             except:
    #                 pass  # 实体可能已被删除
    #         self.entities[shape_name] = []
    
    # def get_entity_group(self, shape_name: str) -> List:
    #     """获取特定形状的实体组，如果不存在则创建"""
    #     if shape_name not in self.entities:
    #         self.entities[shape_name] = []
    #     return self.entities[shape_name]
    
    def refresh_view(self) -> None:
        """刷新CAD视图"""
        if self.is_running():
            try:
                self.doc.Regen(1)  # acAllViewports = 1
            except Exception as e:
                logger.error(f"刷新视图失败: {str(e)}")
    
    def validate_lineweight(self, lineweight) -> int:
        """验证并返回有效的线宽值
        
        如果提供的线宽值不在有效值列表中，则返回默认值0
        
        Args:
            lineweight: 要验证的线宽值
            
        Returns:
            有效的线宽值
        """
        if lineweight is None:
            return None
            
        # 检查线宽是否在有效值列表中
        if lineweight in self.valid_lineweights:
            return lineweight
        else:
            logger.warning(f"线宽值 {lineweight} 无效，将使用默认值 0")
            return 0
    
    def draw_line(self, start_point: Tuple[float, float, float], 
                 end_point: Tuple[float, float, float], layer: str = None, color: int = None, lineweight=None) -> Any:
        """绘制直线"""
        if not self.is_running():
            return None
            
        try:
            # 确保点是三维的
            if len(start_point) == 2:
                start_point = (start_point[0], start_point[1], 0)
            if len(end_point) == 2:
                end_point = (end_point[0], end_point[1], 0)
      
            # 使用VARIANT包装坐标点数据
            start_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [start_point[0], start_point[1], start_point[2]])
            end_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                             [end_point[0], end_point[1], end_point[2]])
            
            # 添加直线
            line = self.doc.ModelSpace.AddLine(start_array, end_array)
            
            # 如果指定了图层，设置图层
            if layer:
                # 确保图层存在
                self.create_layer(layer)
                # 设置实体的图层
                line.Layer = layer
         
            # 如果指定了颜色，设置颜色
            if color is not None:
                line.Color = color

            if lineweight is not None:
                line.LineWeight = self.validate_lineweight(lineweight)
            
            # 刷新视图
            self.refresh_view()

            logger.debug(f"已绘制直线: 起点{start_point}, 终点{end_point}, 图层{layer if layer else '默认'}, 颜色{color if color is not None else '默认'}")
            return line
            
        except Exception as e:
            logger.error(f"绘制直线时出错: {str(e)}")
            return None
    
    def draw_circle(self, center: Tuple[float, float, float], 
                   radius: float, layer: str = None, color: int = None, lineweight=None) -> Any:
        """绘制圆"""
        if not self.is_running():
            return None
            
        try:
            # 确保点是三维的
            if len(center) == 2:
                center = (center[0], center[1], 0)
            
            # 使用VARIANT包装坐标点数据
            center_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [center[0], center[1], center[2]])
            
            # 添加圆
            circle = self.doc.ModelSpace.AddCircle(center_array, radius)
            
            # 如果指定了图层，设置图层
            if layer:
                # 确保图层存在
                self.create_layer(layer)
                # 设置实体的图层
                circle.Layer = layer
            
            # 如果指定了颜色，设置颜色
            if color is not None:
                circle.Color = color

            if lineweight is not None:
                circle.LineWeight = self.validate_lineweight(lineweight)
            
            # 刷新视图
            self.refresh_view()
            
            logger.debug(f"已绘制圆: 中心{center}, 半径{radius}, 图层{layer if layer else '默认'}, 颜色{color if color is not None else '默认'}")
            return circle
            
        except Exception as e:
            logger.error(f"绘制圆时出错: {str(e)}")
            return None
    
    def draw_arc(self, center: Tuple[float, float, float], 
                radius: float, start_angle: float, end_angle: float, layer: str = None, color: int = None, lineweight=None) -> Any:
        """绘制圆弧"""
        if not self.is_running():
            return None
            
        try:
            # 确保点是三维的
            if len(center) == 2:
                center = (center[0], center[1], 0)
                
            # 将角度转换为弧度
            start_rad = math.radians(start_angle)
            end_rad = math.radians(end_angle)
            
            # 使用VARIANT包装坐标点数据
            center_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [center[0], center[1], center[2]])
            
            # 添加圆弧
            arc = self.doc.ModelSpace.AddArc(center_array, radius, start_rad, end_rad)
            
            # 如果指定了图层，设置图层
            if layer:
                # 确保图层存在
                self.create_layer(layer)
                # 设置实体的图层
                arc.Layer = layer
            
            # 如果指定了颜色，设置颜色
            if color is not None:
                arc.Color = color

            if lineweight is not None:
                arc.LineWeight = self.validate_lineweight(lineweight)
            
            # 刷新视图
            self.refresh_view()
            
            logger.debug(f"已绘制圆弧: 中心{center}, 半径{radius}, 起始角度{start_angle}, 结束角度{end_angle}, 图层{layer if layer else '默认'}, 颜色{color if color is not None else '默认'}")
            return arc
        except Exception as e:
            logger.error(f"绘制圆弧失败: {str(e)}")
            return None
    
    def draw_polyline(self, points: List[Tuple[float, float, float]], closed: bool = False, layer: str = None, color: int = None, lineweight=None) -> Any:
        """绘制多段线"""
        if not self.is_running():
            return None
            
        try:
            # 确保所有点都是三维的
            processed_points = []
            for point in points:
                if len(point) == 2:
                    processed_points.append((point[0], point[1], 0))
                else:
                    processed_points.append(point)
            
            # 创建点数组
            point_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                [coord for point in processed_points for coord in point])
            
            # 添加多段线
            polyline = self.doc.ModelSpace.AddPolyline(point_array)
            
            # 如果需要闭合
            if closed and len(processed_points) > 2:
                polyline.Closed = True
            
            # 如果指定了图层，设置图层
            if layer:
                # 确保图层存在
                self.create_layer(layer)
                # 设置实体的图层
                polyline.Layer = layer
            
            # 如果指定了颜色，设置颜色
            if color is not None:
                polyline.Color = color

            if lineweight is not None:
                polyline.LineWeight = self.validate_lineweight(lineweight)
            
            # 刷新视图
            self.refresh_view()

            logger.debug(f"已绘制多段线: {len(points)}个点, {'闭合' if closed else '不闭合'}, 图层{layer if layer else '默认'}, 颜色{color if color is not None else '默认'}")
            return polyline
        except Exception as e:
            logger.error(f"绘制多段线时出错: {str(e)}")
            return None
    
    def draw_rectangle(self, corner1: Tuple[float, float, float], 
                      corner2: Tuple[float, float, float], layer: str = None, color: int = None, lineweight=None) -> Any:
        """绘制矩形"""
        if not self.is_running():
            return None
            
        try:
            # 确保点是三维的
            if len(corner1) == 2:
                corner1 = (corner1[0], corner1[1], 0)
            if len(corner2) == 2:
                corner2 = (corner2[0], corner2[1], 0)
                
            # 计算矩形的四个角点
            x1, y1, z1 = corner1
            x2, y2, z2 = corner2
            
            # 创建矩形的四个点
            points = [
                (x1, y1, z1),
                (x2, y1, z1),
                (x2, y2, z1),
                (x1, y2, z1),
                (x1, y1, z1)  # 闭合矩形
            ]
            
            # 使用多段线绘制矩形
            return self.draw_polyline(points, True, layer, color, lineweight)
        except Exception as e:
            logger.error(f"绘制矩形时出错: {str(e)}")
            return None
    
    def draw_text(self, position: Tuple[float, float, float], 
                 text: str, height: float = 2.5, rotation: float = 0, layer: str = None, color: int = None) -> Any:
        """添加文本"""
        if not self.is_running():
            return None
            
        try:
            # 确保点是三维的
            if len(position) == 2:
                position = (position[0], position[1], 0)
            
            # 使用VARIANT包装坐标点数据
            position_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                 [position[0], position[1], position[2]])
                
            # 添加文本
            text_obj = self.doc.ModelSpace.AddText(text, position_array, height)
            
            # 设置旋转角度
            if rotation != 0:
                text_obj.Rotation = math.radians(rotation)
            
            # 如果指定了图层，设置图层
            if layer:
                # 确保图层存在
                self.create_layer(layer)
                # 设置实体的图层
                text_obj.Layer = layer
            
            # 如果指定了颜色，设置颜色
            if color is not None:
                text_obj.Color = color
            
            # 刷新视图
            self.refresh_view()
                                
            logger.debug(f"已添加文本: '{text}', 位置{position}, 高度{height}, 旋转{rotation}度, 图层{layer if layer else '默认'}, 颜色{color if color is not None else '默认'}")
            return text_obj
        except Exception as e:
            logger.error(f"添加文本时出错: {str(e)}")
            return None
    
    def draw_hatch(self, points: List[Tuple[float, float, float]], 
                  pattern_name: str = "SOLID", scale: float = 1.0, layer: str = None, color: int = None) -> Any:
        """绘制填充图案
        
        Args:
            points: 填充边界的点集，每个点为二维或三维坐标元组
            pattern_name: 填充图案名称，默认为"SOLID"(实体填充)
            scale: 填充图案比例，默认为1.0
            layer: 图层名称，如果为None则使用当前图层
            color: 颜色索引，如果为None则使用默认颜色
            
        Returns:
            成功返回填充对象，失败返回None
        """
        if not self.is_running():
            return None
            
        try:
            # 确保所有点都是有效的
            if not points or len(points) < 3:
                logger.error("创建填充失败: 至少需要3个点来定义填充边界")
                return None
                
            # 创建闭合多段线作为边界
            closed_polyline = self.draw_polyline(points, closed=True, layer=layer)
            if not closed_polyline:
                logger.error("创建填充失败: 无法创建边界多段线")
                return None
                
            # 创建填充对象 (0表示正常填充，True表示关联边界)
            hatch = self.doc.ModelSpace.AddHatch(0, pattern_name, True)
                
            # 添加外部边界循环
            # 使用VARIANT包装对象数组
            object_ids = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [closed_polyline])
            hatch.AppendOuterLoop(object_ids)
                
            # 设置填充图案比例
            hatch.PatternScale = scale
                
            # 如果指定了图层，设置图层
            if layer:
                # 确保图层存在
                self.create_layer(layer)
                # 设置实体的图层
                hatch.Layer = layer
                
            # 如果指定了颜色，设置颜色
            if color is not None:
                hatch.Color = color
                                
            # 更新填充 (计算填充区域)
            hatch.Evaluate()
            
            # 刷新视图
            self.refresh_view()
                
            logger.debug(f"已创建填充: 图案 {pattern_name}, 比例 {scale}, 图层{layer if layer else '默认'}, 颜色{color if color is not None else '默认'}")
            return hatch
        except Exception as e:
            logger.error(f"创建填充时出错: {str(e)}")
            return None
    
    # def erase_entity(self, entity_id: str) -> bool:
    #     """删除指定的实体"""
    #     if not self.is_running():
    #         return False
            
    #     try:
    #         # 在实体字典中查找
    #         for shape_name, entities in self.entities.items():
    #             for i, entity in enumerate(entities):
    #                 if hasattr(entity, 'Handle') and entity.Handle == entity_id:
    #                     entity.Delete()
    #                     entities.pop(i)
    #                     logger.info(f"已删除实体: {entity_id}")
                        
    #                     # 刷新视图
    #                     self.refresh_view()

    #                     return True
            
    #         # 如果在字典中找不到，尝试在当前文档中查找
    #         for obj in self.doc.ModelSpace:
    #             if hasattr(obj, 'Handle') and obj.Handle == entity_id:
    #                 obj.Delete()
    #                 logger.info(f"已删除实体: {entity_id}")
    #                 return True
                                
    #         logger.warning(f"未找到实体: {entity_id}")
    #         return False
    #     except Exception as e:
    #         logger.error(f"删除实体时出错: {str(e)}")
    #         return False
    
    # def move_entity(self, entity_id: str, 
    #                from_point: Tuple[float, float, float], 
    #                to_point: Tuple[float, float, float]) -> bool:
    #     """移动指定的实体"""
    #     if not self.is_running():
    #         return False
            
    #     try:
    #         # 确保点是三维的
    #         if len(from_point) == 2:
    #             from_point = (from_point[0], from_point[1], 0)
    #         if len(to_point) == 2:
    #             to_point = (to_point[0], to_point[1], 0)
                
    #         # 计算移动向量
    #         displacement = (
    #             to_point[0] - from_point[0],
    #             to_point[1] - from_point[1],
    #             to_point[2] - from_point[2]
    #         )
            
    #         # 在实体字典中查找
    #         for shape_name, entities in self.entities.items():
    #             for entity in entities:
    #                 if hasattr(entity, 'Handle') and entity.Handle == entity_id:
    #                     entity.Move(from_point, to_point)
    #                     logger.info(f"已移动实体: {entity_id}")
                        
    #                     # 刷新视图
    #                     self.refresh_view()
            
    #                     return True
            
    #         # 如果在字典中找不到，尝试在当前文档中查找
    #         for obj in self.doc.ModelSpace:
    #             if hasattr(obj, 'Handle') and obj.Handle == entity_id:
    #                 obj.Move(from_point, to_point)
    #                 logger.info(f"已移动实体: {entity_id}")
    #                 return True
                    
    #         logger.warning(f"未找到实体: {entity_id}")
    #         return False
    #     except Exception as e:
    #         logger.error(f"移动实体时出错: {str(e)}")
    #         return False
    
    # def rotate_entity(self, entity_id: str, 
    #                  base_point: Tuple[float, float, float], 
    #                  rotation_angle: float) -> bool:
    #     """旋转指定的实体"""
    #     if not self.is_running():
    #         return False
            
    #     try:
    #         # 确保点是三维的
    #         if len(base_point) == 2:
    #             base_point = (base_point[0], base_point[1], 0)
                
    #         # 将角度转换为弧度
    #         angle_rad = math.radians(rotation_angle)
            
    #         # 在实体字典中查找
    #         for shape_name, entities in self.entities.items():
    #             for entity in entities:
    #                 if hasattr(entity, 'Handle') and entity.Handle == entity_id:
    #                     entity.Rotate(base_point, angle_rad)
    #                     logger.info(f"已旋转实体: {entity_id}, 角度: {rotation_angle}度")

    #                     # 刷新视图
    #                     self.refresh_view()

    #                     return True
            
    #         # 如果在字典中找不到，尝试在当前文档中查找
    #         for obj in self.doc.ModelSpace:
    #             if hasattr(obj, 'Handle') and obj.Handle == entity_id:
    #                 obj.Rotate(base_point, angle_rad)
    #                 logger.info(f"已旋转实体: {entity_id}, 角度: {rotation_angle}度")
    #                 return True
                    
    #         logger.warning(f"未找到实体: {entity_id}")
    #         return False
    #     except Exception as e:
    #         logger.error(f"旋转实体时出错: {str(e)}")
    #         return False
    
    # def scale_entity(self, entity_id: str, 
    #                 base_point: Tuple[float, float, float], 
    #                 scale_factor: float) -> bool:
    #     """缩放指定的实体"""
    #     if not self.is_running():
    #         return False
            
    #     try:
    #         # 确保点是三维的
    #         if len(base_point) == 2:
    #             base_point = (base_point[0], base_point[1], 0)
                
    #         # 在实体字典中查找
    #         for shape_name, entities in self.entities.items():
    #             for entity in entities:
    #                 if hasattr(entity, 'Handle') and entity.Handle == entity_id:
    #                     entity.ScaleEntity(base_point, scale_factor)
    #                     logger.info(f"已缩放实体: {entity_id}, 比例: {scale_factor}")

    #                     # 刷新视图
    #                     self.refresh_view()

    #                     return True
            
    #         # 如果在字典中找不到，尝试在当前文档中查找
    #         for obj in self.doc.ModelSpace:
    #             if hasattr(obj, 'Handle') and obj.Handle == entity_id:
    #                 obj.ScaleEntity(base_point, scale_factor)
    #                 logger.info(f"已缩放实体: {entity_id}, 比例: {scale_factor}")
    #                 return True
                    
    #         logger.warning(f"未找到实体: {entity_id}")
    #         return False
    #     except Exception as e:
    #         logger.error(f"缩放实体时出错: {str(e)}")
    #         return False
    
    def zoom_extents(self) -> bool:
        """缩放视图以显示所有对象"""
        if not self.is_running():
            return False
            
        try:
            self.doc.ActiveViewport.ZoomExtents()
            logger.info("已缩放视图以显示所有对象")
            return True
        except Exception as e:
            logger.error(f"缩放视图时出错: {str(e)}")
            return False
    
    def close(self) -> None:
        """关闭CAD控制器"""
        try:
            # 释放COM资源
            if self.app is not None:
                del self.app
            pythoncom.CoUninitialize()
        except:
            pass

    
    def create_layer(self, layer_name: str) -> bool:    # , color: Union[int, Tuple[int, int, int]] = 7
        """创建新图层
        
        Args:
            layer_name: 图层名称
            color: 颜色值，可以是CAD颜色索引(int)或RGB颜色值(tuple)
            
        Returns:
            操作是否成功
        """
        if not self.is_running():
            return False
        
        try:
            # 检查图层是否已存在
            for i in range(self.doc.Layers.Count):
                if self.doc.Layers.Item(i).Name == layer_name:
                    # 图层已存在，激活它
                    self.doc.ActiveLayer = self.doc.Layers.Item(i)
                    return True
                
            # 创建新图层
            new_layer = self.doc.Layers.Add(layer_name)
            
            # 图层不设置颜色，设置里面的实体颜色
            # # 设置颜色
            # if isinstance(color, int):
            #     # 使用颜色索引
            #     new_layer.Color = color
            # elif isinstance(color, tuple) and len(color) == 3:
            #     # 使用RGB值
            #     r, g, b = color
            #     # 设置TrueColor
            #     new_layer.TrueColor = self._create_true_color(r, g, b)
            
            # 设置为当前图层
            self.doc.ActiveLayer = new_layer
            logger.info(f"已创建新图层: {layer_name}")  #, 颜色: {color}
            return True
        except Exception as e:
            logger.error(f"创建图层时出错: {str(e)}")
            return False

    def create_block(self, block_name: str, insertion_point: Tuple[float, float, float],
                    entities: List[Any] = None) -> Any:
        """创建块定义"""
        if not self.is_running():
            return None
        
        try:
            # 确保点是三维的
            if len(insertion_point) == 2:
                insertion_point = (insertion_point[0], insertion_point[1], 0)
            
            # 开始块定义
            block = self.doc.Blocks.Add(insertion_point, block_name)
            
            # 如果提供了实体列表，复制到块中
            if entities and len(entities) > 0:
                for entity in entities:
                    entity_copy = entity.Copy()
                    block.AddEntity(entity_copy)
            
            logger.info(f"已创建块: {block_name}")
            return block
        except Exception as e:
            logger.error(f"创建块时出错: {str(e)}")
            return None

    def insert_block(self, block_name: str, insertion_point: Tuple[float, float, float],
                   x_scale: float = 1.0, y_scale: float = 1.0, rotation: float = 0.0, layer: str = None) -> Any:
            """插入块引用"""
            if not self.is_running():
                return None
            
            try:
                # 确保点是三维的
                if len(insertion_point) == 2:
                    insertion_point = (insertion_point[0], insertion_point[1], 0)
                
                # 转换角度为弧度
                rotation_rad = math.radians(rotation)
                
                # 使用VARIANT包装坐标点数据
                insertion_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                     [insertion_point[0], insertion_point[1], insertion_point[2]])
                
                # 插入块引用
                block_ref = self.doc.ModelSpace.InsertBlock(insertion_array, block_name, 
                                                          x_scale, y_scale, 1.0, rotation_rad)
                
                # 如果指定了图层，设置图层
                if layer:
                    # 确保图层存在
                    self.create_layer(layer)
                    # 设置实体的图层
                    block_ref.Layer = layer
                
                logger.info(f"已插入块: {block_name} 在 {insertion_point}, 图层{layer if layer else '默认'}")
                return block_ref
            except Exception as e:
                logger.error(f"插入块时出错: {str(e)}")
                return None

    def add_dimension(self, start_point: Tuple[float, float, float], 
                     end_point: Tuple[float, float, float],
                     text_position: Tuple[float, float, float] = None, textheight: float = 5,layer: str = None, color: int=None) -> Any:
            """添加线性标注"""
            if not self.is_running():
                return None
            
            try:
                # 确保点是三维的
                if len(start_point) == 2:
                    start_point = (start_point[0], start_point[1], 0)
                if len(end_point) == 2:
                    end_point = (end_point[0], end_point[1], 0)
                
                # 如果未提供文本位置，自动计算
                if text_position is None:
                    # 在起点和终点之间的中点上方
                    mid_x = (start_point[0] + end_point[0]) / 2
                    mid_y = (start_point[1] + end_point[1]) / 2
                    text_position = (mid_x, mid_y + 5, 0)
                elif len(text_position) == 2:
                    text_position = (text_position[0], text_position[1], 0)
                
                # 使用VARIANT包装坐标点数据
                start_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                 [start_point[0], start_point[1], start_point[2]])
                end_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [end_point[0], end_point[1], end_point[2]])
                text_pos_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                     [text_position[0], text_position[1], text_position[2]])
                
                # 添加对齐标注
                dimension = self.doc.ModelSpace.AddDimAligned(start_array, end_array, text_pos_array)
                
                # 设置文字高度
                if textheight is not None:
                    dimension.TextHeight = textheight
                
                # 如果指定了图层，设置图层
                if layer:
                    # 确保图层存在
                    self.create_layer(layer)
                    # 设置实体的图层
                    dimension.Layer = layer

                # 如果指定了颜色，设置颜色
                if color is not None:
                    dimension.Color = color
                
                # 刷新视图
                self.refresh_view()

                logger.info(f"已添加标注: 从 {start_point} 到 {end_point}, 图层{layer if layer else '默认'}")
                return dimension
            except Exception as e:
                logger.error(f"添加标注时出错: {str(e)}")
                return None

    
    # def draw_wall(self, start_point: Tuple[float, float, float], 
    #              end_point: Tuple[float, float, float], width: float = 10.0, layer: str = None, color: int = None) -> List[Any]:
    #         """绘制墙体（用于建筑图）"""
    #         if not self.is_running():
    #             return None
            
    #         try:
    #             # 确保点是三维的
    #             if len(start_point) == 2:
    #                 start_point = (start_point[0], start_point[1], 0)
    #             if len(end_point) == 2:
    #                 end_point = (end_point[0], end_point[1], 0)
                
    #             # 计算墙体的方向向量
    #             dx = end_point[0] - start_point[0]
    #             dy = end_point[1] - start_point[1]
    #             length = math.sqrt(dx*dx + dy*dy)
                
    #             # 单位法向量
    #             if length > 0:
    #                 nx = -dy / length
    #                 ny = dx / length
    #             else:
    #                 return None
                
    #             # 计算墙体四个角点
    #             half_width = width / 2
    #             p1 = (start_point[0] + nx * half_width, start_point[1] + ny * half_width, start_point[2])
    #             p2 = (start_point[0] - nx * half_width, start_point[1] - ny * half_width, start_point[2])
    #             p3 = (end_point[0] - nx * half_width, end_point[1] - ny * half_width, end_point[2])
    #             p4 = (end_point[0] + nx * half_width, end_point[1] + ny * half_width, end_point[2])
                
    #             # 绘制封闭的多段线表示墙体
    #             points = [p1, p2, p3, p4, p1]  # 闭合多段线
    #             wall = self.draw_polyline(points, closed=True, layer=layer, color=color)
    #             return wall
    #         except Exception as e:
    #             logger.error(f"绘制墙体时出错: {str(e)}")
    #             return None

    # def draw_electrical_symbol(self, symbol_type: str, insertion_point: Tuple[float, float, float],
    #                          scale: float = 1.0, rotation: float = 0.0) -> Any:
    #         """绘制电气符号"""
    #         if not self.is_running():
    #             return None
            
    #         try:
    #             # 确保点是三维的
    #             if len(insertion_point) == 2:
    #                 insertion_point = (insertion_point[0], insertion_point[1], 0)
                
    #             # 根据符号类型创建不同的图形
    #             if symbol_type.lower() == "outlet" or symbol_type.lower() == "插座":
    #                 # 创建插座符号（圆圈）
    #                 circle = self.draw_circle(insertion_point, 2 * scale)
                    
    #                 # 在圆内添加水平线
    #                 p1 = (insertion_point[0] - 1.5 * scale, insertion_point[1], insertion_point[2])
    #                 p2 = (insertion_point[0] + 1.5 * scale, insertion_point[1], insertion_point[2])
    #                 line = self.draw_line(p1, p2)
                    
    #                 # 如果需要旋转
    #                 if rotation != 0:
    #                     self.rotate_entity(circle.Handle, insertion_point, rotation)
    #                     self.rotate_entity(line.Handle, insertion_point, rotation)
                        
    #                 return circle  # 返回主要实体
                    
    #             elif symbol_type.lower() == "switch" or symbol_type.lower() == "开关":
    #                 # 创建开关符号（正方形）
    #                 size = 3 * scale
    #                 p1 = (insertion_point[0] - size/2, insertion_point[1] - size/2, insertion_point[2])
    #                 p2 = (insertion_point[0] + size/2, insertion_point[1] + size/2, insertion_point[2])
    #                 square = self.draw_rectangle(p1, p2)
                    
    #                 # 如果需要旋转
    #                 if rotation != 0:
    #                     self.rotate_entity(square.Handle, insertion_point, rotation)
                        
    #                 return square
                    
    #             elif symbol_type.lower() == "light" or symbol_type.lower() == "灯":
    #                 # 创建灯符号（带叉的圆）
    #                 circle = self.draw_circle(insertion_point, 3 * scale)
                    
    #                 # 添加交叉线
    #                 p1 = (insertion_point[0] - 2 * scale, insertion_point[1] - 2 * scale, insertion_point[2])
    #                 p2 = (insertion_point[0] + 2 * scale, insertion_point[1] + 2 * scale, insertion_point[2])
    #                 line1 = self.draw_line(p1, p2)
                    
    #                 p3 = (insertion_point[0] - 2 * scale, insertion_point[1] + 2 * scale, insertion_point[2])
    #                 p4 = (insertion_point[0] + 2 * scale, insertion_point[1] - 2 * scale, insertion_point[2])
    #                 line2 = self.draw_line(p3, p4)
                    
    #                 # 如果需要旋转
    #                 if rotation != 0:
    #                     self.rotate_entity(circle.Handle, insertion_point, rotation)
    #                     self.rotate_entity(line1.Handle, insertion_point, rotation)
    #                     self.rotate_entity(line2.Handle, insertion_point, rotation)
                        
    #                 return circle
                    
    #             else:
    #                 logger.warning(f"未知的电气符号类型: {symbol_type}")
    #                 return None
    #         except Exception as e:
    #             logger.error(f"绘制电气符号时出错: {str(e)}")
    #             return None


# 以下是发送命令的方式，暂时还没用到
    def draw_circle_command(self, center: Tuple[float, float, float], 
                           radius: float) -> bool:
        """使用命令行方式绘制圆"""
        if not self.is_running():
            return False
        
        try:
            # 格式化坐标和半径
            x, y, z = center
            
            # 修改命令格式，增加更明确的坐标分隔
            cmd = f"_CIRCLE {x},{y},{z} {radius}\n"
            
            # 日志记录
            logger.info(f"发送命令: {cmd}")
            
            # 清除当前命令（防止干扰）
            self.doc.SendCommand("\x03\x03")  # 发送ESC键
            time.sleep(0.1)
            
            # 发送命令并等待执行
            self.doc.SendCommand(cmd)
            time.sleep(0.8)  # 增加等待时间
            
            # 刷新视图
            self.refresh_view()
            
            logger.info("成功执行绘制圆命令")
            return True
            
        except Exception as e:
            logger.error(f"执行绘制圆命令失败: {str(e)}")
            return False

    def end_command(self):
        """结束当前命令"""
        try:
            # 发送回车和ESC组合来确保命令结束
            self.doc.SendCommand("\n\x03")
            time.sleep(0.1)
        except:
            pass

    def draw_line_command(self, start_point: Tuple[float, float, float],
                         end_point: Tuple[float, float, float]) -> bool:
        """使用命令行方式绘制直线"""
        if not self.is_running():
            return False
        
        try:
            # 先结束任何可能正在进行的命令
            self.end_command()
            
            # 格式化坐标
            sx, sy, sz = start_point
            ex, ey, ez = end_point
            
            # 修改命令格式，使用空格分隔坐标
            cmd = f"_LINE {sx},{sy},{sz} {ex},{ey},{ez}\n\n"
            
            # 发送命令
            logger.info(f"发送命令: {cmd}")
            self.doc.SendCommand(cmd)
            
            # 等待命令执行完成
            time.sleep(self.command_delay)
            
            # 确保命令结束
            self.end_command()
            
            # 刷新视图
            self.refresh_view()
            
            logger.info("成功执行绘制直线命令")
            return True
            
        except Exception as e:
            logger.error(f"执行绘制直线命令失败: {str(e)}")
            return False

    def save_drawing_command(self, file_path: str) -> bool:
        """使用命令行方式保存图纸"""
        if not self.is_running():
            return False
        
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            # 方法1: 使用QSAVE命令 - 需要用户交互
            logger.info("使用QSAVE命令保存图纸")
            self.doc.SendCommand("_QSAVE\n")
            
            logger.info("QSAVE命令已发送，请在CAD界面中选择保存位置")
            # 这里需要用户在CAD界面上操作
            
            return True
        except Exception as e:
            logger.error(f"保存图纸失败: {str(e)}")
            return False

    def save_drawing_script(self, file_path: str) -> bool:
        """使用脚本文件方式保存图纸"""
        if not self.is_running():
            return False
        
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            # 创建临时脚本文件
            script_path = os.path.join(os.path.dirname(file_path), "save_script.scr")
            with open(script_path, "w") as f:
                # 写入SAVEAS命令
                abs_path = os.path.abspath(file_path)
                f.write(f'SAVEAS\n"DWG"\n"{abs_path}"\n')
            
            # 运行脚本
            logger.info(f"通过脚本文件保存图纸到: {file_path}")
            self.doc.SendCommand(f'_SCRIPT\n"{script_path}"\n')
            
            # 等待脚本执行完成
            time.sleep(2)
            
            # 删除临时脚本文件
            try:
                os.remove(script_path)
            except:
                pass
            
            logger.info("图纸保存成功")
            return True
        except Exception as e:
            logger.error(f"保存图纸失败: {str(e)}")
            return False