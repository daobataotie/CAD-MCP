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
                 end_point: Tuple[float, float, float], layer: str = None, color: int = None, lineweight=None) -> bool:
        """绘制直线"""
        if not self.is_running():
            return False
            
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
 
    def draw_ellipse(self, center: Tuple[float, float, float], 
                    major_axis: float, minor_axis: float, rotation: float = 0, 
                    layer: str = None, color: int = None, lineweight=None) -> Any:
        """绘制椭圆"""
        if not self.is_running():
            return None
            
        try:
            # 确保点是三维的
            if len(center) == 2:
                center = (center[0], center[1], 0)
            
            if rotation is None:
                rotation = 0

            # 将旋转角度转换为弧度
            rotation_rad = math.radians(rotation)
            
            # 使用VARIANT包装坐标点数据
            center_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [center[0], center[1], center[2]])
            
            # 计算椭圆的主轴向量
            major_x = major_axis * math.cos(rotation_rad)
            major_y = major_axis * math.sin(rotation_rad)
            major_vector = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [major_x, major_y, 0])
            
            # 添加椭圆
            ellipse = self.doc.ModelSpace.AddEllipse(center_array, major_vector, minor_axis / major_axis)
            
            # 如果指定了图层，设置图层
            if layer:
                # 确保图层存在
                self.create_layer(layer)
                # 设置实体的图层
                ellipse.Layer = layer
            
            # 如果指定了颜色，设置颜色
            if color is not None:
                ellipse.Color = color

            if lineweight is not None:
                ellipse.LineWeight = self.validate_lineweight(lineweight)
            
            # 刷新视图
            self.refresh_view()
            
            logger.debug(f"已绘制椭圆: 中心{center}, 长轴{major_axis}, 短轴{minor_axis}, 旋转角度{rotation}, 图层{layer if layer else '默认'}, 颜色{color if color is not None else '默认'}")
            return ellipse
        except Exception as e:
            logger.error(f"绘制椭圆失败: {str(e)}")
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

