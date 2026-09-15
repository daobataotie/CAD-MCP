import logging
import math
import time
import os
import io
import json
import threading
import queue
import concurrent.futures
from contextlib import contextmanager
from typing import Any, Dict, List, Optional, Tuple, Union

# 统一使用配置单例（替代此前各模块重复读取 config.json）
try:
    from .config import get_config
except ImportError:
    from config import get_config

config = get_config()

try:
    import win32com.client
    # pythoncom是pywin32的一部分，不需要单独安装
    import pythoncom
except ImportError:
    logging.error("无法导入win32com.client或pythoncom，请确保已安装pywin32库")
    raise

try:
    import anyio
except ImportError:
    # anyio 是 mcp SDK 的依赖，正常情况下随 mcp 一起安装
    anyio = None

logger = logging.getLogger('cad_controller')

# COM 瞬态错误：调用被忙碌的CAD拒绝（RPC_E_CALL_REJECTED）——
# CAD正在执行命令/弹窗/重生成时会短暂拒绝COM调用，稍后重试即可恢复
_RPC_E_CALL_REJECTED = -2147418111

# AutoCAD 瞬态保存错误（"保存文档时出错"）——
# 保存/打开/关闭等操作进行中再调用SaveAs会报此错，稍后重试即可恢复
_ACAD_SAVE_TRANSIENT = -2145320861

# COM 死引用错误：CAD进程已被用户关闭/崩溃，持有的接口指针失效
# RPC_E_DISCONNECTED(-2147417848)：对象已与客户端断开连接
# RPC_S_SERVER_UNAVAILABLE(-2147023174)：RPC服务器不可用
# 与瞬态错误不同，这类错误重试无意义，应重置连接状态让后续操作重连
_DEAD_COM_HRESULTS = (-2147417848, -2147023174)


def _is_retryable_com_error(e) -> bool:
    """判断是否为可重试的COM瞬态错误（忙碌被拒或保存冲突）"""
    if getattr(e, "hresult", None) == _RPC_E_CALL_REJECTED:
        return True
    # 瞬态保存错误藏在com_error的excepinfo元组里
    info = tuple(getattr(e, "excepinfo", None) or ())
    return _ACAD_SAVE_TRANSIENT in info


def _is_retryable_attr_error(e, retry_attr) -> bool:
    """判断是否为pywin32动态派发把"成员查找被拒"转换成的AttributeError

    动态派发下成员查找(GetIDsOfNames)被忙碌的CAD拒绝时，pywin32会吞掉
    com_error并抛出AttributeError，hresult随之丢失（曾导致打开文档后立即
    关闭报"Open.Close"且调用级重试完全失效）。已知两种消息格式：
    dynamic.py抛出"Open.Close"；client/__init__.py抛出
    "'<COMObject Open>' object has no attribute 'Close'"。
    未指定成员名时一律不重试。
    """
    if not retry_attr or not isinstance(e, AttributeError):
        return False
    msg = str(e)
    return msg.endswith("." + retry_attr) or msg.endswith(f"attribute '{retry_attr}'")


def _call_with_com_retry(fn, args, kwargs, retries: int = 20, delay: float = 0.5, retry_attr: str = None):
    """执行COM调用，忙碌(RPC_E_CALL_REJECTED)或瞬态保存冲突时自动重试

    总等待时间上限约 retries*delay 秒，超过后抛出原始异常。
    retry_attr 见 _is_retryable_attr_error 的说明。
    """
    for attempt in range(retries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            # 仅对瞬态错误重试（COM忙碌/保存冲突/成员查找被拒），其他异常直接抛出
            if (_is_retryable_com_error(e) or _is_retryable_attr_error(e, retry_attr)) \
                    and attempt < retries - 1:
                logger.debug(f"CAD忙碌或保存冲突，COM调用被拒，第{attempt + 1}次重试...")
                time.sleep(delay)
                continue
            raise
    raise RuntimeError("unreachable")  # 逻辑上不可达，仅为类型完整


class COMExecutor:
    """专用COM线程执行器

    Windows COM 的套间(Apartment)模型要求：同一接口指针应在创建它的线程上使用，
    跨线程直接访问会导致调用失败甚至进程崩溃。本执行器将所有 COM 调用串行化到
    单一 STA 线程上执行，带来两个好处：
    1. 线程安全——多个工具并发调用也不会出现跨套间访问；
    2. 不阻塞 asyncio 事件循环——COM 的阻塞调用（如等待 CAD 启动20秒）
       全部发生在专用线程上，服务器仍能响应其他请求。
    """

    def __init__(self):
        self._queue: "queue.Queue" = queue.Queue()
        self._thread = threading.Thread(target=self._worker, name="cad-com", daemon=True)
        self._thread.start()

    def _worker(self) -> None:
        """COM 线程主循环：初始化 COM 后不断从队列取任务执行"""
        # 整个进程生命周期内只初始化一次，不调用 CoUninitialize（保持套间存活）
        pythoncom.CoInitialize()
        while True:
            fn, args, kwargs, future = self._queue.get()
            if fn is None:
                # 停止信号
                future.set_result(True)
                break
            try:
                future.set_result(_call_with_com_retry(fn, args, kwargs))
            except BaseException as e:
                future.set_exception(e)
            finally:
                self._queue.task_done()

    def submit(self, fn, *args, **kwargs):
        """在COM线程上同步执行 fn 并返回其结果（阻塞当前调用线程）"""
        if threading.current_thread() is self._thread:
            # 已在COM线程上（内部嵌套调用），直接执行避免死锁
            return fn(*args, **kwargs)
        future = concurrent.futures.Future()
        self._queue.put((fn, args, kwargs, future))
        return future.result()

    async def run(self, fn, *args, **kwargs):
        """异步执行 COM 调用：把阻塞的 submit 放到线程池，避免阻塞事件循环"""
        if anyio is None:
            # 无 anyio 时退化为同步执行（不应发生）
            return self.submit(fn, *args, **kwargs)
        return await anyio.to_thread.run_sync(lambda: self.submit(fn, *args, **kwargs))


class CADController:
    """CAD控制器类，负责与CAD应用程序交互"""

    def __init__(self):
        """初始化CAD控制器"""
        self.app = None
        self.doc = None
        self.entities = {}  # 存储已创建图形的实体引用，用于后续修改
        # 专用COM线程执行器：所有COM调用经由它串行执行
        self.executor = COMExecutor()
        # 从配置文件加载参数
        self.startup_wait_time = config["cad"]["startup_wait_time"]
        self.command_delay = config["cad"]["command_delay"]
        # 获取CAD类型
        self.cad_type = config["cad"]["type"]
        # 有效的线宽值列表
        self.valid_lineweights = [0, 5, 9, 13, 15, 18, 20, 25, 30, 35, 40, 50, 53, 60, 70, 80, 90, 100, 106, 120, 140, 158, 200, 211]
        logger.info("CAD控制器已初始化")

    def _get_app_info(self) -> Tuple[List[str], str]:
        """根据配置的CAD类型返回 (候选COM ProgID列表, 显示名称)

        注意：COM 的 ProgID 查找不区分大小写，GCAD.Application 与 Gcad.Application 等价。
        同一CAD不同年代版本注册的ProgID可能不同（如浩辰有GCAD/GstarCAD两种），
        返回候选列表由 start_cad 依次尝试，避免目标机器只注册了另一种时连不上。
        """
        # 各CAD软件的COM ProgID映射（未匹配的类型回退到AutoCAD）
        app_map = {
            "autocad": (["AutoCAD.Application"], "AutoCAD"),
            "gcad": (["GCAD.Application", "GstarCAD.Application"], "浩辰CAD"),
            "gstarcad": (["GCAD.Application", "GstarCAD.Application"], "浩辰CAD"),
            "zwcad": (["ZWCAD.Application"], "中望CAD"),
        }
        return app_map.get(self.cad_type.lower(), (["AutoCAD.Application"], "AutoCAD"))

    def start_cad(self) -> bool:
        """启动CAD并创建或打开一个文档"""
        # 存储旧实例引用（如果有）以便后续清理
        old_app = None
        if self.app is not None:
            old_app = self.app
            self.app = None
            self.doc = None

        try:
            # 初始化COM（若已在COM线程上则幂等）
            pythoncom.CoInitialize()

            app_ids, app_name = self._get_app_info()

            # 第一步：尝试连接已运行的实例（遍历候选ProgID）
            for app_id in app_ids:
                try:
                    logger.info(f"尝试连接现有{app_name}实例（{app_id}）...")
                    self.app = win32com.client.GetActiveObject(app_id)
                    logger.info(f"成功连接到已运行的{app_name}实例")
                    break
                except Exception as e:
                    logger.info(f"未找到运行中的{app_name}实例（{app_id}），继续尝试: {str(e)}")

            # 第二步：连接失败则启动新实例（同样遍历候选ProgID）
            if self.app is None:
                for app_id in app_ids:
                    try:
                        logger.info(f"正在启动{app_name}实例（{app_id}）...")
                        self.app = win32com.client.Dispatch(app_id)
                        self._call(setattr, self.app, "Visible", True)
                        break
                    except Exception as e:
                        logger.info(f"启动{app_name}实例失败（{app_id}），继续尝试: {str(e)}")
                if self.app is None:
                    raise Exception(f"无法连接或启动{app_name}（候选ProgID: {app_ids}）")
                # 等待CAD启动
                time.sleep(self.startup_wait_time)  # 使用配置的等待时间

            # 第三步：获取或创建文档（集合与属性访问都可能被忙碌的CAD拒绝，逐调用重试）
            docs = self._call(getattr, self.app, "Documents")
            if self._call(getattr, docs, "Count") == 0:
                logger.info("创建新文档...")
                self.doc = self._call_attr(docs, "Add")
            else:
                try:
                    logger.info("获取活动文档...")
                    self.doc = self._call(getattr, self.app, "ActiveDocument")
                except Exception:
                    # CAD配置为"启动不新建文档"时没有活动文档，回退到手动创建。
                    # 注意：绝不能为"强制获取文档"去关闭用户已打开的图纸——
                    # Close(False)会直接丢弃未保存的修改（旧版兜底逻辑曾这样做，
                    # 属于数据丢失级缺陷），拿不到活动文档时新建一个才是安全做法
                    logger.info("无活动文档，改为新建文档...")
                    self.doc = self._call_attr(docs, "Add")

            # 额外安全检查和等待
            time.sleep(2)  # 给CAD更多时间处理文档创建

            if self.doc is None:
                raise Exception("无法获取有效的Document对象")

            # 尝试读取文档属性以验证其有效性
            try:
                name = self._call(getattr, self.doc, "Name")
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

    def _pt(self, point) -> "win32com.client.VARIANT":
        """将坐标点包装为COM可接受的VARIANT双精度数组

        自动将二维坐标补齐为三维（Z=0），供 AddLine/AddCircle 等方法使用。
        """
        if point is None or len(point) < 2:
            raise ValueError(f"坐标点至少需要两个分量: {point}")
        x, y = float(point[0]), float(point[1])
        z = float(point[2]) if len(point) > 2 else 0.0
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y, z])

    @staticmethod
    def _safe_attr(obj, name: str, default=None):
        """安全读取COM对象属性，实体类型不匹配时返回默认值"""
        try:
            return getattr(obj, name)
        except Exception:
            return default

    @classmethod
    def _read_color(cls, obj) -> Optional[int]:
        """读取对象颜色为ACI索引整数（在COM线程上调用）

        兼容多种COM包装差异：
        1. gen_py静态包装（EnsureDispatch/makepy缓存）对属性名大小写敏感，
           AutoCAD类型库中实体颜色属性为小写color；
        2. 动态派发下IDispatch名字不区分大小写，Color可访问；
        3. 值可能是整数（旧版Color/color属性）或AcadAcCmColor对象（TrueColor）。
        """
        for name in ("color", "Color", "ColorIndex"):
            raw = cls._safe_attr(obj, name)
            if raw is None:
                continue
            if isinstance(raw, int):
                return raw
            # 新版CAD返回AcadAcCmColor对象，取其ColorIndex
            idx = cls._safe_attr(raw, "ColorIndex")
            if idx is not None:
                return idx
        return None

    @staticmethod
    def _write_color(entity, color: int) -> None:
        """设置实体颜色为ACI索引（兼容gen_py小写color与动态派发Color两种包装）"""
        try:
            entity.color = color  # gen_py静态包装下的真实属性名
        except AttributeError:
            entity.Color = color  # 动态派发下IDispatch不区分大小写

    def _call(self, fn, *args, **kwargs):
        """带重试执行单个COM调用（仅在COM线程上调用）

        CAD忙碌(RPC_E_CALL_REJECTED)时调用被拒绝意味着该调用未执行，
        对单个COM调用重试是安全的；对整个业务方法重试则可能重复绘制
        （方法前半段已成功、后半段被拒的场合），因此重试粒度必须在调用级。
        """
        try:
            return _call_with_com_retry(fn, args, kwargs)
        except Exception as e:
            # CAD进程已被用户关闭时，持有的COM引用全部失效（重试无意义）：
            # 重置连接状态，让下一次操作走 is_running()==False 分支自动重连，
            # 否则所有工具会持续报"操作失败"且永远不会触发重启
            if getattr(e, "hresult", None) in _DEAD_COM_HRESULTS:
                logger.warning("检测到CAD进程已退出，重置连接状态（下一次操作将重新启动CAD）")
                self.app = None
                self.doc = None
            raise

    def _call_attr(self, obj, name: str, *args, **kwargs):
        """带重试执行 obj.name(*args)：成员查找与调用均在重试保护内

        动态派发（无gen_py缓存的机器，即最终用户的常态）下 obj.name 的
        查找本身也是一次COM调用（GetIDsOfNames），同样可能被忙碌的CAD
        拒绝，因此不能写成 self._call(obj.name, ...) ——查找发生在进入
        _call之前。本方法将查找与调用合并为一个重试单元，
        并把pywin32转换后的"查找被拒"AttributeError也纳入瞬态重试。
        """
        return _call_with_com_retry(lambda: getattr(obj, name)(*args, **kwargs),
                                    (), {}, retry_attr=name)

    def _ms(self):
        """获取模型空间（属性读取同样可能被忙碌的CAD拒绝，需带重试）"""
        return self._call(getattr, self.doc, "ModelSpace")

    @contextmanager
    def _undo_group(self):
        """将本次工具操作打包为独立撤销组（保证UNDO 1正好撤销一步）

        AutoCAD会把命令行空闲期间的所有COM修改合并为同一个撤销组，
        不加标记时一次UNDO可能把整批绘制全部撤销。
        嵌套使用安全（如draw_rectangle→draw_polyline），内层组归入外层组。
        """
        doc = self.doc
        # 标记调用同样可能被忙碌的CAD拒绝，需带重试
        self._call_attr(doc, "StartUndoMark")
        try:
            yield
        finally:
            try:
                self._call_attr(doc, "EndUndoMark")
            except Exception:
                # 文档已切换等极端情况下无法收尾标记，忽略（标记随文档重置）
                pass

    def _check_target_not_occupied(self, file_path: str) -> None:
        """检查目标保存路径是否被其他已打开的图纸占用

        AutoCAD会锁定已打开文档对应的DWG文件，向这些路径SaveAs会持续报
        "保存文档时出错"(-2145320861)且重试无效（实测等待40秒仍失败），
        故提前检测快速失败。当前文档占用同路径属于原地另存，放行。
        """
        try:
            norm = os.path.normcase(file_path)
            current_name = self._safe_attr(self.doc, "Name")
            docs = self._call(getattr, self.app, "Documents")
            for i in range(self._call(getattr, docs, "Count")):
                doc = self._call_attr(docs, "Item", i)
                full = self._safe_attr(doc, "FullName")
                if full and os.path.normcase(str(full)) == norm:
                    # 打开文档的名称在同一CAD会话内唯一，同名即当前文档
                    if self._safe_attr(doc, "Name") == current_name:
                        continue
                    raise ValueError(
                        f"目标文件已被其他打开的图纸占用，无法覆盖保存: {file_path}（请先关闭占用该文件的图纸）"
                    )
        except ValueError:
            raise
        except Exception:
            # 预检查本身失败（如CAD忙碌）不阻断保存，交由SaveAs及其调用级重试兜底
            pass

    def save_drawing(self, file_path: str) -> bool:
        """保存当前图纸到指定路径

        Raises:
            ValueError: 目标文件正被其他打开的图纸占用时
        """
        if not self.is_running():
            logger.error("CAD未运行，无法保存图纸")
            return False

        try:
            # 相对路径必须先转为绝对路径：SaveAs在CAD进程内解析路径，
            # 其工作目录与MCP服务器的工作目录通常不同（实测相对文件会存到别处）
            file_path = os.path.abspath(file_path)

            # 确保目录存在（abspath后必有dirname；目录可能尚未创建）
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            # 目标被占用时SaveAs会持续报错且重试无效，提前快速失败
            self._check_target_not_occupied(file_path)

            # 保存文件（SaveAs返回后CAD仍可能短暂拒绝后续修改调用，
            # 后续工具的COM调用会经由调用级重试自动等待恢复）
            self._call_attr(self.doc, "SaveAs", file_path)
            logger.info(f"图纸已保存到: {file_path}")

            return True
        except ValueError:
            raise
        except Exception as e:
            logger.error(f"保存图纸失败: {str(e)}")
            return False

    def refresh_view(self) -> None:
        """强制重生成视图（仅用于显式刷新场景，勿在绘图/编辑后调用）

        注意：AutoCAD会把COM修改后的重生成(Regen)记录进撤销栈，导致
        随后的UNDO变成不可见变化的空操作（实测无论是否加撤销标记）。
        因此各绘图/编辑方法不再自动刷新——新增实体会自动显示，
        需要全图刷新时用 zoom_extents 或 send_command 执行 ZOOM。
        """
        if self.is_running():
            try:
                self._call_attr(self.doc, "Regen", 1)  # acAllViewports = 1
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

    def _apply_style(self, entity, layer=None, color=None, lineweight=None) -> None:
        """为实体统一设置图层/颜色/线宽（各绘制方法的公共部分）

        属性设置是幂等操作，逐个带重试执行：单个属性被CAD拒绝时重试，
        不会影响已成功的部分。
        """
        # 如果指定了图层，设置图层
        if layer:
            # 确保图层存在（只创建不激活：避免把用户正在使用的当前图层悄悄切走）
            target_layer = self._ensure_layer(layer)
            # 设置实体的图层（使用CAD返回的真实图层名，避免传入名称大小写不匹配）
            self._call(setattr, entity, "Layer", self._call(getattr, target_layer, "Name"))

        # 如果指定了颜色，设置颜色
        if color is not None:
            self._call(self._write_color, entity, color)

        if lineweight is not None:
            lw = self.validate_lineweight(lineweight)
            # gen_py包装下属性名为Lineweight（与LineWeight的写法差异同Color/color）
            try:
                self._call(setattr, entity, "Lineweight", lw)
            except AttributeError:
                self._call(setattr, entity, "LineWeight", lw)

    def draw_line(self, start_point: Tuple[float, float, float],
                 end_point: Tuple[float, float, float], layer: str = None, color: int = None, lineweight=None) -> Any:
        """绘制直线"""
        if not self.is_running():
            return None

        try:
            # 使用VARIANT包装坐标点数据（_pt自动补齐Z轴）
            start_array = self._pt(start_point)
            end_array = self._pt(end_point)

            # 添加直线（独立撤销组，忙碌时在调用级自动重试）
            with self._undo_group():
                line = self._call_attr(self._ms(), "AddLine", start_array, end_array)

                # 设置图层、颜色、线宽
                self._apply_style(line, layer, color, lineweight)

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

        # 半径必须为正（CAD拒绝0和负半径），提前失败并给出明确日志
        if radius <= 0:
            logger.error(f"绘制圆失败: 半径({radius})必须为正数")
            return None

        try:
            # 使用VARIANT包装坐标点数据（_pt自动补齐Z轴）
            center_array = self._pt(center)

            # 添加圆（独立撤销组，忙碌时在调用级自动重试）
            with self._undo_group():
                circle = self._call_attr(self._ms(), "AddCircle", center_array, radius)

                # 设置图层、颜色、线宽
                self._apply_style(circle, layer, color, lineweight)

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

        # 半径必须为正（CAD拒绝0和负半径），提前失败并给出明确日志
        if radius <= 0:
            logger.error(f"绘制圆弧失败: 半径({radius})必须为正数")
            return None

        try:
            # 将角度转换为弧度
            start_rad = math.radians(start_angle)
            end_rad = math.radians(end_angle)

            # 使用VARIANT包装坐标点数据
            center_array = self._pt(center)

            # 添加圆弧（独立撤销组，忙碌时在调用级自动重试）
            with self._undo_group():
                arc = self._call_attr(self._ms(), "AddArc", center_array, radius, start_rad, end_rad)

                # 设置图层、颜色、线宽
                self._apply_style(arc, layer, color, lineweight)

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

        # 长轴/短轴必须为正，且长轴不小于短轴——AddEllipse的RadiusRatio参数
        # 为短轴/长轴，必须落在(0,1]区间；长轴为0还会导致本地除零
        if major_axis <= 0 or minor_axis <= 0 or minor_axis > major_axis:
            logger.error(f"绘制椭圆失败: 长轴({major_axis})必须为正且不小于短轴({minor_axis})")
            return None

        try:
            if rotation is None:
                rotation = 0

            # 将旋转角度转换为弧度
            rotation_rad = math.radians(rotation)

            # 使用VARIANT包装坐标点数据
            center_array = self._pt(center)

            # 计算椭圆的主轴向量
            major_x = major_axis * math.cos(rotation_rad)
            major_y = major_axis * math.sin(rotation_rad)
            major_vector = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                               [major_x, major_y, 0])

            # 添加椭圆（独立撤销组，忙碌时在调用级自动重试）
            with self._undo_group():
                ellipse = self._call_attr(self._ms(), "AddEllipse", center_array, major_vector, minor_axis / major_axis)

                # 设置图层、颜色、线宽
                self._apply_style(ellipse, layer, color, lineweight)

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

            # 添加多段线（独立撤销组，忙碌时在调用级自动重试）
            with self._undo_group():
                polyline = self._call_attr(self._ms(), "AddPolyline", point_array)

                # 如果需要闭合
                if closed and len(processed_points) > 2:
                    self._call(setattr, polyline, "Closed", True)

                # 设置图层、颜色、线宽
                self._apply_style(polyline, layer, color, lineweight)

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
            # 使用VARIANT包装坐标点数据（_pt自动补齐Z轴）
            position_array = self._pt(position)

            # 添加文本（独立撤销组，忙碌时在调用级自动重试）
            with self._undo_group():
                text_obj = self._call_attr(self._ms(), "AddText", text, position_array, height)

                # 设置旋转角度
                if rotation != 0:
                    self._call(setattr, text_obj, "Rotation", math.radians(rotation))

                # 设置图层、颜色
                self._apply_style(text_obj, layer, color)

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

        closed_polyline = None
        try:
            # 确保所有点都是有效的
            if not points or len(points) < 3:
                logger.error("创建填充失败: 至少需要3个点来定义填充边界")
                return None

            # 创建边界多段线（内部已有独立撤销组，此处外层组使填充+边界成为一个撤销步骤）
            with self._undo_group():
                closed_polyline = self.draw_polyline(points, closed=True, layer=layer)
                if not closed_polyline:
                    logger.error("创建填充失败: 无法创建边界多段线")
                    return None

                # 创建填充对象 (0表示正常填充，True表示关联边界)
                hatch = self._call_attr(self._ms(), "AddHatch", 0, pattern_name, True)

                # 添加外部边界循环
                # 使用VARIANT包装对象数组
                object_ids = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [closed_polyline])
                self._call_attr(hatch, "AppendOuterLoop", object_ids)

                # 设置填充图案比例
                self._call(setattr, hatch, "PatternScale", scale)

                # 设置图层、颜色
                self._apply_style(hatch, layer, color)

                # 更新填充 (计算填充区域)
                self._call_attr(hatch, "Evaluate")

            logger.debug(f"已创建填充: 图案 {pattern_name}, 比例 {scale}, 图层{layer if layer else '默认'}, 颜色{color if color is not None else '默认'}")
            return hatch
        except Exception as e:
            # 失败时清理已创建的边界多段线，避免在图纸上留下孤儿实体
            # （调用方看到的是"失败"，图纸上却不应多出一条闭合多段线）
            if closed_polyline is not None:
                try:
                    self._call_attr(closed_polyline, "Delete")
                except Exception:
                    pass
            logger.error(f"创建填充时出错: {str(e)}")
            return None

    def zoom_extents(self) -> bool:
        """缩放视图以显示所有对象"""
        if not self.is_running():
            return False

        try:
            # ZoomExtents 是 Application 对象的方法（非 ActiveViewport）
            self._call_attr(self.app, "ZoomExtents")
            logger.info("已缩放视图以显示所有对象")
            return True
        except Exception as e:
            logger.error(f"缩放视图时出错: {str(e)}")
            return False

    def close(self) -> None:
        """关闭CAD控制器

        注意：COM 套间由 COMExecutor 线程持有，进程退出时由系统回收，
        此处仅释放引用，不再跨线程调用 CoUninitialize。
        """
        try:
            # 释放COM资源
            if self.app is not None:
                del self.app
        except:
            pass

    def _ensure_layer(self, layer_name: str):
        """确保图层存在并返回图层对象（已存在则直接返回，不改动当前图层）

        绘图时指定图层走这里：只保证实体能落到目标图层，
        不把用户正在使用的当前图层悄悄切走。
        Layers集合与Name读取都可能被忙碌的CAD拒绝，逐调用重试。
        """
        layers = self._call(getattr, self.doc, "Layers")
        for i in range(self._call(getattr, layers, "Count")):
            layer = self._call_attr(layers, "Item", i)
            # CAD图层名不区分大小写，比较时忽略大小写
            # （否则"wall"会对已存在的"WALL"重复Add，被CAD拒绝）
            if self._call(getattr, layer, "Name").casefold() == layer_name.casefold():
                return layer
        # 不存在则创建（独立撤销组）
        with self._undo_group():
            return self._call_attr(layers, "Add", layer_name)

    def create_layer(self, layer_name: str) -> bool:    # , color: Union[int, Tuple[int, int, int]] = 7
        """创建新图层并设为当前图层（已存在则直接激活）

        Args:
            layer_name: 图层名称
            color: 颜色值，可以是CAD颜色索引(int)或RGB颜色值(tuple)

        Returns:
            操作是否成功
        """
        if not self.is_running():
            return False

        try:
            # 确保图层存在（存在则复用，不存在则创建）
            new_layer = self._ensure_layer(layer_name)

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
            self._call(setattr, self.doc, "ActiveLayer", new_layer)
            logger.info(f"已创建并设为当前图层: {layer_name}")  #, 颜色: {color}
            return True
        except Exception as e:
            logger.error(f"创建图层时出错: {str(e)}")
            return False

    def add_dimension(self, start_point: Tuple[float, float, float],
                     end_point: Tuple[float, float, float],
                     text_position: Tuple[float, float, float] = None, textheight: float = 5, layer: str = None, color: int=None) -> Any:
            """添加线性标注"""
            if not self.is_running():
                return None

            try:
                # 如果未提供文本位置，自动计算
                if text_position is None:
                    # 在起点和终点之间的中点上方
                    mid_x = (start_point[0] + end_point[0]) / 2
                    mid_y = (start_point[1] + end_point[1]) / 2
                    text_position = (mid_x, mid_y + 5, 0)

                # 使用VARIANT包装坐标点数据（_pt自动补齐Z轴）
                start_array = self._pt(start_point)
                end_array = self._pt(end_point)
                text_pos_array = self._pt(text_position)

                # 添加对齐标注（独立撤销组，忙碌时在调用级自动重试）
                with self._undo_group():
                    dimension = self._call_attr(self._ms(), "AddDimAligned", start_array, end_array, text_pos_array)

                    # 设置文字高度
                    if textheight is not None:
                        self._call(setattr, dimension, "TextHeight", textheight)

                    # 设置图层、颜色
                    self._apply_style(dimension, layer, color)

                logger.info(f"已添加标注: 从 {start_point} 到 {end_point}, 图层{layer if layer else '默认'}")
                return dimension
            except Exception as e:
                logger.error(f"添加标注时出错: {str(e)}")
                return None

    # ==================== 查询类操作（实时读取CAD文档状态） ====================

    @staticmethod
    def _friendly_type(object_name: str) -> str:
        """将COM对象类型名（如AcDbLine）转换为友好类型名（如line）"""
        mapping = {
            "AcDbLine": "line",
            "AcDbCircle": "circle",
            "AcDbArc": "arc",
            "AcDbEllipse": "ellipse",
            "AcDbPolyline": "polyline",       # 轻量多段线
            "AcDb2dPolyline": "polyline",
            "AcDb3dPolyline": "polyline",
            "AcDbText": "text",
            "AcDbMText": "mtext",
            "AcDbHatch": "hatch",
            "AcDbBlockReference": "block_reference",
            "AcDbDimension": "dimension",
            "AcDbAlignedDimension": "dimension",
            "AcDbRotatedDimension": "dimension",
        }
        return mapping.get(object_name, object_name)

    def get_entity_info(self, entity) -> Dict[str, Any]:
        """获取实体概要信息（句柄/类型/图层/颜色）"""
        return {
            "handle": str(self._safe_attr(entity, "Handle", "")),
            "entity_type": self._friendly_type(str(self._safe_attr(entity, "ObjectName", ""))),
            "layer": self._safe_attr(entity, "Layer"),
            "color": self._read_color(entity),
        }

    def list_layers(self) -> List[Dict[str, Any]]:
        """列出当前图纸的所有图层信息"""
        if not self.is_running():
            return []

        layers = []
        # 集合与属性读取都可能被忙碌的CAD拒绝，逐调用重试
        active_name = self._call(getattr, self._call(getattr, self.doc, "ActiveLayer"), "Name")
        layer_set = self._call(getattr, self.doc, "Layers")
        for i in range(self._call(getattr, layer_set, "Count")):
            layer = self._call_attr(layer_set, "Item", i)
            name = self._call(getattr, layer, "Name")
            layers.append({
                "name": name,
                "color": self._read_color(layer),
                "is_current": name == active_name,
                "is_frozen": bool(self._safe_attr(layer, "Freeze", False)),
                "is_locked": bool(self._safe_attr(layer, "Lock", False)),
                "is_on": bool(self._safe_attr(layer, "LayerOn", True)),
            })
        return layers

    def list_entities(self, entity_type: str = None, limit: int = 50) -> Dict[str, Any]:
        """列出模型空间中的实体（实时查询，与CAD文档保持同步）

        Args:
            entity_type: 按类型过滤（line/circle/arc等，不区分大小写），None为不过滤
            limit: 最多返回的实体数量（大图纸时避免一次性返回过多数据）

        Returns:
            {"total": 模型空间实体总数, "entities": [实体信息]}
        """
        if not self.is_running():
            return {"total": 0, "entities": []}

        ms = self._ms()
        # Count与Item访问都可能被忙碌的CAD拒绝，逐调用重试
        total = self._call(getattr, ms, "Count")
        entities = []
        for i in range(total):
            info = self.get_entity_info(self._call_attr(ms, "Item", i))
            # 精确匹配类型名（子串匹配会让"line"把polyline也带出来）
            if entity_type and info["entity_type"].lower() != entity_type.lower():
                continue
            entities.append(info)
            if len(entities) >= limit:
                break
        return {"total": total, "entities": entities}

    def get_entity_properties(self, handle: str) -> Dict[str, Any]:
        """按句柄获取实体详细属性（几何信息随实体类型而异）"""
        if not self.is_running():
            raise RuntimeError("CAD未运行")

        entity = self._get_entity_by_handle(handle)
        object_name = str(self._safe_attr(entity, "ObjectName", ""))
        props: Dict[str, Any] = {
            "handle": str(handle),
            "entity_type": self._friendly_type(object_name),
            "layer": self._safe_attr(entity, "Layer"),
            "color": self._read_color(entity),
            "linetype": self._safe_attr(entity, "Linetype"),
            "lineweight": self._safe_attr(entity, "Lineweight"),
        }

        # 按实体类型补充几何属性
        detail: Dict[str, Any] = {}
        if object_name == "AcDbLine":
            detail["start_point"] = list(self._safe_attr(entity, "StartPoint") or [])
            detail["end_point"] = list(self._safe_attr(entity, "EndPoint") or [])
            detail["length"] = self._safe_attr(entity, "Length")
        elif object_name == "AcDbCircle":
            detail["center"] = list(self._safe_attr(entity, "Center") or [])
            detail["radius"] = self._safe_attr(entity, "Radius")
            detail["area"] = self._safe_attr(entity, "Area")
        elif object_name == "AcDbArc":
            detail["center"] = list(self._safe_attr(entity, "Center") or [])
            detail["radius"] = self._safe_attr(entity, "Radius")
            # StartAngle/EndAngle为弧度，转换为度便于理解
            start_rad = self._safe_attr(entity, "StartAngle")
            end_rad = self._safe_attr(entity, "EndAngle")
            detail["start_angle"] = math.degrees(start_rad) if start_rad is not None else None
            detail["end_angle"] = math.degrees(end_rad) if end_rad is not None else None
        elif object_name == "AcDbEllipse":
            detail["center"] = list(self._safe_attr(entity, "Center") or [])
            detail["radius_ratio"] = self._safe_attr(entity, "RadiusRatio")
            detail["start_param"] = self._safe_attr(entity, "StartParam")
            detail["end_param"] = self._safe_attr(entity, "EndParam")
        elif object_name == "AcDbText":
            detail["text"] = self._safe_attr(entity, "TextString")
            detail["insertion_point"] = list(self._safe_attr(entity, "InsertionPoint") or [])
            detail["height"] = self._safe_attr(entity, "Height")
            rotation_rad = self._safe_attr(entity, "Rotation")
            detail["rotation"] = math.degrees(rotation_rad) if rotation_rad is not None else None
        elif object_name == "AcDbMText":
            detail["text"] = self._safe_attr(entity, "TextString")
            detail["insertion_point"] = list(self._safe_attr(entity, "InsertionPoint") or [])
        elif object_name in ("AcDbPolyline", "AcDb2dPolyline", "AcDb3dPolyline"):
            # Coordinates为扁平坐标数组，按类型确定每点分量数切分：
            # 轻量多段线(AcDbPolyline)为2分量[x,y,...]，
            # 旧式2D/3D多段线(AcDb2dPolyline/AcDb3dPolyline)为3分量[x,y,z,...]，
            # 统一按2切分会把z当下一点的x，坐标全部错位
            coords = self._safe_attr(entity, "Coordinates") or []
            step = 2 if object_name == "AcDbPolyline" else 3
            detail["coordinates"] = [list(coords[i:i + step]) for i in range(0, len(coords), step)]
            detail["closed"] = bool(self._safe_attr(entity, "Closed", False))
            detail["length"] = self._safe_attr(entity, "Length")
        elif object_name == "AcDbHatch":
            detail["pattern_name"] = self._safe_attr(entity, "PatternName")
            detail["area"] = self._safe_attr(entity, "Area")
        elif object_name == "AcDbBlockReference":
            detail["name"] = self._safe_attr(entity, "Name")
            detail["insertion_point"] = list(self._safe_attr(entity, "InsertionPoint") or [])
            detail["x_scale"] = self._safe_attr(entity, "XScaleFactor")
            detail["y_scale"] = self._safe_attr(entity, "YScaleFactor")
            rotation_rad = self._safe_attr(entity, "Rotation")
            detail["rotation"] = math.degrees(rotation_rad) if rotation_rad is not None else None

        props["properties"] = detail
        return props

    def get_drawing_stats(self) -> Dict[str, Any]:
        """获取当前文档的实时统计信息（实体数/图层数/文档名）"""
        if not self.is_running():
            return {"running": False}
        return {
            "running": True,
            # 各属性读取都可能被忙碌的CAD拒绝，逐调用重试
            "entity_count": self._call(getattr, self._ms(), "Count"),
            "layer_count": self._call(getattr, self._call(getattr, self.doc, "Layers"), "Count"),
            "document_name": self._call(getattr, self.doc, "Name"),
        }

    def screenshot(self) -> bytes:
        """截取CAD主窗口画面，返回PNG格式字节流"""
        if not self.is_running():
            raise RuntimeError("CAD未运行")

        import win32gui
        import win32con
        from PIL import ImageGrab

        # 截图前强制全视口重生成：浩辰CAD经COM批量绘图后画布不会自动刷新，
        # 不重生成会截到绘制前的旧画面（实体已入库，仅显示滞后）
        try:
            self._call_attr(self.doc, "Regen", 1)  # 1 = acAllViewports
        except Exception as e:
            logger.warning(f"截图前重生成失败（仍会尝试截图）: {str(e)}")

        # CAD主窗口句柄（属性读取可能被忙碌的CAD拒绝，需带重试）
        hwnd = self._call(getattr, self.app, "HWND")
        # 窗口最小化时GetWindowRect返回的是无效坐标（-32000附近），
        # 直接截会得到垃圾图，先把窗口还原
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)
        # 尽量把CAD窗口带到前台，避免被其他窗口遮挡
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
        except Exception as e:
            logger.warning(f"无法将CAD窗口置前（仍会尝试截图）: {str(e)}")

        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        # all_screens=True 以支持多显示器/负坐标场景
        img = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    # ==================== 编辑类操作 ====================

    def _get_entity_by_handle(self, handle: str):
        """按句柄查找实体对象（查找为只读操作，调用级重试安全）"""
        entity = self._call_attr(self.doc, "HandleToObject", str(handle))
        if entity is None:
            raise ValueError(f"未找到句柄为 {handle} 的实体")
        return entity

    def erase_entity(self, handle: str) -> bool:
        """删除指定实体"""
        if not self.is_running():
            return False
        try:
            # 独立撤销组：UNDO 1正好撤销本次删除
            with self._undo_group():
                self._call_attr(self._get_entity_by_handle(handle), "Delete")
            return True
        except Exception as e:
            logger.error(f"删除实体失败: {str(e)}")
            return False

    def move_entity(self, handle: str, displacement: List[float]) -> bool:
        """移动实体（按位移向量 [dx, dy, dz]）"""
        if not self.is_running():
            return False
        try:
            entity = self._get_entity_by_handle(handle)
            # Move(基点, 第二点) 实现平移：基点取原点，第二点即位移向量
            with self._undo_group():
                self._call_attr(entity, "Move", self._pt((0, 0, 0)), self._pt(displacement))
            return True
        except Exception as e:
            logger.error(f"移动实体失败: {str(e)}")
            return False

    def rotate_entity(self, handle: str, base_point: List[float], angle_degrees: float) -> bool:
        """绕基点旋转实体（角度单位：度，逆时针为正）"""
        if not self.is_running():
            return False
        try:
            entity = self._get_entity_by_handle(handle)
            with self._undo_group():
                self._call_attr(entity, "Rotate", self._pt(base_point), math.radians(angle_degrees))
            return True
        except Exception as e:
            logger.error(f"旋转实体失败: {str(e)}")
            return False

    def scale_entity(self, handle: str, base_point: List[float], scale_factor: float) -> bool:
        """绕基点缩放实体"""
        if not self.is_running():
            return False
        try:
            entity = self._get_entity_by_handle(handle)
            with self._undo_group():
                self._call_attr(entity, "ScaleEntity", self._pt(base_point), scale_factor)
            return True
        except Exception as e:
            logger.error(f"缩放实体失败: {str(e)}")
            return False

    def copy_entity(self, handle: str, displacement: List[float]):
        """复制实体并平移指定位移，返回新实体"""
        if not self.is_running():
            return None
        try:
            entity = self._get_entity_by_handle(handle)
            # 复制+平移打包为一个撤销组（Copy与Move分开重试，均幂等安全）
            with self._undo_group():
                copy = self._call_attr(entity, "Copy")
                if copy is not None and displacement:
                    self._call_attr(copy, "Move", self._pt((0, 0, 0)), self._pt(displacement))
            return copy
        except Exception as e:
            logger.error(f"复制实体失败: {str(e)}")
            return None

    def mirror_entity(self, handle: str, point1: List[float], point2: List[float]):
        """沿两点确定的镜像轴镜像实体，返回镜像产生的新实体（保留原实体）"""
        if not self.is_running():
            return None
        try:
            entity = self._get_entity_by_handle(handle)
            with self._undo_group():
                mirrored = self._call_attr(entity, "Mirror", self._pt(point1), self._pt(point2))
            return mirrored
        except Exception as e:
            logger.error(f"镜像实体失败: {str(e)}")
            return None

    def offset_entity(self, handle: str, distance: float) -> List[Any]:
        """偏移实体（正负值决定偏移方向），返回新实体列表"""
        if not self.is_running():
            return []
        try:
            entity = self._get_entity_by_handle(handle)
            with self._undo_group():
                # Offset返回新实体数组（可能为空）
                result = self._call_attr(entity, "Offset", distance)
            return list(result) if result else []
        except Exception as e:
            logger.error(f"偏移实体失败: {str(e)}")
            return []

    def array_linear_entity(self, handle: str, count: int, displacement: List[float]) -> List[Any]:
        """矩形阵列：沿位移向量方向复制 count-1 份（含原实体共 count 个）"""
        if not self.is_running():
            return []
        try:
            src = self._get_entity_by_handle(handle)
            dx, dy = float(displacement[0]), float(displacement[1])
            dz = float(displacement[2]) if len(displacement) > 2 else 0.0
            copies = []
            # 整个阵列打包为一个撤销组；循环内逐调用重试避免部分完成后重复
            with self._undo_group():
                for i in range(1, count):
                    copy = self._call_attr(src, "Copy")
                    self._call_attr(copy, "Move", self._pt((0, 0, 0)), self._pt((dx * i, dy * i, dz * i)))
                    copies.append(copy)
            return copies
        except Exception as e:
            logger.error(f"线性阵列失败: {str(e)}")
            return []

    def array_polar_entity(self, handle: str, count: int, center: List[float], total_angle_degrees: float = 360.0) -> List[Any]:
        """环形阵列：绕中心点在 total_angle 角度范围内复制 count-1 份"""
        if not self.is_running():
            return []
        try:
            src = self._get_entity_by_handle(handle)
            step = math.radians(total_angle_degrees / count)
            center_pt = self._pt(center)
            copies = []
            # 整个阵列打包为一个撤销组；循环内逐调用重试避免部分完成后重复
            with self._undo_group():
                for i in range(1, count):
                    copy = self._call_attr(src, "Copy")
                    self._call_attr(copy, "Rotate", center_pt, step * i)
                    copies.append(copy)
            return copies
        except Exception as e:
            logger.error(f"环形阵列失败: {str(e)}")
            return []

    def undo(self) -> bool:
        """撤销上一步操作"""
        if not self.is_running():
            return False
        try:
            # 注意：实测 AutoCAD 上 "_.UNDO 1" 带前缀+数字参数的组合不生效，
            # 纯英文命令名 "UNDO 1" 可靠（中文版CAD原生支持英文命令名）
            self._call_attr(self.doc, "SendCommand", "UNDO 1\n")
            # SendCommand是异步的：等待命令处理完再返回，确保后续调用能看到撤销结果
            time.sleep(self.command_delay)
            return True
        except Exception as e:
            logger.error(f"撤销失败: {str(e)}")
            return False

    def redo(self) -> bool:
        """重做被撤销的操作"""
        if not self.is_running():
            return False
        try:
            self._call_attr(self.doc, "SendCommand", "_.REDO\n")
            # SendCommand是异步的：等待命令处理完再返回
            time.sleep(self.command_delay)
            return True
        except Exception as e:
            logger.error(f"重做失败: {str(e)}")
            return False

    # ==================== 图层操作 ====================

    def set_current_layer(self, layer_name: str) -> bool:
        """切换当前图层（图层必须已存在）"""
        if not self.is_running():
            return False
        try:
            # Layers集合与Name读取都可能被忙碌的CAD拒绝，逐调用重试
            layers = self._call(getattr, self.doc, "Layers")
            for i in range(self._call(getattr, layers, "Count")):
                layer = self._call_attr(layers, "Item", i)
                # CAD图层名不区分大小写，比较时忽略大小写
                if self._call(getattr, layer, "Name").casefold() == layer_name.casefold():
                    self._call(setattr, self.doc, "ActiveLayer", layer)
                    logger.info(f"已切换当前图层: {layer_name}")
                    return True
            logger.warning(f"图层 {layer_name} 不存在")
            return False
        except Exception as e:
            logger.error(f"切换图层失败: {str(e)}")
            return False

    # ==================== 图纸管理 ====================

    def new_drawing(self) -> bool:
        """新建图纸文档"""
        if self.app is None:
            return False
        try:
            # Documents集合的获取与Add调用都可能被忙碌的CAD拒绝，逐调用重试
            docs = self._call(getattr, self.app, "Documents")
            self.doc = self._call_attr(docs, "Add")
            # Name读取可能被刚创建文档的CAD短暂拒绝，需带重试（否则日志误显示None）
            logger.info(f"已新建文档: {self._call(getattr, self.doc, 'Name')}")
            return self.doc is not None
        except Exception as e:
            logger.error(f"新建文档失败: {str(e)}")
            return False

    def open_drawing(self, file_path: str) -> bool:
        """打开已有图纸文件"""
        if self.app is None:
            return False
        # 相对路径必须先转为绝对路径：Documents.Open在CAD进程内解析路径，
        # 其工作目录与MCP服务器的工作目录通常不同（与save_drawing同理）
        file_path = os.path.abspath(file_path)
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        try:
            # 打开后CAD短暂忙碌，调用级重试确保Open本身成功
            docs = self._call(getattr, self.app, "Documents")
            opened = self._call_attr(docs, "Open", file_path)
            # 部分CAD的Documents.Open不返回文档对象，回退到取活动文档
            self.doc = opened if opened is not None else self._call(getattr, self.app, "ActiveDocument")
            logger.info(f"已打开文档: {file_path}")
            return True
        except Exception as e:
            logger.error(f"打开文档失败: {str(e)}")
            return False

    def close_drawing(self, save: bool = False) -> bool:
        """关闭当前文档（save=False 时不保存直接关闭）"""
        if not self.is_running():
            return False
        try:
            # 关闭期间CAD忙碌会拒绝调用，调用级重试确保Close本身成功
            self._call_attr(self.doc, "Close", save)
            # 尝试切换到其他已打开的文档
            try:
                self.doc = self._call(getattr, self.app, "ActiveDocument")
            except Exception:
                self.doc = None
            logger.info(f"已关闭当前文档（保存: {save}）")
            return True
        except Exception as e:
            logger.error(f"关闭文档失败: {str(e)}")
            return False

    # ==================== 图块操作 ====================

    def create_block(self, name: str, base_point: List[float] = None):
        """创建图块定义（若同名图块已存在则直接返回它）

        注意：图块内部实体可通过 send_command 执行 BEDIT 等命令编辑。
        """
        if not self.is_running():
            return None
        try:
            # 检查同名图块是否已存在（Blocks集合与Name读取都可能被忙碌的CAD拒绝，逐调用重试）
            blocks = self._call(getattr, self.doc, "Blocks")
            for i in range(self._call(getattr, blocks, "Count")):
                block = self._call_attr(blocks, "Item", i)
                if self._call(getattr, block, "Name") == name:
                    return block
            # 独立撤销组，忙碌时在调用级自动重试
            with self._undo_group():
                return self._call_attr(blocks, "Add", self._pt(base_point or (0, 0, 0)), name)
        except Exception as e:
            logger.error(f"创建图块失败: {str(e)}")
            return None

    def insert_block(self, name: str, position: List[float],
                     x_scale: float = 1.0, y_scale: float = 1.0, z_scale: float = 1.0,
                     rotation_degrees: float = 0.0):
        """在模型空间插入图块引用"""
        if not self.is_running():
            return None
        try:
            # 独立撤销组，忙碌时在调用级自动重试
            with self._undo_group():
                block_ref = self._call_attr(
                    self._ms(), "InsertBlock",
                    self._pt(position), name, x_scale, y_scale, z_scale,
                    math.radians(rotation_degrees)
                )
            return block_ref
        except Exception as e:
            logger.error(f"插入图块失败: {str(e)}")
            return None

    def list_blocks(self) -> List[Dict[str, Any]]:
        """列出当前图纸中的用户图块定义"""
        if not self.is_running():
            return []
        blocks = []
        try:
            # Blocks集合与Name读取都可能被忙碌的CAD拒绝，逐调用重试
            block_set = self._call(getattr, self.doc, "Blocks")
            for i in range(self._call(getattr, block_set, "Count")):
                block = self._call_attr(block_set, "Item", i)
                name = str(self._call(getattr, block, "Name"))
                # 跳过布局占位块（*Model_Space、*Paper_Space等匿名块）
                if not name.startswith("*"):
                    blocks.append({
                        "name": name,
                        "entity_count": self._safe_attr(block, "Count", 0),
                    })
        except Exception as e:
            logger.error(f"列出图块失败: {str(e)}")
        return blocks

    # ==================== 命令直通 ====================

    def send_command(self, command: str) -> bool:
        """直通CAD命令行执行任意命令（如 "_.ZOOM E"）

        这是功能逃生舱：任何未封装为工具的CAD命令都可以通过它执行。
        """
        if not self.is_running():
            return False
        try:
            # 若命令未以回车/空格结尾则补一个回车，确保命令被提交执行
            if not command.endswith(("\n", "\r", " ")):
                command += "\n"
            # SendCommand本身可能被忙碌的CAD拒绝，调用级重试确保命令入队
            self._call_attr(self.doc, "SendCommand", command)
            # SendCommand是异步的：等待命令执行（如ZOOM/保存类命令执行中，
            # 后续的SaveAs等操作会报错），时长由配置command_delay控制
            time.sleep(self.command_delay)
            return True
        except Exception as e:
            logger.error(f"执行命令失败: {str(e)}")
            return False
