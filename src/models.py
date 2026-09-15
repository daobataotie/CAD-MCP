"""工具结构化输出模型（MCP Structured Output，2025-06-18 规范）

所有工具统一返回 Pydantic 模型，FastMCP 会据此自动生成 outputSchema，
客户端（Claude / Cursor 等）可直接获得结构化的 structuredContent，
而不是此前难以解析的 COM 对象字符串。
"""

from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


class ToolResult(BaseModel):
    """通用工具执行结果"""
    success: bool = Field(description="执行是否成功")
    message: str = Field(description="人类可读的结果说明")
    entity_handles: List[str] = Field(
        default_factory=list,
        description="本次操作涉及/产生的实体句柄列表，可用于后续编辑或查询"
    )
    error: Optional[str] = Field(default=None, description="失败时的错误信息")


class EntityInfo(BaseModel):
    """实体概要信息"""
    handle: str = Field(description="实体句柄（CAD内唯一标识）")
    entity_type: str = Field(description="实体类型（line/circle/arc等）")
    layer: Optional[str] = Field(default=None, description="所在图层名称")
    color: Optional[int] = Field(default=None, description="颜色索引（ACI）")


class ListEntitiesResult(BaseModel):
    """实体列表查询结果"""
    success: bool = Field(description="查询是否成功")
    message: str = Field(description="结果说明")
    total: int = Field(default=0, description="模型空间实体总数")
    returned: int = Field(default=0, description="本次实际返回的实体数量（受limit限制）")
    entities: List[EntityInfo] = Field(default_factory=list, description="实体列表")


class LayerInfo(BaseModel):
    """图层信息"""
    name: str = Field(description="图层名称")
    color: Optional[int] = Field(default=None, description="图层颜色索引（ACI）")
    is_current: bool = Field(default=False, description="是否为当前图层")
    is_frozen: bool = Field(default=False, description="是否冻结")
    is_locked: bool = Field(default=False, description="是否锁定")
    is_on: bool = Field(default=True, description="是否打开（可见）")


class ListLayersResult(BaseModel):
    """图层列表查询结果"""
    success: bool = Field(description="查询是否成功")
    message: str = Field(description="结果说明")
    layers: List[LayerInfo] = Field(default_factory=list, description="图层列表")


class EntityPropertiesResult(BaseModel):
    """实体详细属性查询结果"""
    success: bool = Field(description="查询是否成功")
    message: str = Field(description="结果说明")
    handle: str = Field(default="", description="实体句柄")
    entity_type: str = Field(default="", description="实体类型")
    layer: Optional[str] = Field(default=None, description="所在图层")
    color: Optional[int] = Field(default=None, description="颜色索引（ACI）")
    linetype: Optional[str] = Field(default=None, description="线型名称")
    lineweight: Optional[int] = Field(default=None, description="线宽")
    properties: Dict[str, Any] = Field(
        default_factory=dict,
        description="按实体类型而异的几何属性（坐标/半径/文本内容等）"
    )
    error: Optional[str] = Field(default=None, description="失败时的错误信息")


class BlockInfo(BaseModel):
    """图块信息"""
    name: str = Field(description="图块名称")
    entity_count: int = Field(default=0, description="图块内实体数量")


class ListBlocksResult(BaseModel):
    """图块列表查询结果"""
    success: bool = Field(description="查询是否成功")
    message: str = Field(description="结果说明")
    blocks: List[BlockInfo] = Field(default_factory=list, description="图块列表")
