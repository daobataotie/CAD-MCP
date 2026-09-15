# CAD-MCP Server (CAD Model Context Protocol Server)

[English](/README_en.md) | [中文](/README_zh.md) 

## 项目介绍

CAD-MCP是一个创新的CAD控制服务，允许通过自然语言指令控制CAD软件进行绘图操作。该项目结合了自然语言处理和CAD自动化技术，使用户能够通过简单的文本命令创建和修改CAD图纸，无需手动操作CAD界面。

v2.0 版本基于最新 MCP 规范（FastMCP、结构化输出、工具注解、进度通知、Streamable HTTP）全面升级，并新增了查询、编辑、图层、图纸管理、图块与命令直通共 35 个工具，形成完整的"绘制-查看-修正"闭环。

## 功能特点

### CAD控制功能

- **多CAD软件支持**：支持AutoCAD、浩辰CAD(GCAD/GstarCAD)和中望CAD(ZWCAD)等主流CAD软件
- **基础绘图功能**：
  - 直线绘制
  - 圆形绘制
  - 圆弧绘制
  - 椭圆绘制
  - 矩形绘制
  - 多段线绘制
  - 文本添加
  - 图案填充
  - 尺寸标注
- **查询功能（AI能看到图纸实际内容）**：
  - 图层列表查询
  - 实体列表查询（返回可编辑的实体句柄）
  - 实体详细属性查询（坐标/半径/文本内容等）
  - CAD窗口截图预览
- **编辑功能**：删除、移动、旋转、缩放、复制、镜像、偏移、阵列（线性/环形）、撤销/重做
- **图层管理**：创建和切换图层、查询图层状态
- **图纸管理**：新建、打开、关闭、保存图纸（DWG/DXF）
- **图块操作**：创建、插入、查询图块
- **命令直通**：通过 `send_command` 执行任意CAD命令（功能逃生舱）

### MCP 协议特性

- **结构化输出**：所有工具返回带 `outputSchema` 的 `structuredContent`，客户端可直接解析结果
- **工具注解**：查询类工具标记 `readOnlyHint`，破坏性操作标记 `destructiveHint`，便于客户端做确认策略
- **进度通知**：CAD启动/保存等耗时操作会向客户端发送消息通知
- **双传输模式**：支持 `stdio`（本地单客户端）与 `streamable-http`（远程/多客户端共享同一CAD实例）
- **错误语义**：参数校验错误作为工具错误返回，支持模型自我修正

### 自然语言处理功能（遗留接口）

- **命令解析**：将自然语言指令解析为CAD操作参数
- **颜色识别**：从文本中提取颜色信息并应用到绘图对象
- **形状关键词映射**：支持多种形状描述词的识别
- **动作关键词映射**：识别各种绘图和编辑动作

## Demo

The following is the demo video.

![Demo](imgs/demo.gif)

## 安装要求

### 依赖项

```
pywin32>=306       # Windows COM接口支持
mcp>=1.12,<2       # MCP Python SDK（含FastMCP、结构化输出、Streamable HTTP）
pydantic>=2.0.0    # 数据验证
Pillow>=10.0.0     # 截图预览PNG编码
```

### 系统要求

- Windows操作系统
- 已安装的CAD软件（AutoCAD、浩辰CAD或中望CAD）
- Python 3.10+

## 配置说明

配置文件位于`src/config.json`，包含以下主要设置：

```json
{
    "server": {
        "name": "CAD MCP 服务器",
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
```

- **server**: 服务器名称、版本与传输配置
  - `transport`: 传输方式，`stdio`（默认，本地单客户端）或 `streamable-http`（远程/多客户端）
  - `host` / `port`: Streamable HTTP 模式的监听地址和端口（也可通过命令行 `--host` / `--port` 覆盖）
- **cad**: 
  - `type`: CAD软件类型（AutoCAD、GCAD、GstarCAD或ZWCAD）
  - `startup_wait_time`: CAD启动等待时间（秒）
  - `command_delay`: 命令执行延迟（秒）
- **output**: 输出文件设置

## 使用方法

### 启动服务

```bash
# 方式一：直接运行（与旧版配置完全兼容）
python src/server.py

# 方式二：安装为命令行工具
pip install -e .
cad-mcp

# 方式三：以 Streamable HTTP 传输运行（远程/多客户端共享同一CAD实例）
python src/server.py --transport streamable-http --port 8000
# 客户端连接 http://127.0.0.1:8000/mcp
```

### 接入 MCP 客户端

服务器默认使用 `stdio` 传输，同一份 JSON 配置适用于所有 MCP 客户端（Claude Desktop、Cursor、Windsurf、TRAE 等）。将其添加到客户端的 MCP 服务器配置文件中（如 `claude_desktop_config.json`、Cursor 的 `mcp.json`），路径替换为你自己的：

```json
{
    "mcpServers": {
        "CAD": {
            "command": "python",
            "args": [
                # 你的实际路径，如："C:\\cad-mcp\\src\\server.py"
            ]
        }
    }
}
```

若使用 Streamable HTTP 模式，客户端改为连接 `http://127.0.0.1:8000/mcp`。

### MCP Inspector

```bash
# Note：use your path  
npx -y @modelcontextprotocol/inspector python C:\\cad-mcp\\src\\server.py
```

### 服务API

服务器提供以下工具（共35个）：

**基本绘图：**
- `draw_line`: 绘制直线
- `draw_circle`: 绘制圆
- `draw_arc`: 绘制弧
- `draw_ellipse`: 绘制椭圆
- `draw_polyline`: 绘制多段线
- `draw_rectangle`: 绘制矩形
- `draw_text`: 添加文本
- `draw_hatch`: 绘制填充
- `add_dimension`: 添加线性标注

**查询：**
- `list_layers`: 列出所有图层
- `list_entities`: 列出模型空间实体（返回可编辑的句柄）
- `get_entity_properties`: 按句柄查询实体详细属性
- `screenshot`: 截取CAD窗口画面确认绘图效果

**编辑：**
- `erase_entity` / `move_entity` / `rotate_entity` / `scale_entity`: 删除/移动/旋转/缩放实体
- `copy_entity` / `mirror_entity` / `offset_entity`: 复制/镜像/偏移实体
- `array_linear_entity` / `array_polar_entity`: 线性/环形阵列
- `undo` / `redo`: 撤销/重做

**图层与图纸管理：**
- `create_layer` / `set_current_layer`: 图层管理
- `new_drawing` / `open_drawing` / `close_drawing` / `save_drawing`: 图纸管理
- `create_block` / `insert_block` / `list_blocks`: 图块操作
- `send_command`: 直通CAD命令行执行任意命令

**遗留接口：**
- `process_command`: 处理自然语言命令

## 项目结构

```
CAD-MCP/
├── imgs/                # 图像和视频资源
│   └── CAD-mcp.mp4     # 演示视频
├── pyproject.toml       # 打包与安装配置（pip install -e .）
├── requirements.txt     # 项目依赖
└── src/                 # 源代码
    ├── __init__.py     # 包初始化
    ├── cad_controller.py # CAD控制器（COM层：绘图/查询/编辑/图纸/图块）
    ├── config.py        # 配置单例加载
    ├── config.json     # 配置文件
    ├── models.py        # 工具结构化输出模型（outputSchema）
    ├── nlp_processor.py # 自然语言处理器
    └── server.py       # MCP服务器实现（FastMCP）
```

## 许可证

MIT License
