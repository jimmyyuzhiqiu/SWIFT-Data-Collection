# SWIFT Data Collection

一个高效的 SWIFT 报文解析工具，支持从 Outlook .msg 文件中提取数据并生成 Excel 报表。

## ✨ 功能

- 📧 **SWIFT 报文解析** - 从 .msg 文件中自动提取 SWIFT 字段
- 📊 **Excel 生成** - 生成 `Step3_Final` 和 `Debug` 两个 sheet
- 🔄 **数据回写** - 将 CP SWIFT 信息回写到 DW Excel（账号优先、金额兜底匹配）
- 🎨 **图形界面** - PySide6 深色主题 GUI，一键运行
- 📦 **独立 EXE** - 支持打包成单文件 Windows exe，无需 Python 环境

## 🚀 快速开始

### 方式一：GUI（推荐）

**Windows 用户 - 直接运行 EXE：**
```bash
SWIFT_Data_Collection.exe
```

**或从源码运行：**
```bash
python swfit_app.py
```

### 方式二：命令行

```bash
python swift_core.py
```

### 方式三：回写 DW

```bash
python update_cp_swift.py
```

## 📋 安装

### 环境要求
- Python 3.9+
- Windows 7 或更高版本

### 安装依赖

```bash
pip install -r requirements.txt
```

## 📁 项目结构

```
SWIFT-Data-Collection/
├── swfit_app.py              # GUI 入口（PySide6）
├── swift_core.py             # 核心解析逻辑
├── update_cp_swift.py        # DW 回写脚本
├── build.py                  # PyInstaller 打包脚本
├── build.bat                 # Windows 一键打包
├── requirements.txt          # 依赖列表
└── README.md                 # 本文件
```

## 🔧 使用说明

### GUI 界面

1. 启动 `swfit_app.py` 或 `SWIFT_Data_Collection.exe`
2. 确认/修改以下路径：
   - **MSG 文件夹** - 包含 .msg 文件的目录
   - **输出文件夹** - 生成 Excel 的目录
   - **Mapping 文件** - 账户映射 Excel
   - **Sheet 名称** - 默认 `ACCT Mapping`
3. 点击"▶ 运行"
4. 成功后自动打开输出的 Excel

### 核心功能

#### swift_core.py - SWIFT 报文解析

**输入：** .msg 文件

**输出：** `YYYYMMDD_Swift.xlsx`

**Step3_Final 列：**
- `Client Acct` - 客户账号
- `PRIM ID` - 主账号 ID
- `DATE` - 交易日期
- `CCY` - 货币
- `AMT` - 金额
- `CP NAME` - 交易对手名称
- `CP A/C` - 交易对手账号
- `CP SWIFT` - 交易对手 SWIFT 码
- `CP BANK NAME` - 交易对手银行名称
- `DIRECTION` - 交易方向（IN/OUT）

**字段提取规则：**

| 方向 | 货币 | 字段 | 来源 |
|------|------|------|------|
| OUT | USD/EUR | ClientAcct | 50K 或 50F Line1 |
| IN | USD | ClientAcct | 59F 或 59K Line1 |
| IN | EUR | ClientAcct | 59 Line1 |
| IN | USD/EUR | CP NAME | 50K 或 50F Line2 |
| IN | USD/EUR | CP A/C | 50K 或 50F Line1 |

#### update_cp_swift.py - DW 回写

**功能：** 将 Step3_Final 的 CP SWIFT 写回到 DW Excel

**匹配逻辑：**
1. 账号优先匹配（`CP A/C` → `交易对手存款账户编码`）
2. 金额兜底匹配（`AMT` → `存款发生金额`，范围 `[AMT-DELTA, AMT]`）

**冲突处理：**
- `ALLOW_OVERWRITE_ON_CONFLICT = False` - 保留第一次写入，标橙提示冲突
- `ALLOW_OVERWRITE_ON_CONFLICT = True` - 允许覆盖

## 📦 打包成 EXE

### 快速打包（推荐）

**Windows 用户 - 双击运行：**
```bash
build.bat
```

自动完成：
1. 检查 Python 环境
2. 安装依赖
3. 打包成 exe

**输出：** `dist/SWIFT_Data_Collection.exe`

### 手动打包

```bash
pip install -r requirements.txt
python build.py
```

## ⚙️ 配置

### Mapping 文件要求

`swift_core.py` 默认读取 `ACCT Mapping` sheet，需要以下列：
- `PRIMARY ID` - 主账号
- `CCY` - 货币
- `R-TAG` - 账户标签

### 默认路径

编辑 `swift_core.py` 顶部的配置：

```python
DEFAULT_MSG_FOLDER = r"Z:\To Jimmy Yu\Swift Data Collection\Swift"
DEFAULT_OUTPUT_FOLDER = r"Z:\To Jimmy Yu\Swift Data Collection"
DEFAULT_MAPPING_FILE = r"Z:\To Jimmy Yu\Swift Data Collection\Swift Data Collection.xlsx"
```

## 🐛 常见问题

**Q: .msg 文件解析失败？**
- A: 工具支持两种模式：
  - 优先使用 `extract-msg` 库（Outlook 格式）
  - 回退到原始文本解码（文本导出格式）

**Q: 打包后 exe 很大？**
- A: 正常现象，包含 PySide6、pandas 等库。首次运行会解压到临时目录。

**Q: 能否在 Mac/Linux 上运行？**
- A: 可以运行 Python 脚本，但 GUI 的 Excel 自动打开功能仅在 Windows 上有效。

**Q: 如何修改 GUI 样式？**
- A: 编辑 `swfit_app.py` 中的 QSS 样式表。

## 📝 许可证

MIT License

## 👤 作者

Jimmy Yu

## 🔗 相关链接

- [GitHub](https://github.com/jimmyyuzhiqiu/SWIFT-Data-Collection)
- [Issues](https://github.com/jimmyyuzhiqiu/SWIFT-Data-Collection/issues)
