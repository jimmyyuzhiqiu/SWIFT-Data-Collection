# 打包成 EXE 说明

## 快速打包（推荐）

### Windows 用户
直接双击运行 `build.bat`，自动完成以下步骤：
1. 检查 Python 环境
2. 安装依赖
3. 打包成 exe

输出文件：`dist/SWIFT_Data_Collection.exe`

---

## 手动打包

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. 运行打包脚本
```bash
python build.py
```

### 3. 获取 exe
打包完成后，exe 文件位于：`dist/SWIFT_Data_Collection.exe`

---

## 打包参数说明

- `--onefile`: 生成单文件 exe（所有依赖打包在一个文件中）
- `--windowed`: 无控制台窗口（GUI 应用）
- `--hidden-import`: 显式导入隐藏的模块
- `--collect-all`: 收集所有相关文件

---

## 常见问题

### Q: 打包后 exe 很大？
A: 这是正常的，因为包含了 PySide6、pandas 等大型库。首次运行会解压到临时目录。

### Q: 如何减小 exe 大小？
A: 可以使用 UPX 压缩，但会增加启动时间。不推荐。

### Q: 能否分发 exe？
A: 可以，exe 是独立的，无需 Python 环境。直接分发即可。

### Q: 如何修改 icon？
A: 将 `app.ico` 放在项目根目录，重新打包即可。

---

## 环境要求

- Python 3.9+
- Windows 7 或更高版本
- 至少 2GB 可用磁盘空间（打包过程）
