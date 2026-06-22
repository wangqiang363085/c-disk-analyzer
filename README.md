# C 盘空间分析工具

一个 Windows 专用的 C 盘空间分析工具，**无需安装任何依赖**，双击即用。

## 功能

- 扫描 C 盘各目录大小，找出空间占用大户
- 检测 AppData 缓存（剪映、美图、WPS、微信开发者工具等）
- 生成带 **饼图** 的 Excel 报告 + 清理建议
- 可选一键清理 npm / pip 缓存
- 全中文界面

## 使用方法

### 方式一：直接运行 exe（推荐普通用户）

```
dist/C盘分析工具.exe
```

**建议右键 → 以管理员身份运行**，扫描结果更完整。

### 方式二：运行 Python 源码（推荐开发者）

```bash
pip install openpyxl
python c_disk_analyzer.py
```

### 方式三：批处理版（兼容 XP ~ Win11）

```
C盘分析工具.bat
```

无需 Python，双击运行，生成文本报告 + CSV。

## 输出

运行后自动在桌面生成：
- `C盘分析报告_时间戳.xlsx` — 带饼图的 Excel 报告（仅 exe / py 版本）
- `C盘目录说明.csv` — 各目录能否删除的速查表（仅 bat 版本）

## 项目结构

```
c-disk-analyzer/
├── c_disk_analyzer.py      # Python 源码（Win7+，需安装 Python）
├── C盘分析工具.bat          # 批处理版本（XP ~ Win11 全兼容）
├── requirements.txt         # Python 依赖
├── .gitignore
└── README.md
```

## 注意事项

- 该工具**只读不写**，不会自动删除任何文件
- 如需清理缓存，运行时会询问用户确认
- npm / pip 缓存清理命令需已安装对应环境
