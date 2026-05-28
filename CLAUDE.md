# 订货单生成器 (Order Form Generator)

> 零技术门槛的桌面程序 — 传模板、选照片、填信息、一键生成订货单。

## 目录结构

```
order-app/
├── src/
│   ├── main.py              # 入口 — python src/main.py
│   ├── gui.py               # Tkinter 界面
│   ├── generator.py         # xlsx 生成
│   ├── parser.py            # 文本解析
│   └── template_reader.py   # 模板解析
├── docs/                    # 文档
├── dist/                    # PyInstaller 打包输出
├── requirements.txt         # 依赖
└── CLAUDE.md
```

## 红线

- **不写死模板格式** — 必须通过 template_reader 读取模板动态生成
- **不使用 data_only=True 后保存** — 会丢失公式，永久替换为值
- **公式优先** — 总数量、总金额用 Excel 公式，不用 Python 计算硬编码
- **中文变量/注释** — 面向用户的部分用中文，方便用户理解

## 命令速查

```bash
python src/main.py            # 启动 GUI
pip install -r requirements.txt  # 安装依赖
pyinstaller --onefile --windowed --add-data "src;src" src/main.py  # 打包
```

## 深入文档

- [架构说明](docs/architecture.md) — 模块职责、数据流
- [用户手册](docs/user-guide.md) — 妈妈看的操作指南
- [打包指南](docs/build-guide.md) — PyInstaller 打包步骤
