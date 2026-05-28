# 打包指南

将 Python 程序打包为单文件 .exe，妈妈双击即用。

## 前提

已安装 Python 和所需依赖：

```bash
pip install -r requirements.txt
pip install pyinstaller
```

## 打包命令

```bash
pyinstaller --onefile --windowed --name "订货单生成器" src/main.py
```

参数说明：
- `--onefile`：打包成单个 .exe 文件
- `--windowed`：不显示控制台窗口（纯 GUI）
- `--name`：输出文件名

## 输出

打包后的 .exe 在 `dist/订货单生成器.exe`

## 常见问题

**杀毒软件报毒？**
PyInstaller 打包的程序有时会被误报。可以添加白名单，或者用 UPX 压缩壳。

**文件太大？**
安装 upx 并加 `--upx-dir` 参数可以减小体积。

**在别的电脑上打不开？**
确保打包时用的 Python 架构和目标电脑一致（都是 64 位 Windows）。
