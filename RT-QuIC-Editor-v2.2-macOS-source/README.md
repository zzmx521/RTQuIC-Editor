# RT-QuIC Editor v2.2 — macOS Source Kit

本源码包用于在 macOS 上构建 RT-QuIC Editor v2.2。源码包本身小于 25 MB；构建后的完整 DMG 因包含 Python、PyQt5 和 QtWebEngine，通常会超过 100 MB。

## Included files

- `rt_quic_editor.py` — 主程序
- `build_viewer.py` — RT-QuIC 报告生成模块
- `icons.py` — 界面图标模块
- `plotly.min.js` — 离线交互图表库
- `RT-QuIC-Editor-macOS.spec` — PyInstaller 构建配置
- `requirements-macos.txt` — Python 依赖版本
- `build_macos.sh` — macOS 一键构建脚本
- `.github/workflows/build-macos.yml` — GitHub Actions 自动构建配置

## Build locally on macOS

需要 macOS 12 或更高版本及 Python 3.12。Apple Silicon Mac 会生成原生 Apple Silicon 版本；Intel Mac 会生成 Intel 版本。

在终端中执行：

```bash
cd RT-QuIC-Editor-v2.2-macOS-source
chmod +x build_macos.sh
./build_macos.sh
```

生成的 DMG 位于 `release` 文件夹。

## Build with GitHub Actions

1. 将本文件夹内的全部文件上传到 GitHub 仓库根目录。
2. 打开仓库的 **Actions** 页面。
3. 选择 **Build macOS DMG**。
4. 点击 **Run workflow**。
5. 构建完成后，在运行记录底部下载生成的 artifact。

GitHub Actions 生成的安装包不需要从本地上传，因此不受网页上传源码时 25 MB 限制的影响。

## Notes

- 当前应用使用临时代码签名，没有 Apple Developer ID 签名或公证。
- 首次打开时，macOS 可能显示安全提示。可在 Finder 中右键应用并选择 **Open**。
- 发布前建议使用已知 RT-QuIC 数据验证导入、分析、图表、PDF/TIFF 和 Prism CSV 导出功能。
