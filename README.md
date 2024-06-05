# Office Automation Tools

软件测试自动化办公工具集

## 🔆介绍

🤗软件测试自动化办公工具，目前已经实现了单元测试表格转文档。

## 💫使用

你可以在[github](https://github.com/anyfan/yy_oa_tools/releases)下载最新版本

1. 命令行

```bash
./oa_tools.exe 单元测试协同表格.xlsx
```

2. 界面

拖动`单元测试协同表格.xlsx`文件到`oa_tools.exe`文件，完成转换。

## ⚙️开发构建

### 安装依赖库
```bash
pip install -r requirements.txt
```

### 打包成可执行文件
```bash
Pyinstaller build.spec
```

### 其他

1. `docx`模板文件建议使用`WPS`编辑。已知`ms office`容易导致块错乱。

## 📝未来计划

- [ ] 图形化界面，pyqt5