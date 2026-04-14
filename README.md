# EnglishSpiker

一个 Windows 桌面小工具：复制英语单词后自动朗读，并显示中文意思。

## 使用方法

1. 双击 `start_english_reader.bat`。
2. 复制一个英语单词，程序会自动朗读。
3. 按默认快捷键 `Ctrl+Alt+R` 可以重读剪贴板里的单词。
4. 在窗口里可以换声音、调语速、修改快捷键。

## 说明

- 朗读使用 Windows 自带语音功能，不需要额外安装 Python 包。
- 声音好不好听取决于 Windows 已安装的语音；可以在程序里的“声音”下拉框切换。
- 中文释义优先使用在线词典接口，网络不可用时会给出提示。

## 命令行启动

```powershell
py english_reader_app.py
```
