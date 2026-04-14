# EnglishSpiker

一个 Windows 原生桌面小工具：复制英语单词后自动朗读，并显示中文意思。

## 使用方法

1. 双击 `start_english_reader.bat`。
2. 程序会显示一个小浮窗，默认始终置顶。
3. 复制一个英语单词，程序会自动朗读并显示中文释义。
4. 按默认快捷键 `Ctrl+Alt+R` 可以重读剪贴板里的单词。
5. 点窗口右上角的 `×` 不会退出程序，只会隐藏到屏幕右下角托盘。
6. 右键小浮窗或右键托盘图标都可以打开设置菜单。

## 说明

- 朗读使用 Windows 自带语音功能，项目使用 C# / WPF。
- 设置菜单里有：显示/隐藏、自动朗读、自动释义、置顶、声音、语速、快捷键、退出。
- 声音好不好听取决于 Windows 已安装的语音；可以在托盘右键菜单里切换。
- 中文释义优先使用在线词典接口，网络不可用时会给出提示。

## 命令行启动

```powershell
dotnet run --project EnglishSpiker.csproj
```

## 直接运行

打包后可以直接双击：

```text
publish\EnglishSpiker.exe
```

## 构建检查

```powershell
dotnet build EnglishSpiker.csproj
```

## 打包

```powershell
dotnet publish EnglishSpiker.csproj -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o publish
```
