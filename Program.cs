using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Forms = System.Windows.Forms;

namespace EnglishSpiker;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        var app = new Application
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown
        };
        app.Run(new MainWindow());
    }
}

internal sealed class MainWindow : Window
{
    private const int HotkeyId = 9281;
    private const int WmHotkey = 0x0312;
    private const uint ModAlt = 0x0001;
    private const uint ModControl = 0x0002;
    private const uint ModShift = 0x0004;

    private static readonly Dictionary<string, uint> HotkeyModifiers = new()
    {
        ["Ctrl+Alt"] = ModControl | ModAlt,
        ["Ctrl+Shift"] = ModControl | ModShift,
        ["Alt+Shift"] = ModAlt | ModShift,
        ["Ctrl+Alt+Shift"] = ModControl | ModAlt | ModShift,
    };

    private static readonly string[] HotkeyKeys =
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ".Select(c => c.ToString()).Concat(Enumerable.Range(1, 12).Select(i => $"F{i}")).ToArray();

    private static readonly Dictionary<string, uint> VirtualKeys =
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToDictionary(c => c.ToString(), c => (uint)c);

    private readonly System.Windows.Controls.TextBox _wordBox;
    private readonly TextBlock _meaningText;
    private readonly TextBlock _statusText;
    private readonly DispatcherTimer _clipboardTimer;
    private readonly DictionaryClient _dictionary = new();
    private readonly SpeechService _speech = new();

    private Forms.NotifyIcon? _trayIcon;
    private HwndSource? _source;
    private string _lastClipboardText = "";
    private string _lastLookupWord = "";
    private bool _autoRead = true;
    private bool _autoLookup = true;
    private bool _alwaysOnTop = true;
    private string _voice = "系统默认";
    private int _rate;
    private string _hotkeyModifier = "Ctrl+Alt";
    private string _hotkeyKey = "R";

    public MainWindow()
    {
        foreach (var number in Enumerable.Range(1, 12))
        {
            VirtualKeys[$"F{number}"] = (uint)(0x70 + number - 1);
        }

        Width = 320;
        Height = 120;
        MinWidth = 320;
        MinHeight = 120;
        MaxWidth = 380;
        MaxHeight = 160;
        ResizeMode = ResizeMode.NoResize;
        WindowStyle = WindowStyle.None;
        AllowsTransparency = false;
        Background = Brush("#f8f9fa");
        Topmost = true;
        ShowInTaskbar = false;
        Title = "Lexicon";

        var border = new Border
        {
            BorderBrush = Brush("#dbe4e7"),
            BorderThickness = new Thickness(1),
            Background = Brush("#f8f9fa"),
            Padding = new Thickness(12, 8, 12, 8),
            Child = BuildLayout(out _wordBox, out _meaningText, out _statusText)
        };
        Content = border;

        MouseLeftButtonDown += (_, _) => DragMove();
        MouseRightButtonUp += (_, _) => ShowContextMenu();
        Closing += OnClosing;
        SourceInitialized += OnSourceInitialized;

        _clipboardTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(180) };
        _clipboardTimer.Tick += (_, _) => PollClipboard();
        _clipboardTimer.Start();
    }

    private static UIElement BuildLayout(out System.Windows.Controls.TextBox wordBox, out TextBlock meaningText, out TextBlock statusText)
    {
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var header = new Grid();
        header.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        header.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        Grid.SetRow(header, 0);

        header.Children.Add(new TextBlock
        {
            Text = "Lexicon",
            Foreground = Brush("#2b3437"),
            FontSize = 11,
            FontWeight = FontWeights.Bold,
            VerticalAlignment = VerticalAlignment.Center
        });

        var close = new TextBlock
        {
            Text = "×",
            Foreground = Brush("#586064"),
            FontSize = 16,
            FontWeight = FontWeights.Bold,
            Cursor = Cursors.Hand,
            VerticalAlignment = VerticalAlignment.Center,
            Padding = new Thickness(8, 0, 0, 0)
        };
        close.MouseLeftButtonUp += (_, _) => Application.Current.MainWindow.Hide();
        Grid.SetColumn(close, 1);
        header.Children.Add(close);
        root.Children.Add(header);

        var body = new Grid { Margin = new Thickness(0, 2, 0, 0) };
        body.RowDefinitions.Add(new RowDefinition { Height = new GridLength(34) });
        body.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
        body.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        Grid.SetRow(body, 1);

        wordBox = new System.Windows.Controls.TextBox
        {
            Text = "Lexicon",
            IsReadOnly = true,
            BorderThickness = new Thickness(0),
            Background = Brush("#f8f9fa"),
            Foreground = Brush("#2b3437"),
            FontSize = 24,
            FontWeight = FontWeights.ExtraBold,
            Padding = new Thickness(0),
            VerticalContentAlignment = VerticalAlignment.Center
        };
        body.Children.Add(wordBox);

        meaningText = new TextBlock
        {
            Text = "复制英语单词后自动朗读",
            Foreground = Brush("#586064"),
            FontSize = 13,
            FontFamily = new FontFamily("Microsoft YaHei UI"),
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetRow(meaningText, 1);
        body.Children.Add(meaningText);

        var footer = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom
        };
        statusText = new TextBlock
        {
            Text = "Listening...",
            Foreground = Brush("#586064"),
            FontSize = 9,
            FontWeight = FontWeights.Bold,
            VerticalAlignment = VerticalAlignment.Center
        };
        footer.Children.Add(statusText);
        footer.Children.Add(new TextBlock
        {
            Text = " ●",
            Foreground = Brush("#005ac2"),
            FontSize = 10,
            VerticalAlignment = VerticalAlignment.Center
        });
        Grid.SetRow(footer, 2);
        body.Children.Add(footer);

        root.Children.Add(body);
        return root;
    }

    private void OnSourceInitialized(object? sender, EventArgs e)
    {
        _source = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
        _source?.AddHook(WndProc);
        RegisterHotkey();
        CreateTrayIcon();
    }

    private void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        e.Cancel = true;
        Hide();
    }

    private void CreateTrayIcon()
    {
        _trayIcon = new Forms.NotifyIcon
        {
            Icon = System.Drawing.SystemIcons.Application,
            Text = "EnglishSpiker",
            Visible = true
        };
        _trayIcon.MouseClick += (_, e) =>
        {
            if (e.Button == Forms.MouseButtons.Left)
            {
                ToggleWindow();
            }
            else if (e.Button == Forms.MouseButtons.Right)
            {
                ShowContextMenu();
            }
        };
    }

    private void ShowContextMenu()
    {
        var menu = new ContextMenu();
        menu.Items.Add(MenuItem(IsVisible ? "隐藏窗口" : "显示窗口", ToggleWindow));
        menu.Items.Add(CheckItem("复制后自动朗读", _autoRead, value => _autoRead = value));
        menu.Items.Add(CheckItem("自动显示中文释义", _autoLookup, value => _autoLookup = value));
        menu.Items.Add(CheckItem("窗口始终置顶", _alwaysOnTop, value =>
        {
            _alwaysOnTop = value;
            Topmost = value;
        }));
        menu.Items.Add(new Separator());

        var voiceMenu = new MenuItem { Header = "选择声音" };
        foreach (var voiceName in _speech.GetVoices())
        {
            var item = new MenuItem
            {
                Header = voiceName,
                IsCheckable = true,
                IsChecked = voiceName == _voice
            };
            item.Click += (_, _) => _voice = voiceName;
            voiceMenu.Items.Add(item);
        }
        menu.Items.Add(voiceMenu);

        var speedMenu = new MenuItem { Header = "语速" };
        speedMenu.Items.Add(RateItem("慢", -2));
        speedMenu.Items.Add(RateItem("正常", 0));
        speedMenu.Items.Add(RateItem("快", 2));
        menu.Items.Add(speedMenu);

        menu.Items.Add(new Separator());
        menu.Items.Add(MenuItem($"设置快捷键：{_hotkeyModifier}+{_hotkeyKey}", ShowHotkeyDialog));
        menu.Items.Add(new Separator());
        menu.Items.Add(MenuItem("退出程序", ExitApp));
        menu.IsOpen = true;
    }

    private MenuItem RateItem(string label, int value)
    {
        var item = new MenuItem { Header = label, IsCheckable = true, IsChecked = _rate == value };
        item.Click += (_, _) => _rate = value;
        return item;
    }

    private static MenuItem MenuItem(string label, Action action)
    {
        var item = new MenuItem { Header = label };
        item.Click += (_, _) => action();
        return item;
    }

    private static MenuItem CheckItem(string label, bool value, Action<bool> action)
    {
        var item = new MenuItem { Header = label, IsCheckable = true, IsChecked = value };
        item.Click += (_, _) => action(item.IsChecked);
        return item;
    }

    private void ShowHotkeyDialog()
    {
        var dialog = new Window
        {
            Title = "设置快捷键",
            Width = 280,
            Height = 130,
            ResizeMode = ResizeMode.NoResize,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            Topmost = true,
            Background = Brush("#f8f9fa")
        };
        var panel = new StackPanel { Margin = new Thickness(14) };
        panel.Children.Add(new TextBlock
        {
            Text = "重读剪贴板快捷键",
            FontWeight = FontWeights.Bold,
            Foreground = Brush("#2b3437"),
            Margin = new Thickness(0, 0, 0, 10)
        });

        var row = new StackPanel { Orientation = Orientation.Horizontal };
        var modifierBox = new ComboBox { Width = 130, ItemsSource = HotkeyModifiers.Keys.ToArray(), SelectedItem = _hotkeyModifier };
        var keyBox = new ComboBox { Width = 70, ItemsSource = HotkeyKeys, SelectedItem = _hotkeyKey, Margin = new Thickness(8, 0, 0, 0) };
        row.Children.Add(modifierBox);
        row.Children.Add(keyBox);
        panel.Children.Add(row);

        var apply = new Button { Content = "应用", Width = 72, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
        apply.Click += (_, _) =>
        {
            _hotkeyModifier = modifierBox.SelectedItem?.ToString() ?? "Ctrl+Alt";
            _hotkeyKey = keyBox.SelectedItem?.ToString() ?? "R";
            RegisterHotkey();
            dialog.Close();
        };
        panel.Children.Add(apply);
        dialog.Content = panel;
        dialog.ShowDialog();
    }

    private void PollClipboard()
    {
        if (!_autoRead)
        {
            return;
        }

        string text;
        try
        {
            text = Clipboard.GetText();
        }
        catch
        {
            return;
        }

        var word = NormalizeWord(text);
        if (string.IsNullOrWhiteSpace(word) || text == _lastClipboardText)
        {
            return;
        }

        _lastClipboardText = text;
        ProcessWord(word, speak: true, lookup: _autoLookup);
    }

    private void ReadClipboard()
    {
        try
        {
            var word = NormalizeWord(Clipboard.GetText());
            if (!string.IsNullOrWhiteSpace(word))
            {
                ProcessWord(word, speak: true, lookup: _autoLookup);
            }
        }
        catch
        {
            _statusText.Text = "Clipboard busy";
        }
    }

    private async void ProcessWord(string word, bool speak, bool lookup)
    {
        _wordBox.Text = word.Length > 28 ? word[..28] : word;
        _statusText.Text = "Reading...";
        if (speak)
        {
            _speech.Speak(word, _voice, _rate, () => Dispatcher.Invoke(() => _statusText.Text = "Listening..."));
        }
        if (lookup)
        {
            await LookupWord(word);
        }
    }

    private async Task LookupWord(string word)
    {
        if (word == _lastLookupWord)
        {
            return;
        }
        _lastLookupWord = word;
        _meaningText.Text = "查询中...";
        try
        {
            var result = await _dictionary.Lookup(word);
            _meaningText.Text = result.Meaning.Split('\n')[0].Trim();
        }
        catch
        {
            _meaningText.Text = "暂时没有查到中文释义";
        }
    }

    private void ToggleWindow()
    {
        if (IsVisible)
        {
            Hide();
        }
        else
        {
            Show();
            Activate();
            Topmost = _alwaysOnTop;
        }
    }

    private void RegisterHotkey()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        UnregisterHotKey(hwnd, HotkeyId);
        var modifiers = HotkeyModifiers.GetValueOrDefault(_hotkeyModifier, ModControl | ModAlt);
        var key = VirtualKeys.GetValueOrDefault(_hotkeyKey, (uint)'R');
        _statusText.Text = RegisterHotKey(hwnd, HotkeyId, modifiers, key) ? "Listening..." : "Hotkey used";
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmHotkey && wParam.ToInt32() == HotkeyId)
        {
            ReadClipboard();
            handled = true;
        }
        return IntPtr.Zero;
    }

    private void ExitApp()
    {
        _clipboardTimer.Stop();
        var hwnd = new WindowInteropHelper(this).Handle;
        UnregisterHotKey(hwnd, HotkeyId);
        _trayIcon?.Dispose();
        Application.Current.Shutdown();
    }

    private static string NormalizeWord(string text)
    {
        var match = Regex.Match(text.Trim(), @"[A-Za-z][A-Za-z' -]*");
        return match.Success ? Regex.Replace(match.Value.Trim(' ', '\'', '-'), @"\s+", " ").ToLowerInvariant() : "";
    }

    private static SolidColorBrush Brush(string hex) => new((Color)ColorConverter.ConvertFromString(hex));

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}

internal sealed record LookupResult(string Word, string Meaning, string Source);

internal sealed class DictionaryClient
{
    private static readonly HttpClient Http = new();
    private static readonly Dictionary<string, string> Fallback = new()
    {
        ["hello"] = "你好；喂",
        ["world"] = "世界",
        ["book"] = "书；书籍",
        ["read"] = "读；阅读",
        ["write"] = "写；书写",
        ["good"] = "好的；优秀的",
        ["bad"] = "坏的；糟糕的",
        ["happy"] = "快乐的；幸福的",
        ["study"] = "学习；研究",
        ["computer"] = "电脑；计算机",
        ["word"] = "单词；词语",
        ["apple"] = "苹果",
        ["water"] = "水",
        ["time"] = "时间；次数",
        ["love"] = "爱；喜爱",
    };

    public async Task<LookupResult> Lookup(string word)
    {
        var youdao = await TryLookupYoudao(word);
        if (!string.IsNullOrWhiteSpace(youdao))
        {
            return new LookupResult(word, youdao, "有道词典");
        }

        var translated = await TryLookupMyMemory(word);
        if (!string.IsNullOrWhiteSpace(translated))
        {
            return new LookupResult(word, translated, "MyMemory 翻译");
        }

        if (Fallback.TryGetValue(word.ToLowerInvariant(), out var fallback))
        {
            return new LookupResult(word, fallback, "本地小词库");
        }

        throw new InvalidOperationException("No meaning found.");
    }

    private static async Task<string> TryLookupYoudao(string word)
    {
        try
        {
            var url = $"https://dict.youdao.com/suggest?num=5&ver=3.0&doctype=json&cache=false&le=en&q={Uri.EscapeDataString(word)}";
            using var doc = JsonDocument.Parse(await Http.GetStringAsync(url));
            var entries = doc.RootElement.GetProperty("data").GetProperty("entries");
            var lines = new List<string>();
            foreach (var entry in entries.EnumerateArray())
            {
                if (entry.TryGetProperty("explain", out var explain))
                {
                    var value = explain.GetString();
                    if (!string.IsNullOrWhiteSpace(value) && !lines.Contains(value))
                    {
                        lines.Add(value);
                    }
                }
            }
            return string.Join('\n', lines);
        }
        catch
        {
            return "";
        }
    }

    private static async Task<string> TryLookupMyMemory(string word)
    {
        try
        {
            var url = $"https://api.mymemory.translated.net/get?q={Uri.EscapeDataString(word)}&langpair=en%7Czh-CN";
            using var doc = JsonDocument.Parse(await Http.GetStringAsync(url));
            var translated = doc.RootElement.GetProperty("responseData").GetProperty("translatedText").GetString();
            return string.Equals(translated, word, StringComparison.OrdinalIgnoreCase) ? "" : translated ?? "";
        }
        catch
        {
            return "";
        }
    }
}

internal sealed class SpeechService
{
    public string[] GetVoices()
    {
        try
        {
            var voice = CreateVoice();
            dynamic voices = voice.GetVoices();
            var names = new List<string> { "系统默认" };
            for (var index = 0; index < voices.Count; index++)
            {
                names.Add((string)voices.Item(index).GetDescription());
            }
            return names.Distinct().ToArray();
        }
        catch
        {
            return ["系统默认"];
        }
    }

    public void Speak(string word, string voiceName, int rate, Action onFinished)
    {
        var thread = new Thread(() =>
        {
            try
            {
                dynamic voice = CreateVoice();
                voice.Rate = rate;
                voice.Volume = 100;
                if (voiceName != "系统默认")
                {
                    dynamic voices = voice.GetVoices();
                    for (var index = 0; index < voices.Count; index++)
                    {
                        dynamic candidate = voices.Item(index);
                        if ((string)candidate.GetDescription() == voiceName)
                        {
                            voice.Voice = candidate;
                            break;
                        }
                    }
                }
                voice.Speak(word);
            }
            catch
            {
                // The UI keeps listening even when a particular voice fails.
            }
            finally
            {
                onFinished();
            }
        });
        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    private static dynamic CreateVoice()
    {
        var type = Type.GetTypeFromProgID("SAPI.SpVoice") ?? throw new InvalidOperationException("SAPI is unavailable.");
        return Activator.CreateInstance(type)!;
    }
}
