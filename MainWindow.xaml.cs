using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace NetworkDiagramBuilder
{
    // ====================================================
    // MODELS
    // ====================================================
    public class Device
    {
        public string Name  { get; set; } = "";
        public string Model { get; set; } = "";
        public string IP    { get; set; } = "";
        public string DisplayText =>
            Name
            + (string.IsNullOrEmpty(Model) ? "" : $"  [{Model}]")
            + (string.IsNullOrEmpty(IP)    ? "" : $"  {IP}");
    }

    public class ChildConn
    {
        public Device? Device   { get; set; }
        public bool    IsNone   { get; set; }   // 「接続なし（端末）」
        public bool    IsDouble { get; set; }   // 二重接続（既配置機器）
        public string  PortSrc  { get; set; } = "";  // 自側ポート（例: port1）
        public string  PortDst  { get; set; } = "";  // 相手側ポート（例: GE0/1）
    }

    public class ConnBlock
    {
        public int          Id            { get; set; }
        public Device?      Parent        { get; set; }
        public List<ChildConn> Children   { get; set; } = new();
        // WPF UI references
        public Border?      UIRoot        { get; set; }
        public ComboBox?    ParentCombo   { get; set; }
        public StackPanel?  ChildrenPanel { get; set; }
        public ComboBox?    ChildCombo    { get; set; }
        public Grid?        AddChildRow   { get; set; }
    }

    // ====================================================
    // MAIN WINDOW
    // ====================================================
    public partial class MainWindow : Window
    {
        private readonly List<Device>    _devices      = new();
        private string                   _memo         = "";
        private readonly List<ConnBlock> _blocks       = new();
        private readonly Dictionary<int, ConnBlock> _blockDict = new();  // O(1) Id lookup
        private int                      _blockId      = 0;
        private string                   _xml          = "";
        private Device?                  _editingDevice = null;   // 編集中の機器（null=追加モード）
        // O(1) device index lookup (rebuilt on parse/load)
        private readonly Dictionary<Device, int>   _deviceIndexMap = new();

        // Shared brushes（インスタンスフィールド: テーマ切替で .Color を書き換えるため static 不可）
        private readonly SolidColorBrush C_Bg       = MutableBrush("#0D1117");
        private readonly SolidColorBrush C_Surface  = MutableBrush("#161B22");
        private readonly SolidColorBrush C_Surface2 = MutableBrush("#21262D");
        private readonly SolidColorBrush C_Border   = MutableBrush("#30363D");
        private readonly SolidColorBrush C_Accent   = MutableBrush("#238636");
        private readonly SolidColorBrush C_Accent2  = MutableBrush("#1F6FEB");
        private readonly SolidColorBrush C_Text     = MutableBrush("#E6EDF3");
        private readonly SolidColorBrush C_Muted    = MutableBrush("#8B949E");
        private readonly SolidColorBrush C_Dim      = MutableBrush("#484F58");
        private readonly SolidColorBrush C_Danger   = MutableBrush("#DA3633");
        private readonly SolidColorBrush C_Green    = MutableBrush("#56D364");
        private readonly SolidColorBrush C_GreenBg  = MutableBrush("#0D2A0D");
        private readonly SolidColorBrush C_GreenBd  = MutableBrush("#238636");
        private readonly SolidColorBrush C_Amber    = MutableBrush("#C9922A");
        private readonly SolidColorBrush C_BlueBg   = MutableBrush("#0D1F35");
        private readonly SolidColorBrush C_NoneBg   = MutableBrush("#1A1A2E");
        private readonly SolidColorBrush C_NoneBd   = MutableBrush("#333344");

        // Cached FontFamily objects (creating new FontFamily() is expensive)
        private static readonly FontFamily FF_Consolas = new FontFamily("Consolas");

        // Freeze しない SolidColorBrush を生成（テーマ切替で .Color を後から書き換えるため）
        private static SolidColorBrush MutableBrush(string hex)
            => new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));

        // 後方互換エイリアス
        private static SolidColorBrush Brush(string hex) => MutableBrush(hex);

        // ── バージョン情報 ──
        private const string AppVersion    = "1.0.0";
        private const string GitHubOwner   = "your-github-username";   // ← 要変更
        private const string GitHubRepo    = "NetworkDiagramBuilder";  // ← 要変更
        private const string GitHubApiUrl  = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";
        private const string GitHubRelUrl  = $"https://github.com/{GitHubOwner}/{GitHubRepo}/releases/latest";
        private const string DrawioDownUrl = "https://github.com/jgraph/drawio-desktop/releases";

        private string _latestGitHubTag = "";

        // ====================================================
        // THEME
        // ====================================================
        private bool _isDark = false;   // デフォルト: ライトモード
        private int  _currentStep = 1;

        // (dark, light) のカラーペア定義
        private static readonly (string key, string dark, string light)[] ThemeColors = {
            ("BgBrush",          "#0D1117", "#FFFFFF"),
            ("SurfaceBrush",     "#161B22", "#F6F8FA"),
            ("Surface2Brush",    "#21262D", "#EFF1F3"),
            ("BorderBrush",      "#30363D", "#D0D7DE"),
            ("TextBrush",        "#E6EDF3", "#1F2328"),
            ("TextMuted",        "#8B949E", "#656D76"),
            ("TextDim",          "#484F58", "#8C959F"),
            ("ComboHighBrush",   "#2D333B", "#E2E5E9"),
            ("AmberBrush",       "#C9922A", "#9A6700"),
            ("AmberBorderBrush", "#9E6A03", "#C9922A"),
            ("AmberHoverBrush",  "#2D1A00", "#FFF8E6"),
        };

        private void ThemeToggle_Click(object s, MouseButtonEventArgs e)
        {
            _isDark = !_isDark;
            ApplyTheme();
        }

        private void ApplyTheme()
        {
            // ── XAML DynamicResource ブラシを新インスタンスで差し替え ──
            // XAML で生成されたブラシは WPF が Freeze するため .Color 変更不可。
            // Resources[] に新しい SolidColorBrush を代入することで DynamicResource が追従する。
            foreach (var (key, dark, light) in ThemeColors)
                Resources[key] = MutableBrush(_isDark ? dark : light);

            // ── C_* インスタンスブラシ（コードで生成したコントロール用）──
            // MutableBrush で作成済みのため .Color 変更が可能。
            // Color を変えるとそのブラシを参照している全コントロールが自動再描画される。
            void SetC(SolidColorBrush b, string dark, string light) =>
                b.Color = (Color)ColorConverter.ConvertFromString(_isDark ? dark : light);

            SetC(C_Bg,       "#0D1117", "#FFFFFF");
            SetC(C_Surface,  "#161B22", "#F6F8FA");
            SetC(C_Surface2, "#21262D", "#EFF1F3");
            SetC(C_Border,   "#30363D", "#D0D7DE");
            SetC(C_Text,     "#E6EDF3", "#1F2328");
            SetC(C_Muted,    "#8B949E", "#656D76");
            SetC(C_Dim,      "#484F58", "#8C959F");
            SetC(C_Green,    "#56D364", "#1A7F37");
            SetC(C_GreenBg,  "#0D2A0D", "#DAFBE1");
            SetC(C_GreenBd,  "#238636", "#2DA44E");
            SetC(C_Amber,    "#C9922A", "#9A6700");
            SetC(C_BlueBg,   "#0D1F35", "#DDF4FF");
            SetC(C_NoneBg,   "#1A1A2E", "#F6F8FA");
            SetC(C_NoneBd,   "#333344", "#D0D7DE");

            // ── テーマアイコン更新 ──
            ThemeToggleIcon.Text = _isDark ? "☀" : "🌙";

            // ── ステップインジケーター再描画（コードで色を設定しているため）──
            GoToStep(_currentStep);

            // ── ドロップゾーンのホバー色を現在テーマに合わせてリセット ──
            DropZone.Background  = C_Surface;
            DropZone.BorderBrush = C_Border;
        }

        public MainWindow()
        {
            InitializeComponent();
            ApplyTheme();           // 起動時にライトモードを適用
            _ = CheckForUpdateAsync();
        }

        // ====================================================
        // UPDATE CHECK
        // ====================================================
        private async System.Threading.Tasks.Task CheckForUpdateAsync()
        {
            try
            {
                using var http = new HttpClient();
                http.DefaultRequestHeaders.Add("User-Agent", "NetworkDiagramBuilder");
                http.Timeout = TimeSpan.FromSeconds(8);

                var json = await http.GetStringAsync(GitHubApiUrl);

                // "tag_name": "v1.2.3" を正規表現で抽出（外部ライブラリ不要）
                var m = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"");
                if (!m.Success) return;

                _latestGitHubTag = m.Groups[1].Value.TrimStart('v');

                if (IsNewerVersion(_latestGitHubTag, AppVersion))
                {
                    // UI スレッドで更新バナーを表示
                    await Dispatcher.InvokeAsync(() =>
                    {
                        UpdateBannerText.Text =
                            $"🎉 新しいバージョン v{_latestGitHubTag} が公開されています！（現在: v{AppVersion}）";
                        UpdateBanner.Visibility = Visibility.Visible;
                    });
                }
            }
            catch
            {
                // 更新確認失敗は無視（オフライン・ネットワーク不可など）
            }
        }

        /// <summary>latest が current より新しいバージョンか判定（セマンティックバージョニング）</summary>
        private static bool IsNewerVersion(string latest, string current)
        {
            try
            {
                var l = latest.Split('.').Select(int.Parse).ToArray();
                var c = current.Split('.').Select(int.Parse).ToArray();
                int len = Math.Max(l.Length, c.Length);
                for (int i = 0; i < len; i++)
                {
                    int lv = i < l.Length ? l[i] : 0;
                    int cv = i < c.Length ? c[i] : 0;
                    if (lv > cv) return true;
                    if (lv < cv) return false;
                }
                return false;
            }
            catch { return false; }
        }

        private void UpdateLink_Click(object s, MouseButtonEventArgs e)
        {
            OpenUrl(GitHubRelUrl);
        }

        private void UpdateBannerClose_Click(object s, MouseButtonEventArgs e)
        {
            UpdateBanner.Visibility = Visibility.Collapsed;
        }

        private void DrawioLink_Click(object s, MouseButtonEventArgs e)
        {
            OpenUrl(DrawioDownUrl);
        }

        private static void OpenUrl(string url)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName        = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ブラウザを開けませんでした。\n{url}\n\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ====================================================
        // DRAG-DROP
        // ====================================================
        private void DropZone_DragOver(object s, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy : DragDropEffects.None;
            DropZone.BorderBrush = C_Accent2;
            DropZone.Background  = C_BlueBg;
            e.Handled = true;
        }

        private void DropZone_DragLeave(object s, DragEventArgs e)
        {
            DropZone.BorderBrush = C_Border;
            DropZone.Background  = C_Surface;
        }

        private void DropZone_Drop(object s, DragEventArgs e)
        {
            DropZone.BorderBrush = C_Border;
            DropZone.Background  = C_Surface;
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
                LoadFile(files[0]);
        }

        private void DropZone_Click(object s, MouseButtonEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "テキストファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*",
                Title  = "メモファイルを選択"
            };
            if (dlg.ShowDialog() == true) LoadFile(dlg.FileName);
        }

        private void LoadFile(string path)
        {
            try   { PasteArea.Text = File.ReadAllText(path, Encoding.UTF8); }
            catch (Exception ex)
            { MessageBox.Show($"読み込み失敗:\n{ex.Message}", "エラー",
                MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        // ====================================================
        // PARSE
        // ====================================================
        private void ParseButton_Click(object s, RoutedEventArgs e) => DoParse();

        private void DoParse()
        {
            try
            {
                string text = PasteArea.Text.Trim();
                if (string.IsNullOrEmpty(text))
                {
                    MessageBox.Show("テキストを入力してください。", "確認",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _devices.Clear();
                _deviceIndexMap.Clear();
                _memo = "";

                // 【構成機器】～【構成以上】
                var devMatch = Regex.Match(text, @"【構成機器】([\s\S]*?)【構成以上】");
                if (devMatch.Success)
                {
                    foreach (var line in devMatch.Groups[1].Value.Split('\n'))
                    {
                        var t = line.Trim().Replace("\r", "");
                        if (string.IsNullOrEmpty(t)) continue;
                        var parts = Regex.Split(t, @"[\t　 ]+")
                                         .Where(p => !string.IsNullOrEmpty(p)).ToArray();
                        var dev = new Device
                        {
                            Name  = parts.ElementAtOrDefault(0) ?? "",
                            Model = parts.ElementAtOrDefault(1) ?? "",
                            IP    = parts.ElementAtOrDefault(2) ?? ""
                        };
                        _deviceIndexMap[dev] = _devices.Count;
                        _devices.Add(dev);
                    }
                }

                // 【以下メモ欄】～【メモ欄以上】
                var memoMatch = Regex.Match(text, @"【以下メモ欄】([\s\S]*?)【メモ欄以上】");
                if (memoMatch.Success) _memo = memoMatch.Groups[1].Value.Trim();

                if (_devices.Count == 0)
                {
                    MessageBox.Show(
                        "構成機器が見つかりませんでした。\n\n" +
                        "フォーマット確認:\n" +
                        "【構成機器】\n機器名　型番　IPアドレス\n【構成以上】",
                        "解析エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _blocks.Clear();
                _blockDict.Clear();
                _blockId = 0;
                BlocksPanel.Children.Clear();
                CloseAddDevicePanel();  // BUG-D: 解析時に編集フォームをリセット

                UpdateMemoPanel();
                UpdateDeviceChips();
                GoToStep(2);

                // AddBlock は GoToStep(2) でパネルが表示された後に呼ぶ
                // BUG-X1: 解析ボタン連続クリックで InvokeAsync が複数キューに積まれると
                // AddBlock が複数回実行され初期ブロックが2件以上追加される。
                // ボタンを一時無効化して InvokeAsync 完了後に再有効化することで防止する。
                ParseButton.IsEnabled = false;
                Dispatcher.InvokeAsync(() =>
                {
                    AddBlock();
                    ParseButton.IsEnabled = true;
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"解析中にエラーが発生しました:\n\n{ex.GetType().Name}\n{ex.Message}\n\n{ex.StackTrace}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ====================================================
        // MEMO & DEVICE CHIPS
        // ====================================================
        private void UpdateMemoPanel()
        {
            // MemoDisplay は編集可能な TextBox
            // TextChanged イベントで _memo に同期するため、ここでは一方向に Text をセットする
            // イベントの再帰呼び出しを防ぐため一時的にハンドラを外す
            MemoDisplay.TextChanged -= MemoDisplay_TextChanged;
            MemoDisplay.Text        = _memo;
            MemoDisplay.Foreground  = C_Muted;
            MemoDisplay.FontStyle   = FontStyles.Normal;
            MemoDisplay.TextChanged += MemoDisplay_TextChanged;
        }

        private void MemoDisplay_TextChanged(object s, TextChangedEventArgs e)
        {
            _memo = MemoDisplay.Text;
        }

        private void UpdateDeviceChips()
        {
            DeviceChips.Children.Clear();
            var asChild  = ChildDevices();
            var asParent = ParentDevices();
            // 親も子も「配置済み」として同じ緑で表示
            var placed = new HashSet<Device>(asChild);
            placed.UnionWith(asParent);

            foreach (var d in _devices)
            {
                bool isPlaced = placed.Contains(d);

                var (bg, bd, fg) = isPlaced
                    ? (C_GreenBg, C_GreenBd, C_Green)
                    : (C_Surface2, C_Border, C_Text);

                var chip = new Border
                {
                    Background      = bg,
                    BorderBrush     = bd,
                    BorderThickness = new Thickness(1),
                    CornerRadius    = new CornerRadius(6),
                    Padding         = new Thickness(8, 5, 8, 5),
                    Margin          = new Thickness(0, 0, 0, 4)
                };

                var g = new StackPanel();

                var nameT  = new TextBlock { Text = d.Name,  Foreground = fg, FontWeight = FontWeights.SemiBold, FontSize = 11 };
                var modelT = new TextBlock { Text = d.Model, Foreground = C_Muted, FontSize = 9, FontFamily = FF_Consolas };
                var ipT    = new TextBlock { Text = d.IP,    Foreground = C_Muted, FontSize = 9, FontFamily = FF_Consolas };
                g.Children.Add(nameT);
                if (!string.IsNullOrEmpty(d.Model)) g.Children.Add(modelT);
                if (!string.IsNullOrEmpty(d.IP))    g.Children.Add(ipT);

                if (isPlaced)
                {
                    var tagBd = new Border
                    {
                        Background   = bd,
                        CornerRadius = new CornerRadius(8),
                        Margin       = new Thickness(0, 3, 0, 0),
                        Padding      = new Thickness(5, 1, 5, 1),
                        HorizontalAlignment = HorizontalAlignment.Left
                    };
                    tagBd.Child = new TextBlock
                    {
                        Text       = "配置済",
                        Foreground = C_Bg,
                        FontSize   = 9,
                        FontWeight = FontWeights.Bold
                    };
                    g.Children.Add(tagBd);
                }

                chip.Child = g;

                // ── 右上に編集ボタンを重ねる ──
                // BUG-2: chip の Margin(0,0,0,4) は chipGrid 内の余白にしかならないため
                //        chipGrid 自体に Margin を設定してチップ間隔を確保する
                var chipGrid = new Grid { Margin = new Thickness(0, 0, 0, 4) };
                chipGrid.Children.Add(chip);

                var editBtn = new Button
                {
                    Content             = "✎",
                    Background          = Brushes.Transparent,
                    BorderThickness     = new Thickness(0),
                    Foreground          = C_Dim,
                    FontSize            = 11,
                    Cursor              = Cursors.Hand,
                    Padding             = new Thickness(4, 2, 4, 2),
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment   = VerticalAlignment.Top,
                    ToolTip             = "この機器を編集",
                    Focusable           = false,    // UX-1: Tab フォーカスが余分に回らないように
                };
                var capturedDevice = d;
                editBtn.Click      += (_, _) => OpenEditDevice(capturedDevice);
                editBtn.MouseEnter += (_, _) => editBtn.Foreground = C_Accent2;
                editBtn.MouseLeave += (_, _) => editBtn.Foreground = C_Dim;
                chipGrid.Children.Add(editBtn);

                // ドラッグ開始
                chip.Cursor = Cursors.Hand;
                chip.MouseMove += (sender, me) =>
                {
                    if (me.LeftButton == MouseButtonState.Pressed)
                    {
                        var data = new DataObject("Device", capturedDevice);
                        DragDrop.DoDragDrop(chip, data, DragDropEffects.Copy);
                    }
                };
                // ドラッグ中のビジュアルフィードバック
                chip.MouseEnter += (_, _) =>
                {
                    if (!placed.Contains(capturedDevice))
                        chip.BorderBrush = C_Accent2;
                };
                chip.MouseLeave += (_, _) => chip.BorderBrush = bd;

                DeviceChips.Children.Add(chipGrid);
            }
        }

        private HashSet<Device> ChildDevices()
        {
            var s = new HashSet<Device>();
            foreach (var b in _blocks)
                foreach (var c in b.Children)
                    if (c.Device != null) s.Add(c.Device);
            return s;
        }

        private HashSet<Device> ParentDevices()
        {
            var s = new HashSet<Device>();
            foreach (var b in _blocks)
                if (b.Parent != null) s.Add(b.Parent);
            return s;
        }

        // ====================================================
        // NAVIGATION
        // ====================================================
        private void GoToStep(int step)
        {
            _currentStep = step;
            PanelStep1.Visibility = step == 1 ? Visibility.Visible : Visibility.Collapsed;
            PanelStep2.Visibility = step == 2 ? Visibility.Visible : Visibility.Collapsed;
            PanelStep3.Visibility = step == 3 ? Visibility.Visible : Visibility.Collapsed;

            SetStepStyle(Si1, Si1Text, step, 1);
            SetStepStyle(Si2, Si2Text, step, 2);
            SetStepStyle(Si3, Si3Text, step, 3);
        }

        private void SetStepStyle(Border bd, TextBlock tb, int cur, int me)
        {
            if (cur > me)
            {
                bd.BorderBrush = C_Accent;  bd.Background = C_GreenBg; tb.Foreground = C_Accent;
            }
            else if (cur == me)
            {
                bd.BorderBrush = C_Accent2; bd.Background = C_BlueBg;  tb.Foreground = C_Accent2;
            }
            else
            {
                bd.BorderBrush = C_Border;  bd.Background = Brushes.Transparent; tb.Foreground = C_Dim;
            }
        }

        private void BackToStep1_Click(object s, RoutedEventArgs e) => GoToStep(1);
        private void BackToStep2_Click(object s, RoutedEventArgs e) => GoToStep(2);

        // ====================================================
        // CONNECTION BLOCKS
        // ====================================================
        private void AddBlock_Click(object s, RoutedEventArgs e)
        {
            try { AddBlock(); }
            catch (Exception ex)
            {
                MessageBox.Show($"ブロック追加中にエラー:\n{ex.Message}", "エラー",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // BlocksPanelの末尾に「＋ 追加」ボタンを配置する
        private void AppendAddButton()
        {
            // 既存の追加ボタンを削除
            var existing = BlocksPanel.Children
                .OfType<Border>()
                .FirstOrDefault(b => b.Tag?.ToString() == "__add_btn__");
            if (existing != null) BlocksPanel.Children.Remove(existing);

            var wrapper = new Border
            {
                Tag    = "__add_btn__",
                Margin = new Thickness(0, 4, 0, 4)
            };
            var btn = MakeButton("＋  接続ブロックを追加", BtnKind.Secondary);
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Click += AddBlock_Click;
            wrapper.Child = btn;
            BlocksPanel.Children.Add(wrapper);
        }

        private void AddBlock()
        {
            var block = new ConnBlock { Id = ++_blockId };
            RegisterBlock(block);
            AppendAddButton();
        }

        // ブロックを登録してUIを構築（インポート時など外部から呼ぶ共通処理）
        private void RegisterBlock(ConnBlock block)
        {
            if (block.Id == 0) block.Id = ++_blockId;
            _blocks.Add(block);
            _blockDict[block.Id] = block;
            BuildBlockUI(block);
            FillParentCombo(block);
        }

        private void BuildBlockUI(ConnBlock block)
        {
            // ── Root border ──
            var root = new Border
            {
                Background    = C_Surface,
                BorderBrush   = C_Border,
                BorderThickness = new Thickness(1),
                CornerRadius  = new CornerRadius(10),
                Margin        = new Thickness(0, 0, 0, 12)
            };
            block.UIRoot = root;

            var outer = new StackPanel();

            // ── Header ──
            var hdr = new Border
            {
                Background    = C_Surface2,
                BorderBrush   = C_Border,
                BorderThickness = new Thickness(0, 0, 0, 1),
                CornerRadius  = new CornerRadius(9, 9, 0, 0),
                Padding       = new Thickness(14, 10, 14, 10)
            };
            var hg = new Grid();
            hg.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            hg.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            hg.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var numBd = new Border
            {
                Width = 24, Height = 24, CornerRadius = new CornerRadius(12),
                Background = C_Accent2, Margin = new Thickness(0, 0, 10, 0)
            };
            numBd.Child = new TextBlock
            {
                Text = block.Id.ToString(), Foreground = Brushes.White,
                FontSize = 10, FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center
            };
            var lbl = new TextBlock
            {
                Text = $"接続ブロック #{block.Id}", Foreground = C_Muted,
                FontSize = 12, VerticalAlignment = VerticalAlignment.Center
            };
            var delBtn = MakeButton("✕  削除", BtnKind.Danger);
            delBtn.Tag    = block.Id;
            delBtn.Click += DeleteBlock_Click;

            Grid.SetColumn(numBd,  0);
            Grid.SetColumn(lbl,    1);
            Grid.SetColumn(delBtn, 2);
            hg.Children.Add(numBd);
            hg.Children.Add(lbl);
            hg.Children.Add(delBtn);
            hdr.Child = hg;

            // ── Body ──
            var body = new StackPanel
            {
                Margin    = new Thickness(16, 12, 16, 12),
                AllowDrop = true
            };

            // ドロップゾーンのイベント
            body.DragOver += (_, de) =>
            {
                de.Effects = de.Data.GetDataPresent("Device")
                    ? DragDropEffects.Copy : DragDropEffects.None;
                de.Handled = true;
            };
            body.DragEnter += (_, _) =>
            {
                root.BorderBrush = C_Accent2;
                root.Background  = C_BlueBg;  // BUG-B: #0D1F2D はダーク専用 → テーマ対応ブラシに変更
            };
            body.DragLeave += (_, _) =>
            {
                // 状態に応じた枠色に戻す
                root.BorderBrush = block.Children.Any(c => !c.IsNone && c.Device != null) ? C_Accent
                                 : block.Parent != null ? C_Accent2
                                 : C_Border;
                root.Background  = C_Surface;
            };
            body.Drop += (_, de) =>
            {
                root.Background  = C_Surface;
                if (de.Data.GetData("Device") is not Device dropped)
                {
                    // ドロップデータが Device でない場合は現在の状態に応じた枠色に戻す
                    root.BorderBrush = block.Children.Any(c => !c.IsNone && c.Device != null) ? C_Accent
                                     : block.Parent != null ? C_Accent2
                                     : C_Border;
                    return;
                }

                // 親が未設定ならドロップした機器を親にセット
                if (block.Parent == null)
                {
                    // ParentCombo から該当アイテムを選択（ParentCombo_Changedが枠色を更新する）
                    foreach (ComboBoxItem ci2 in block.ParentCombo!.Items)
                        if (ci2.Tag == dropped) { block.ParentCombo.SelectedItem = ci2; break; }
                }
                else
                {
                    // 親設定済み → 接続先として追加
                    bool isDouble = ChildDevices().Contains(dropped) ||
                                   ParentDevices().Contains(dropped);
                    // 同一ブロック内で重複しないようチェック
                    if (block.Children.Any(c => c.Device == dropped))
                    {
                        root.BorderBrush = C_Accent;
                        return;
                    }
                    block.Children.Add(new ChildConn { Device = dropped, IsDouble = isDouble });
                    RenderChildren(block);
                    RefreshChildCombosAndChips();  // ChildDevices()は1回のみ計算
                    root.BorderBrush = C_Accent;
                }
                de.Handled = true;
            };

            // Parent row
            var pRow = new Grid { Margin = new Thickness(0, 0, 0, 10) };
            pRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(72) });
            pRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var pLbl = new TextBlock
            {
                Text = "親機器:", Foreground = C_Muted, FontSize = 11,
                FontWeight = FontWeights.Bold, VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(pLbl, 0);
            pRow.Children.Add(pLbl);
            var pCombo = MakeCombo();
            pCombo.Tag              = block.Id;
            pCombo.SelectionChanged += ParentCombo_Changed;
            block.ParentCombo       = pCombo;
            Grid.SetColumn(pCombo, 1);
            pRow.Children.Add(pCombo);
            body.Children.Add(pRow);

            // Children list panel
            var cpanel = new StackPanel();
            block.ChildrenPanel = cpanel;
            body.Children.Add(cpanel);

            // Add-child row (hidden until parent selected)
            var acRow = new Grid { Visibility = Visibility.Collapsed, Margin = new Thickness(0, 8, 0, 0) };
            acRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(72) });
            acRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            block.AddChildRow = acRow;

            var acLbl = new TextBlock
            {
                Text = "接続先:", Foreground = C_Muted, FontSize = 11,
                FontWeight = FontWeights.Bold, VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(acLbl, 0);

            var cCombo = MakeCombo();
            cCombo.Tag = block.Id;
            block.ChildCombo = cCombo;
            // 選択したら即追加（ボタン不要）
            cCombo.SelectionChanged += ChildCombo_Changed;
            Grid.SetColumn(cCombo, 1);

            acRow.Children.Add(acLbl);
            acRow.Children.Add(cCombo);
            body.Children.Add(acRow);

            outer.Children.Add(hdr);
            outer.Children.Add(body);
            root.Child = outer;

            BlocksPanel.Children.Add(root);
        }


        // ====================================================
        // COMBO POPULATION
        // ====================================================
        // 単体呼び出し用
        private void FillParentCombo(ConnBlock block) => FillParentComboCore(block, ParentDevices());

        // 一括更新用（ParentDevices を外から渡してO(n²)を防ぐ）
        private void FillParentComboCore(ConnBlock block, HashSet<Device> allParents)
        {
            var combo   = block.ParentCombo!;
            var usedAsP = new HashSet<Device>(allParents);
            if (block.Parent != null) usedAsP.Remove(block.Parent); // 自分の親は表示させる

            combo.SelectionChanged -= ParentCombo_Changed;
            combo.Items.Clear();
            combo.Items.Add(MakeComboItem("-- 選択してください --", null, C_Dim));

            foreach (var d in _devices)
            {
                if (usedAsP.Contains(d)) continue;
                combo.Items.Add(MakeComboItem(d.DisplayText, d, C_Text));
            }

            // Restore selection
            if (block.Parent != null)
                foreach (ComboBoxItem ci in combo.Items)
                    if (ci.Tag == block.Parent) { combo.SelectedItem = ci; break; }
            else
                combo.SelectedIndex = 0;

            combo.SelectionChanged += ParentCombo_Changed;
        }

        // 単体呼び出し用（ChildDevices・ParentDevices を内部で計算）
        private void FillChildCombo(ConnBlock block) =>
            FillChildComboCore(block, ChildDevices(), ParentDevices());

        // 一括更新用（両セットを外から渡してO(n²)を防ぐ）
        private void FillChildComboCore(ConnBlock block,
            HashSet<Device> allChildDevices, HashSet<Device> allParentDevices)
        {
            var combo = block.ChildCombo;
            if (combo == null) return;  // BuildBlockUI前に呼ばれた場合の安全ガード;
            // 他ブロックの親 = 全親セット から 自ブロックの親を除いたもの（O(1)）
            var otherParents = new HashSet<Device>(allParentDevices);
            if (block.Parent != null) otherParents.Remove(block.Parent);

            // 自ブロックの既存の子機器（既に追加済み→コンボに出さない）
            var thisBlockChildren = new HashSet<Device>(
                block.Children.Where(c => c.Device != null).Select(c => c.Device!));

            combo.Items.Clear();
            combo.Items.Add(MakeComboItem("-- 選択してください --", null,       C_Dim));
            // 接続なし（端末）は既に追加済みなら選択肢から非表示
            if (!block.Children.Any(c => c.IsNone))
                combo.Items.Add(MakeComboItem("── 接続なし（端末）", "__none__", C_Muted, FontStyles.Italic));

            // 配置済み = 他ブロックで子として使用済み OR 他ブロックで親として使用済み
            var otherBlockChildren = new HashSet<Device>(allChildDevices.Except(thisBlockChildren));
            var placedSet = new HashSet<Device>(otherBlockChildren);
            placedSet.UnionWith(otherParents);
            // 自ブロックの子・自ブロックの親 は表示対象外
            var displayable = _devices.Where(d => !thisBlockChildren.Contains(d) && d != block.Parent).ToList();
            var available   = displayable.Where(d => !placedSet.Contains(d)).ToList();
            var placed      = displayable.Where(d =>  placedSet.Contains(d)).ToList();

            foreach (var d in available)
                combo.Items.Add(MakeComboItem(d.DisplayText, d, C_Text));

            if (placed.Count > 0)
            {
                combo.Items.Add(MakeComboItem("── 配置済み機器（二重接続用）──", null, C_Dim, isHeader: true));
                foreach (var d in placed)
                {
                    // _deviceIndexMap で O(1) 取得（同名機器の衝突回避）
                    int devIdx = _deviceIndexMap.TryGetValue(d, out var idx) ? idx : _devices.IndexOf(d);
                    combo.Items.Add(MakeComboItem("[配置済] " + d.DisplayText, $"__used__{devIdx}", C_Amber));
                }
            }

            combo.SelectedIndex = 0;
        }

        // ====================================================
        // EVENT HANDLERS
        // ====================================================
        private void ParentCombo_Changed(object s, SelectionChangedEventArgs e)
        {
            if (s is not ComboBox combo) return;
            if (combo.Tag == null) return;
            var block = GetBlock((int)combo.Tag);
            if (block == null) return;

            var ci = combo.SelectedItem as ComboBoxItem;
            block.Parent = ci?.Tag as Device;
            block.Children.Clear();
            RenderChildren(block);

            if (block.Parent != null)
            {
                block.AddChildRow!.Visibility = Visibility.Visible;
                block.UIRoot!.BorderBrush = C_Accent2;
            }
            else
            {
                block.AddChildRow!.Visibility = Visibility.Collapsed;
                block.UIRoot!.BorderBrush     = C_Border;
            }

            // 親変更で「配置済み」状態が変わるため、全ブロックの子コンボを更新（UpdateDeviceChipsも内包）
            RefreshChildCombosAndChips();
            RefreshAllParentCombos(block.Id);
        }

        // 接続先コンボ：選択したら即追加
        private void ChildCombo_Changed(object s, SelectionChangedEventArgs e)
        {
            if (s is not ComboBox combo) return;
            if (combo.Tag == null) return;
            var block = GetBlock((int)combo.Tag);
            if (block == null) return;

            var ci = combo.SelectedItem as ComboBoxItem;
            if (ci == null || combo.SelectedIndex == 0) return;
            // ヘッダー行は無視
            if (ci.IsEnabled == false) { combo.SelectedIndex = 0; return; }

            string? tag = ci.Tag?.ToString();

            if (tag == "__none__")
            {
                // 接続なし（端末）は1つだけ追加可能
                if (!block.Children.Any(c => c.IsNone))
                    block.Children.Add(new ChildConn { IsNone = true });
            }
            else if (tag?.StartsWith("__used__") == true)
            {
                // インデックスでデバイスを特定（同名機器の衝突回避）
                if (int.TryParse(tag.Substring("__used__".Length), out int devIdx)
                    && devIdx >= 0 && devIdx < _devices.Count)
                {
                    var dev = _devices[devIdx];
                    // 同一ブロック内に既に存在しない場合のみ追加
                    if (!block.Children.Any(c => c.Device == dev))
                        block.Children.Add(new ChildConn { Device = dev, IsDouble = true });
                }
            }
            else if (ci.Tag is Device d)
            {
                // 同一ブロック内に既に存在しない場合のみ追加
                if (!block.Children.Any(c => c.Device == d))
                    block.Children.Add(new ChildConn { Device = d });
            }

            // 追加後はコンボをリセット（続けて別の機器を選択できるように）
            combo.SelectionChanged -= ChildCombo_Changed;
            combo.SelectedIndex = 0;
            combo.SelectionChanged += ChildCombo_Changed;

            RenderChildren(block);
            RefreshChildCombosAndChips();  // ChildDevices()は1回のみ計算

            // 実際の機器接続がある場合のみ緑（接続なし端末のみは青のまま）
            bool hasRealConn = block.Children.Any(c => !c.IsNone && c.Device != null);
            block.UIRoot!.BorderBrush = hasRealConn ? C_Accent : C_Accent2;
        }

        private void DeleteBlock_Click(object s, RoutedEventArgs e)
        {
            if (s is not Button btn) return;
            if (btn.Tag is not int tagId) return;   // null や非int を安全に除外
            var block = GetBlock(tagId);
            if (block == null) return;

            _blocks.Remove(block);
            _blockDict.Remove(block.Id);
            BlocksPanel.Children.Remove(block.UIRoot);

            RefreshAllCombosAndChips();
            AppendAddButton();
        }

        // ====================================================
        // RENDER CHILDREN
        // ====================================================
        private void RenderChildren(ConnBlock block)
        {
            var panel = block.ChildrenPanel!;
            panel.Children.Clear();

            for (int i = 0; i < block.Children.Count; i++)
            {
                var ch  = block.Children[i];
                var row = new Grid { Margin = new Thickness(0, 0, 0, 5) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var arrow = new TextBlock
                {
                    Text = "└─", Foreground = C_Dim,
                    FontFamily = FF_Consolas,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(arrow, 0);

                Border badge;
                if (ch.IsNone)
                {
                    badge = new Border
                    {
                        Background = C_NoneBg, BorderBrush = C_NoneBd,
                        BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6),
                        Padding = new Thickness(10, 5, 10, 5)
                    };
                    badge.Child = new TextBlock
                    {
                        Text = "接続なし（端末）", Foreground = C_Dim,
                        FontSize = 11, FontStyle = FontStyles.Italic
                    };
                }
                else
                {
                    badge = new Border
                    {
                        Background = C_Surface2, BorderBrush = C_Border,
                        BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6),
                        Padding = new Thickness(10, 5, 10, 5)
                    };
                    var sp = new StackPanel { Orientation = Orientation.Horizontal };
                    sp.Children.Add(new TextBlock
                    {
                        Text = ch.Device!.Name, Foreground = C_Text,
                        FontSize = 11, FontWeight = FontWeights.SemiBold,
                        FontFamily = FF_Consolas
                    });
                    if (!string.IsNullOrEmpty(ch.Device.Model))
                        sp.Children.Add(new TextBlock { Text = "  " + ch.Device.Model, Foreground = C_Muted, FontSize = 10, FontFamily = FF_Consolas, VerticalAlignment = VerticalAlignment.Center });
                    if (!string.IsNullOrEmpty(ch.Device.IP))
                        sp.Children.Add(new TextBlock { Text = "  " + ch.Device.IP, Foreground = C_Accent2, FontSize = 10, FontFamily = FF_Consolas, VerticalAlignment = VerticalAlignment.Center });
                    if (ch.IsDouble)
                        sp.Children.Add(new TextBlock { Text = "  [二重接続]", Foreground = C_Amber, FontSize = 9, VerticalAlignment = VerticalAlignment.Center });
                    badge.Child = sp;
                }
                Grid.SetColumn(badge, 1);

                int ci = i, bid = block.Id;
                var rmBtn = new Button
                {
                    Content = "✕", Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0), Foreground = C_Dim,
                    FontSize = 13, Cursor = Cursors.Hand,
                    Padding = new Thickness(4, 0, 4, 0), Margin = new Thickness(4, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center
                };
                rmBtn.Click      += (_, _) => RemoveChild(bid, ci);
                rmBtn.MouseEnter += (_, _) => rmBtn.Foreground = C_Danger;
                rmBtn.MouseLeave += (_, _) => rmBtn.Foreground = C_Dim;
                Grid.SetColumn(rmBtn, 2);

                row.Children.Add(arrow);
                row.Children.Add(badge);
                row.Children.Add(rmBtn);
                panel.Children.Add(row);

                // ポート番号入力 アコーディオン
                if (!ch.IsNone)
                {
                    var capturedCh = ch;

                    // 入力パネル（初期は折りたたみ、既入力があれば展開）
                    bool hasPort = !string.IsNullOrEmpty(capturedCh.PortSrc) || !string.IsNullOrEmpty(capturedCh.PortDst);

                    var portPanel = new StackPanel
                    {
                        Visibility = hasPort ? Visibility.Visible : Visibility.Collapsed,
                        Margin = new Thickness(0, 2, 0, 4)
                    };

                    // トグルボタン
                    var toggleBtn = new Button
                    {
                        Background = Brushes.Transparent,
                        BorderThickness = new Thickness(0),
                        Padding = new Thickness(0, 2, 0, 2),
                        Cursor = Cursors.Hand,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        HorizontalContentAlignment = HorizontalAlignment.Left,
                    };

                    // トグルボタンの内容（テキスト＋矢印）
                    var toggleSp = new StackPanel { Orientation = Orientation.Horizontal };
                    var arrowTb = new TextBlock
                    {
                        Text = hasPort ? "▼" : "▶",
                        Foreground = C_Muted, FontSize = 9,
                        Margin = new Thickness(0, 0, 4, 0),
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    var toggleLbl = new TextBlock
                    {
                        Text = "ポート番号入力",
                        Foreground = C_Muted, FontSize = 10,
                        FontFamily = FF_Consolas,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    toggleSp.Children.Add(arrowTb);
                    toggleSp.Children.Add(toggleLbl);
                    toggleBtn.Content = toggleSp;
                    toggleBtn.MouseEnter += (_, _) => { arrowTb.Foreground = C_Text; toggleLbl.Foreground = C_Text; };
                    toggleBtn.MouseLeave += (_, _) => { arrowTb.Foreground = C_Muted; toggleLbl.Foreground = C_Muted; };
                    toggleBtn.Click += (_, _) =>
                    {
                        bool open = portPanel.Visibility == Visibility.Visible;
                        portPanel.Visibility = open ? Visibility.Collapsed : Visibility.Visible;
                        arrowTb.Text = open ? "▶" : "▼";
                    };
                    panel.Children.Add(toggleBtn);

                    // 接続元ポート行
                    var portRow1 = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                    portRow1.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(86) });
                    portRow1.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    var portLbl1 = new TextBlock
                    {
                        Text = "  接続元:", Foreground = C_Dim, FontSize = 10,
                        FontFamily = FF_Consolas, VerticalAlignment = VerticalAlignment.Center
                    };
                    Grid.SetColumn(portLbl1, 0);
                    var portBox1 = new TextBox
                    {
                        Text = capturedCh.PortSrc,
                        Background = C_Surface2, Foreground = C_Text,
                        BorderBrush = C_Border, BorderThickness = new Thickness(1),
                        FontFamily = FF_Consolas, FontSize = 10,
                        Padding = new Thickness(6, 3, 6, 3), CaretBrush = C_Text,
                        ToolTip = "接続元（自側）のポート番号（例: port1 / GE1/0/1）"
                    };
                    portBox1.TextChanged += (_, _) => capturedCh.PortSrc = portBox1.Text;
                    Grid.SetColumn(portBox1, 1);
                    portRow1.Children.Add(portLbl1);
                    portRow1.Children.Add(portBox1);
                    portPanel.Children.Add(portRow1);

                    // 接続先ポート行
                    var portRow2 = new Grid { Margin = new Thickness(0, 0, 0, 2) };
                    portRow2.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(86) });
                    portRow2.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    var portLbl2 = new TextBlock
                    {
                        Text = "  接続先:", Foreground = C_Dim, FontSize = 10,
                        FontFamily = FF_Consolas, VerticalAlignment = VerticalAlignment.Center
                    };
                    Grid.SetColumn(portLbl2, 0);
                    var portBox2 = new TextBox
                    {
                        Text = capturedCh.PortDst,
                        Background = C_Surface2, Foreground = C_Text,
                        BorderBrush = C_Border, BorderThickness = new Thickness(1),
                        FontFamily = FF_Consolas, FontSize = 10,
                        Padding = new Thickness(6, 3, 6, 3), CaretBrush = C_Text,
                        ToolTip = "接続先（相手側）のポート番号（例: GE1/0/1 / ethernet1/1）"
                    };
                    portBox2.TextChanged += (_, _) => capturedCh.PortDst = portBox2.Text;
                    Grid.SetColumn(portBox2, 1);
                    portRow2.Children.Add(portLbl2);
                    portRow2.Children.Add(portBox2);
                    portPanel.Children.Add(portRow2);

                    panel.Children.Add(portPanel);
                }
            }
        }

        private void RemoveChild(int blockId, int idx)
        {
            var block = GetBlock(blockId);
            if (block == null) return;
            if (idx < 0 || idx >= block.Children.Count) return;  // 二重クリック等による境界外アクセス防止
            block.Children.RemoveAt(idx);
            RenderChildren(block);
            RefreshChildCombosAndChips();  // ChildDevices()は1回のみ計算
            // 枠色を接続状態に応じて更新（機器削除後もIsNoneのみ残る可能性があるため）
            if (block.Parent != null)
            {
                bool hasRealConn = block.Children.Any(c => !c.IsNone && c.Device != null);
                block.UIRoot!.BorderBrush = hasRealConn ? C_Accent : C_Accent2;
            }
        }

        // ====================================================
        // REFRESH COMBOS
        // ====================================================
        private void RefreshAllParentCombos(int exceptId)
        {
            var parentSet = ParentDevices();   // 1回だけ計算
            foreach (var b in _blocks)
                if (b.Id != exceptId) FillParentComboCore(b, parentSet);
        }

        /// <summary>子コンボ全ブロック＋チップを一括更新（両セット計算は各1回）</summary>
        private void RefreshChildCombosAndChips()
        {
            var childSet  = ChildDevices();
            var parentSet = ParentDevices();
            foreach (var b in _blocks)
                if (b.Parent != null) FillChildComboCore(b, childSet, parentSet);
            UpdateDeviceChips();
        }

        /// <summary>親コンボ・子コンボ全ブロック＋チップを一括更新（両セット計算は各1回）</summary>
        private void RefreshAllCombosAndChips()
        {
            var childSet  = ChildDevices();
            var parentSet = ParentDevices();
            foreach (var b in _blocks)
            {
                FillParentComboCore(b, parentSet);
                if (b.Parent != null) FillChildComboCore(b, childSet, parentSet);
            }
            UpdateDeviceChips();
        }

        // ====================================================
        // CONFIG SAVE / LOAD
        // ====================================================
        private void SaveConfig_Click(object s, RoutedEventArgs e)
        {
            if (_devices.Count == 0)
            {
                MessageBox.Show("保存するデータがありません。先に機器情報を解析してください。",
                    "確認", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter   = "構成図設定ファイル (*.ndb.json)|*.ndb.json|JSON ファイル (*.json)|*.json",
                FileName = SanitizeFileName(DiagramNameBox.Text.Trim(), "network_diagram") + ".ndb.json",
                Title    = "設定を保存"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                // シリアライズ用の匿名オブジェクトを構築
                var data = new
                {
                    version     = "1.0",
                    diagramName = DiagramNameBox.Text.Trim(),
                    memo        = _memo,
                    pasteText   = PasteArea.Text,
                    devices     = _devices.Select(d => new { d.Name, d.Model, d.IP }).ToArray(),
                    blocks      = _blocks
                        .Where(b => b.Parent != null)
                        .Select(b => new
                        {
                            parentIndex = _deviceIndexMap.TryGetValue(b.Parent!, out var pi) ? pi : -1,
                            children    = b.Children.Select(ch => new
                            {
                                deviceIndex = ch.Device != null && _deviceIndexMap.TryGetValue(ch.Device, out var ci) ? ci : -1,
                                ch.IsNone,
                                ch.IsDouble,
                                ch.PortSrc,
                                ch.PortDst
                            }).ToArray()
                        }).ToArray()
                };

                var json = JsonSerializer.Serialize(data,
                    new JsonSerializerOptions { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
                File.WriteAllText(dlg.FileName, json, new System.Text.UTF8Encoding(false));

                MessageBox.Show($"設定を保存しました。\n{Path.GetFileName(dlg.FileName)}",
                    "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存に失敗しました。\n\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadConfig_Click(object s, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "構成図設定ファイル (*.ndb.json)|*.ndb.json|JSON ファイル (*.json)|*.json",
                Title  = "設定を読み込む"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var json = File.ReadAllText(dlg.FileName, System.Text.Encoding.UTF8);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                // ── 機器リスト復元 ──
                _devices.Clear();
                _deviceIndexMap.Clear();
                if (root.TryGetProperty("devices", out var devsEl))
                {
                    foreach (var dEl in devsEl.EnumerateArray())
                    {
                        var loadedDev = new Device
                        {
                            Name  = dEl.TryGetProperty("Name",  out var n) ? n.GetString() ?? "" : "",
                            Model = dEl.TryGetProperty("Model", out var m) ? m.GetString() ?? "" : "",
                            IP    = dEl.TryGetProperty("IP",    out var ip) ? ip.GetString() ?? "" : ""
                        };
                        _deviceIndexMap[loadedDev] = _devices.Count;
                        _devices.Add(loadedDev);
                    }
                }

                if (_devices.Count == 0)
                {
                    MessageBox.Show("機器情報が見つかりませんでした。ファイルが破損している可能性があります。",
                        "読み込みエラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // ── メモ・テキスト復元 ──
                _memo = root.TryGetProperty("memo", out var memoEl) ? memoEl.GetString() ?? "" : "";
                if (root.TryGetProperty("pasteText", out var pasteEl))
                    PasteArea.Text = pasteEl.GetString() ?? "";
                if (root.TryGetProperty("diagramName", out var nameEl))
                    DiagramNameBox.Text = nameEl.GetString() ?? "";

                // ── UI リセット ──
                _blocks.Clear();
                _blockDict.Clear();
                _blockId = 0;
                BlocksPanel.Children.Clear();
                CloseAddDevicePanel();  // BUG-A/C: 読込時に編集フォームと _editingDevice をリセット
                UpdateMemoPanel();
                UpdateDeviceChips();
                GoToStep(2);

                // ── ブロック復元 ──
                if (root.TryGetProperty("blocks", out var blocksEl))
                {
                    foreach (var bEl in blocksEl.EnumerateArray())
                    {
                        int pIdx = bEl.TryGetProperty("parentIndex", out var piEl) ? piEl.GetInt32() : -1;
                        if (pIdx < 0 || pIdx >= _devices.Count) continue;

                        var block = new ConnBlock { Id = ++_blockId };
                        _blocks.Add(block);
                        _blockDict[block.Id] = block;
                        BuildBlockUI(block);

                        // 親を設定（FillParentCombo内で SelectedItem も復元される、イベント発火なし）
                        block.Parent = _devices[pIdx];
                        FillParentCombo(block);
                        block.AddChildRow!.Visibility = Visibility.Visible;
                        block.UIRoot!.BorderBrush = C_Accent2;

                        // 子を復元
                        if (bEl.TryGetProperty("children", out var childrenEl))
                        {
                            foreach (var chEl in childrenEl.EnumerateArray())
                            {
                                bool isNone   = chEl.TryGetProperty("IsNone",   out var inEl) && inEl.GetBoolean();
                                bool isDouble = chEl.TryGetProperty("IsDouble", out var idEl) && idEl.GetBoolean();
                                int  dIdx     = chEl.TryGetProperty("deviceIndex", out var diEl) ? diEl.GetInt32() : -1;
                                string pSrc   = chEl.TryGetProperty("PortSrc", out var psEl) ? psEl.GetString() ?? "" : "";
                                string pDst   = chEl.TryGetProperty("PortDst", out var pdEl) ? pdEl.GetString() ?? "" : "";

                                if (isNone)
                                {
                                    block.Children.Add(new ChildConn { IsNone = true });
                                }
                                else if (dIdx >= 0 && dIdx < _devices.Count)
                                {
                                    block.Children.Add(new ChildConn
                                    {
                                        Device   = _devices[dIdx],
                                        IsDouble = isDouble,
                                        PortSrc  = pSrc,
                                        PortDst  = pDst
                                    });
                                }
                            }
                        }

                        RenderChildren(block);
                        // 子コンボは末尾の RefreshAllCombosAndChips で一括更新

                        // 枠色
                        bool hasReal = block.Children.Any(c => !c.IsNone && c.Device != null);
                        block.UIRoot!.BorderBrush = hasReal ? C_Accent : C_Accent2;
                    }
                }

                // 全ブロック読み込み後に親・子コンボを一括更新
                RefreshAllCombosAndChips();
                AppendAddButton();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"読み込みに失敗しました。\n\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ====================================================
        // IMPORT FROM DRAWIO
        // ====================================================
        private void ImportDrawio_Click(object s, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "draw.io ファイル (*.drawio;*.xml)|*.drawio;*.xml|すべてのファイル (*.*)|*.*",
                Title  = "draw.io ファイルから構成を復元"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var xml = File.ReadAllText(dlg.FileName, System.Text.Encoding.UTF8);
                ImportFromDrawioXml(xml);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"draw.io ファイルの読み込みに失敗しました。\n\n{ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ImportFromDrawioXml(string xmlText)
        {
            // ── draw.io XML パース ──
            var doc   = System.Xml.Linq.XDocument.Parse(xmlText);
            var cells = doc.Descendants("mxCell").ToList();

            var nodeById  = new Dictionary<string, string>();
            var edgeCells = new List<System.Xml.Linq.XElement>();
            var textCells = new List<System.Xml.Linq.XElement>();
            var nodeGeom  = new Dictionary<string, (double x, double y, double w, double h)>();
            string memoExtracted = "";  // draw.io メモセルから復元したメモ本文

            // draw.io の diagram name を取得
            string drawioName = doc.Descendants("diagram").FirstOrDefault()?.Attribute("name")?.Value ?? "";

            foreach (var cell in cells)
            {
                string id       = cell.Attribute("id")?.Value ?? "";
                string value    = cell.Attribute("value")?.Value ?? "";
                string style    = cell.Attribute("style")?.Value ?? "";
                bool   isVertex = cell.Attribute("vertex")?.Value == "1";
                bool   isEdge   = cell.Attribute("edge")?.Value   == "1";
                string? src     = cell.Attribute("source")?.Value;
                string? tgt     = cell.Attribute("target")?.Value;

                if (id == "0" || id == "1") continue;

                if (isEdge && src != null && tgt != null)
                {
                    edgeCells.Add(cell);
                }
                else if (isVertex)
                {
                    bool isLabel = cell.Attribute("connectable")?.Value == "0"
                                || style.Contains("text;")
                                || style.Contains("edgeLabel");

                    var geomEl = cell.Element("mxGeometry");
                    double gx = 0, gy = 0, gw = 0, gh = 0;
                    if (geomEl != null)
                    {
                        double.TryParse(geomEl.Attribute("x")?.Value,      System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out gx);
                        double.TryParse(geomEl.Attribute("y")?.Value,      System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out gy);
                        double.TryParse(geomEl.Attribute("width")?.Value,  System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out gw);
                        double.TryParse(geomEl.Attribute("height")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out gh);
                    }

                    if (isLabel)
                    {
                        // メモ相当の大きなテキストセルはポートラベル候補から除外
                        // （style="text;..." でも幅200超 or 高さ100超は装飾メモとみなす）
                        // ただし本ツール出力のメモセル（黄色付箋スタイル）は _memo として復元する
                        if (gw > 200 || gh > 100)
                        {
                            // 本ツール出力のメモセル判定: fillColor=#fffacd かつ value が「メモ」ヘッダで始まる
                            if (style.Contains("fillColor=#fffacd") && style.Contains("strokeColor=#d6b656"))
                            {
                                // value 例: "&lt;b&gt;メモ&lt;/b&gt;&lt;br&gt;内容..."
                                // StripHtml でデコードすると "メモ\n内容..."
                                string memoRaw = StripHtml(value);
                                // 先頭の "メモ" ヘッダ行を除去して本文だけ取り出す
                                int nlIdx = memoRaw.IndexOf('\n');
                                memoExtracted = nlIdx >= 0 ? memoRaw.Substring(nlIdx + 1).Trim() : memoRaw.Trim();
                            }
                            continue;
                        }
                        textCells.Add(cell);
                    }
                    else if (!string.IsNullOrWhiteSpace(id))
                    {
                        string clean = StripHtml(value);
                        // メモセル除外: 幅350px超かつ行数6行超
                        if (gw > 350 && clean.Count(c => c == '\n') > 5) continue;
                        nodeById[id]  = clean;
                        nodeGeom[id]  = (gx, gy, gw, gh);
                    }
                }
            }

            if (nodeById.Count == 0)
            {
                MessageBox.Show("ノード（機器）が見つかりませんでした。\ndraw.io ファイルの形式を確認してください。",
                    "復元エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ── Device 生成 ──
            var idToDevice = new Dictionary<string, Device>();
            foreach (var (id, label) in nodeById)
            {
                var parts = label.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0).ToArray();
                string name  = parts.Length > 0 ? parts[0] : id;
                string model = parts.Length > 1 ? parts[1] : "";
                string ip    = parts.FirstOrDefault(l => Regex.IsMatch(l, @"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")) ?? "";
                if (model == ip) model = "";
                idToDevice[id] = new Device { Name = name, Model = model, IP = ip };
            }

            // ── ポート情報収集 ──
            var edgePortSrc = new Dictionary<string, string>();
            var edgePortDst = new Dictionary<string, string>();
            // スコア辞書（絶対座標ラベルの競合解決用）
            var edgePortSrcScore = new Dictionary<string, double>();
            var edgePortDstScore = new Dictionary<string, double>();

            // BUG修正: edgeCells.First()はO(n²)+例外リスク → Dictionary<id, XElement>に変更
            var edgeById = edgeCells
                .Where(ec => (ec.Attribute("id")?.Value ?? "").Length > 0)
                .ToDictionary(ec => ec.Attribute("id")!.Value);

            foreach (var tc in textCells)
            {
                string parentId = tc.Attribute("parent")?.Value ?? "";
                string tcVal    = StripHtml(tc.Attribute("value")?.Value ?? "").Trim();
                if (string.IsNullOrWhiteSpace(tcVal)) continue;

                var geomEl = tc.Element("mxGeometry");

                if (edgeById.TryGetValue(parentId, out var parentEdge))
                {
                    // edgeLabel (connectable=0, parent=エッジID)
                    string eSrc2 = parentEdge.Attribute("source")?.Value ?? "";
                    string eTgt2 = parentEdge.Attribute("target")?.Value ?? "";
                    string key   = $"{eSrc2}→{eTgt2}";

                    // x値 (relative=-1〜1): x<0=src寄り, x=0=中央(スキップ), x>0=dst寄り
                    // draw.io 仕様: x=0 はエッジ中点の中央ラベルなので src/dst に割り当てない
                    string? xAttr = geomEl?.Attribute("x")?.Value;
                    if (xAttr == null) continue; // x属性なし=中央ラベルはスキップ
                    double.TryParse(xAttr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double gx);
                    if (gx == 0.0) continue; // x=0 は中央ラベル (IMP-05: 旧 gx<0.5 では src 誤割当)
                    if (gx < 0.0)
                        edgePortSrc.TryAdd(key, tcVal);
                    else
                        edgePortDst.TryAdd(key, tcVal);
                }
                else if (parentId == "1")
                {
                    // 絶対座標テキストセル（本ツール出力形式 + 一般 draw.io）
                    if (geomEl == null) continue;
                    double.TryParse(geomEl.Attribute("x")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double lx);
                    double.TryParse(geomEl.Attribute("y")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double ly);
                    double.TryParse(geomEl.Attribute("width")?.Value,  System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double lw);
                    double.TryParse(geomEl.Attribute("height")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double lh);

                    // ──────────────────────────────────────────────────────────────
                    // 本ツールの出力ルール（GenerateXML参照）:
                    //   PortSrc: lx = node.x + exitX * NodeW - 30    ly = node.y + NodeH + 4
                    //   PortDst: lx = node.x + entryX * NodeW - 30   ly = node.y - 20
                    //
                    // 逆算:
                    //   exitX  = (lx + 30 - node.x) / NodeW
                    //   entryX = (lx + 30 - node.x) / NodeW
                    //
                    // この逆算値をエッジの exitX/entryX と比較して最も近いエッジに割り当てる。
                    // 一般 draw.io ファイルのラベルも同様にスコアリングしてフォールバック。
                    // ──────────────────────────────────────────────────────────────
                    const double NodeW_f = 160.0;
                    const double SrcYOffset = 4.0;   // PortSrc: node_bottom + 4
                    const double DstYOffset = -20.0; // PortDst: node_top - 20

                    // (スコア, エッジsrc, エッジtgt, isSrc)
                    double bestScore = double.MaxValue;
                    string bestEdgeSrc = "";
                    string bestEdgeTgt = "";
                    bool   bestIsSrc   = false;

                    foreach (var ec in edgeCells)
                    {
                        string eSrc2 = ec.Attribute("source")?.Value ?? "";
                        string eTgt2 = ec.Attribute("target")?.Value ?? "";
                        if (!nodeGeom.ContainsKey(eSrc2) || !nodeGeom.ContainsKey(eTgt2)) continue;

                        // エッジの exitX / entryX を取得
                        string eStyle = ec.Attribute("style")?.Value ?? "";
                        double exitX  = ParseStyleDouble(eStyle, "exitX",  0.5);
                        double entryX = ParseStyleDouble(eStyle, "entryX", 0.5);

                        var (sx, sy, sw, sh) = nodeGeom[eSrc2];
                        var (dx, dy, dw, dh) = nodeGeom[eTgt2];

                        // --- Src 側スコア ---
                        // 本ツール出力の PortSrc Y位置: sy + sh + SrcYOffset
                        // 一般配置のフォールバック: sy + sh (ノード下辺) までの Y 距離
                        double srcNodeW = sw > 0 ? sw : NodeW_f;
                        double recoveredExitX = (lx + 30.0 - sx) / srcNodeW;  // 本ツール逆算
                        double dxNorm_src     = Math.Abs(recoveredExitX - exitX);
                        double srcExpectedY   = sy + sh + SrcYOffset;
                        double dyScore_src    = Math.Abs(ly - srcExpectedY);

                        // 本ツール形式に近い場合（Y誤差 ≤ 8px）: 正規化X差のみで判定（高精度）
                        // 一般形式フォールバック（Y誤差 > 8px）: Y距離を重く加算
                        double score_src = dyScore_src <= 8.0
                            ? dxNorm_src * 1000.0 + dyScore_src
                            : dxNorm_src * 50.0   + dyScore_src * 5.0;

                        // Y が src ノード側にない場合は大ペナルティ
                        if (ly < sy + sh - 10 || ly > sy + sh + 60) score_src += 9999.0;

                        // --- Dst 側スコア ---
                        double dstNodeW = dw > 0 ? dw : NodeW_f;
                        double recoveredEntryX = (lx + 30.0 - dx) / dstNodeW;
                        double dxNorm_dst      = Math.Abs(recoveredEntryX - entryX);
                        double dstExpectedY    = dy + DstYOffset;
                        double dyScore_dst     = Math.Abs(ly - dstExpectedY);

                        double score_dst = dyScore_dst <= 8.0
                            ? dxNorm_dst * 1000.0 + dyScore_dst
                            : dxNorm_dst * 50.0   + dyScore_dst * 5.0;

                        if (ly > dy + 10 || ly < dy - 60) score_dst += 9999.0;

                        // 良い方を採用
                        if (score_src < bestScore)
                        {
                            bestScore = score_src; bestEdgeSrc = eSrc2; bestEdgeTgt = eTgt2; bestIsSrc = true;
                        }
                        if (score_dst < bestScore)
                        {
                            bestScore = score_dst; bestEdgeSrc = eSrc2; bestEdgeTgt = eTgt2; bestIsSrc = false;
                        }
                    }

                    if (bestScore < 9000.0 && bestEdgeSrc.Length > 0)
                    {
                        string key = $"{bestEdgeSrc}→{bestEdgeTgt}";
                        if (bestIsSrc)
                        {
                            // スコア比較: 既存より良いスコアなら上書き（TryAdd は最初のもの勝ちになるため使わない）
                            if (!edgePortSrcScore.TryGetValue(key, out double prevScore) || bestScore < prevScore)
                            {
                                edgePortSrc[key]      = tcVal;
                                edgePortSrcScore[key] = bestScore;
                            }
                        }
                        else
                        {
                            if (!edgePortDstScore.TryGetValue(key, out double prevScore) || bestScore < prevScore)
                            {
                                edgePortDst[key]      = tcVal;
                                edgePortDstScore[key] = bestScore;
                            }
                        }
                    }
                }
            }

            // ── リセット ──
            _devices.Clear();
            _deviceIndexMap.Clear();
            _blocks.Clear();
            _blockDict.Clear();
            _blockId = 0;          // BUG-31: リセットがなく#2から始まっていた
            _memo = memoExtracted;  // draw.io メモセルから復元（なければ空文字）
            BlocksPanel.Children.Clear();
            CloseAddDevicePanel();  // BUG-B/C: インポート時に編集フォームと _editingDevice をリセット
            UpdateMemoPanel();

            foreach (var dev in idToDevice.Values)
            {
                _deviceIndexMap[dev] = _devices.Count;
                _devices.Add(dev);
            }

            // ── ConnBlock 構築（全エッジを処理してブロックを作成）──
            var parentToBlock = new Dictionary<string, ConnBlock>();
            // Y座標でソート：まず全ブロックを作成する
            var sortedEdges = edgeCells
                .Where(ec => nodeById.ContainsKey(ec.Attribute("source")?.Value ?? "")
                          && nodeById.ContainsKey(ec.Attribute("target")?.Value ?? ""))
                .OrderBy(ec => nodeGeom.TryGetValue(ec.Attribute("source")?.Value ?? "", out var g) ? g.y : 0)
                .ToList();

            foreach (var ec in sortedEdges)
            {
                string eSrc = ec.Attribute("source")?.Value ?? "";
                string eTgt = ec.Attribute("target")?.Value ?? "";
                string key  = $"{eSrc}→{eTgt}";

                if (!idToDevice.TryGetValue(eSrc, out var parentDev)) continue;
                if (!idToDevice.TryGetValue(eTgt, out var childDev))  continue;

                if (!parentToBlock.TryGetValue(eSrc, out var block))
                {
                    block = new ConnBlock { Parent = parentDev };
                    parentToBlock[eSrc] = block;
                    RegisterBlock(block);
                    // BUG-32: RegisterBlock/BuildBlockUI は AddChildRow を Collapsed で生成する
                    // 親が確定しているので Visible に設定して子の追加を可能にする
                    if (block.AddChildRow != null)
                        block.AddChildRow.Visibility = Visibility.Visible;
                }

                edgePortSrc.TryGetValue(key, out string? pSrc);
                edgePortDst.TryGetValue(key, out string? pDst);
                // BUG-E: draw.io で同一ノード間に複数エッジがある場合の重複追加を防止
                if (block.Children.Any(c => c.Device == childDev)) continue;
                block.Children.Add(new ChildConn { Device = childDev, PortSrc = pSrc ?? "", PortDst = pDst ?? "" });
            }

            // ── IsDouble フラグを設定 ──
            // 正しい定義: 同じデバイスが「複数の異なるブロックの子」として登場している場合のみ IsDouble=true
            // NG例: 「他ブロックの親」は通常のツリー接続であり二重接続ではない
            var childAppearCount = new Dictionary<Device, int>();
            foreach (var block in _blocks)
                foreach (var ch in block.Children.Where(c => c.Device != null))
                {
                    childAppearCount.TryGetValue(ch.Device!, out int cnt);
                    childAppearCount[ch.Device!] = cnt + 1;
                }
            foreach (var block in _blocks)
                foreach (var ch in block.Children.Where(c => c.Device != null))
                    ch.IsDouble = childAppearCount.TryGetValue(ch.Device!, out int n) && n > 1;

            // ── ブロックをトポロジカルソートで並び替え ──
            // 「親が他のブロックの子に含まれていない」ブロック（ルート）を先に置く
            // → 上位ノードのブロックが上に表示される
            _blocks.Clear();
            _blockDict.Clear();
            BlocksPanel.Children.Clear();

            // id→Block マップを parentToBlock から再構築
            var allBlocks = parentToBlock.Values.ToList();

            // 各ブロックの親デバイスが「他ブロックの子」に含まれているか判定
            var blockChildDevs = new HashSet<Device>(
                allBlocks.SelectMany(b => b.Children).Where(c => c.Device != null).Select(c => c.Device!));

            // BFS トポロジカルソート
            // 入次数: そのブロックの親が他ブロックの子に含まれる数
            var inDegree = new Dictionary<ConnBlock, int>();
            var depGraph = new Dictionary<ConnBlock, List<ConnBlock>>();  // 親→依存するブロック
            foreach (var b in allBlocks)
            {
                inDegree[b]  = 0;
                depGraph[b]  = new List<ConnBlock>();
            }
            foreach (var b in allBlocks)
            {
                if (b.Parent == null) continue;
                // b の親が他ブロックの子であれば、その「他ブロック」が b の先行
                foreach (var other in allBlocks.Where(o => o != b))
                {
                    if (other.Children.Any(c => c.Device == b.Parent))
                    {
                        inDegree[b]++;
                        depGraph[other].Add(b);
                    }
                }
            }

            // 入次数0のブロック（ルート）をキューに、Y座標の小さい順で処理
            var queue = new Queue<ConnBlock>(
                allBlocks.Where(b => inDegree[b] == 0)
                         .OrderBy(b => b.Parent != null && nodeGeom.TryGetValue(
                             parentToBlock.FirstOrDefault(kv => kv.Value == b).Key, out var g) ? g.y : 0));

            var sorted = new List<ConnBlock>();
            while (queue.Count > 0)
            {
                var b = queue.Dequeue();
                sorted.Add(b);
                foreach (var dep in depGraph[b])
                {
                    inDegree[dep]--;
                    if (inDegree[dep] == 0)
                        queue.Enqueue(dep);
                }
            }
            // 循環参照が残っている場合（実際の draw.io グラフで起きうる）は Y座標順で追加
            foreach (var b in allBlocks.Where(b => !sorted.Contains(b))
                                       .OrderBy(b => b.Parent != null && nodeGeom.TryGetValue(
                                           parentToBlock.FirstOrDefault(kv => kv.Value == b).Key, out var g) ? g.y : 0))
                sorted.Add(b);

            // ソート済み順で _blocks / _blockDict / BlocksPanel を再構成
            // ※ ID は RegisterBlock で確定済み。IDを変えるとUI Tag(削除ボタン等)と不一致になるため変更しない
            _blocks.Clear();
            _blockDict.Clear();
            BlocksPanel.Children.Clear();

            foreach (var b in sorted)
            {
                _blocks.Add(b);
                _blockDict[b.Id] = b;
                if (b.UIRoot != null)
                    BlocksPanel.Children.Add(b.UIRoot);
            }

            // ── UI 更新 ──
            foreach (var block in _blocks)
            {
                RenderChildren(block);
                if (block.UIRoot != null)
                    block.UIRoot.BorderBrush = block.Children.Any(c => !c.IsNone && c.Device != null) ? C_Accent : C_Accent2;
            }

            RefreshAllCombosAndChips();
            AppendAddButton();

            // 図名を設定
            if (!string.IsNullOrWhiteSpace(drawioName))
                DiagramNameBox.Text = drawioName;

            GoToStep(2);

            MessageBox.Show(
                $"draw.io から復元しました。\n\n" +
                $"・機器数: {idToDevice.Count}\n" +
                $"・接続数: {sortedEdges.Count}\n" +
                $"・ポート情報: {edgePortSrc.Count + edgePortDst.Count} 件\n" +
                (string.IsNullOrEmpty(memoExtracted) ? "" : "・メモ: 復元済み\n") +
                "\n接続ブロックを確認・修正してください。",
                "復元完了", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>HTMLタグを除去してプレーンテキストに変換</summary>
        // draw.io スタイル文字列から "key=value;" の value を double で取得する
        // 例: "exitX=0.3333;exitY=1;..." → ParseStyleDouble(style, "exitX", 0.5) → 0.3333
        private static double ParseStyleDouble(string style, string key, double defaultVal)
        {
            int idx = style.IndexOf(key + "=", StringComparison.Ordinal);
            if (idx < 0) return defaultVal;
            // BUG-D: 語境界チェック。"noExitX=..." のような接頭辞付きキーへの誤マッチを防ぐ
            // key の直前が英数字またはアンダースコアの場合は別のキーの一部なので skip
            if (idx > 0 && (char.IsLetterOrDigit(style[idx - 1]) || style[idx - 1] == '_'))
            {
                // 次の出現を探す（接頭辞なしのものが後方にある場合）
                int next = style.IndexOf(key + "=", idx + 1, StringComparison.Ordinal);
                while (next >= 0)
                {
                    if (next == 0 || (!char.IsLetterOrDigit(style[next - 1]) && style[next - 1] != '_'))
                    { idx = next; break; }
                    next = style.IndexOf(key + "=", next + 1, StringComparison.Ordinal);
                }
                if (next < 0) return defaultVal;
            }
            int start = idx + key.Length + 1;
            int end   = style.IndexOf(';', start);
            string raw = end >= 0 ? style.Substring(start, end - start) : style.Substring(start);
            return double.TryParse(raw.Trim(), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double v) ? v : defaultVal;
        }

        private static string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html)) return "";
            // &lt; &gt; &amp; などをデコード
            html = html.Replace("&lt;",  "<")
                       .Replace("&gt;",  ">")
                       .Replace("&amp;", "&")
                       .Replace("&nbsp;"," ")
                       .Replace("<br>",  "\n")
                       .Replace("<br/>", "\n")
                       .Replace("<br />","\n");
            // <b>/<strong> などのタグを除去
            html = Regex.Replace(html, "<[^>]+>", "");
            return html.Trim();
        }

        // ====================================================
        // GENERATE XML
        // ====================================================
        private void GenerateXML_Click(object s, RoutedEventArgs e)
        {
            // 有効ブロック = 親あり かつ 実際の機器接続が1つ以上（接続なし端末のみは除外）
            var valid = _blocks.Where(b =>
                b.Parent != null &&
                b.Children.Any(c => !c.IsNone && c.Device != null)).ToList();

            if (valid.Count == 0)
            {
                MessageBox.Show("接続ブロックを少なくとも1つ完成させてください。\n（親機器 + 接続先機器を設定）",
                    "確認", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string diagName = DiagramNameBox.Text.Trim();
            if (string.IsNullOrEmpty(diagName)) diagName = "ネットワーク構成図";
            _xml = DrawioXml.Build(valid, diagName, _memo);
            RenderOutputPanel(valid);
            GoToStep(3);
        }

        private void RenderOutputPanel(List<ConnBlock> valid)
        {
            OutputPanel.Children.Clear();

            // ── Summary ──
            var sumBd = new Border
            {
                Background = C_Surface, BorderBrush = C_Border, BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10), Margin = new Thickness(0, 0, 0, 12)
            };
            var sumSp = new StackPanel();

            var sumHdr = new Border
            {
                Background = C_Surface2, BorderBrush = C_Border,
                BorderThickness = new Thickness(0, 0, 0, 1),
                CornerRadius = new CornerRadius(9, 9, 0, 0), Padding = new Thickness(16, 10, 16, 10)
            };
            var shg = new Grid();
            shg.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            shg.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            shg.Children.Add(new TextBlock { Text = "📊  接続関係サマリー", Foreground = C_Text, FontWeight = FontWeights.Bold, FontSize = 12 });
            var cntTb = new TextBlock { Text = $"{valid.Count} ブロック", Foreground = C_Muted, FontSize = 11, FontFamily = FF_Consolas };
            Grid.SetColumn(cntTb, 1);
            shg.Children.Add(cntTb);
            sumHdr.Child = shg;
            sumSp.Children.Add(sumHdr);

            var body = new StackPanel { Margin = new Thickness(16, 12, 16, 12) };
            foreach (var b in valid)
            {
                var pt = new TextBlock { FontFamily = FF_Consolas, FontSize = 12, Margin = new Thickness(0, 0, 0, 4) };
                pt.Inlines.Add(new Run("📦 " + b.Parent!.Name) { Foreground = C_Text, FontWeight = FontWeights.SemiBold });
                if (!string.IsNullOrEmpty(b.Parent.IP))
                    pt.Inlines.Add(new Run($" ({b.Parent.IP})") { Foreground = C_Muted });
                body.Children.Add(pt);

                foreach (var ch in b.Children)
                {
                    var ct = new TextBlock { FontFamily = FF_Consolas, FontSize = 11, Margin = new Thickness(20, 0, 0, 2) };
                    ct.Inlines.Add(new Run("└─ ") { Foreground = C_Dim });
                    if (ch.IsNone || ch.Device == null)
                    {
                        ct.Inlines.Add(new Run("[接続なし]") { Foreground = C_Dim, FontStyle = FontStyles.Italic });
                    }
                    else
                    {
                        ct.Inlines.Add(new Run(ch.Device.Name) { Foreground = C_Muted });
                        if (!string.IsNullOrEmpty(ch.Device.IP))
                            ct.Inlines.Add(new Run($" ({ch.Device.IP})") { Foreground = C_Accent2 });
                        if (ch.IsDouble)
                            ct.Inlines.Add(new Run(" [二重接続]") { Foreground = C_Amber });
                    }
                    body.Children.Add(ct);
                }
            }
            sumSp.Children.Add(body);
            sumBd.Child = sumSp;
            OutputPanel.Children.Add(sumBd);

            // ── XML preview ──
            var xmlBd = new Border
            {
                Background = C_Surface, BorderBrush = C_Accent, BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10)
            };
            var xmlSp = new StackPanel();
            var xmlHdr = new Border
            {
                Background = C_GreenBg, BorderBrush = C_GreenBd,
                BorderThickness = new Thickness(0, 0, 0, 1),
                CornerRadius = new CornerRadius(9, 9, 0, 0), Padding = new Thickness(16, 10, 16, 10)
            };
            xmlHdr.Child = new TextBlock { Text = "draw.io XML プレビュー", Foreground = C_Green, FontWeight = FontWeights.Bold, FontSize = 12 };
            xmlSp.Children.Add(xmlHdr);
            xmlSp.Children.Add(new TextBox
            {
                Text = _xml, Background = C_Surface, Foreground = C_Muted,
                BorderThickness = new Thickness(0), FontFamily = FF_Consolas,
                FontSize = 11, Padding = new Thickness(16), IsReadOnly = true,
                TextWrapping = TextWrapping.NoWrap,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
                MaxHeight = 260
            });
            xmlBd.Child = xmlSp;
            OutputPanel.Children.Add(xmlBd);
        }

        // ====================================================
        // OUTPUT
        // ====================================================
        private void CopyXML_Click(object s, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_xml))
            {
                Clipboard.SetText(_xml);
                MessageBox.Show("XMLをクリップボードにコピーしました！", "完了",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private static string SanitizeFileName(string name, string fallback)
        {
            if (string.IsNullOrEmpty(name)) return fallback;
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c.ToString(), "_");
            return name;
        }

        private void SaveXML_Click(object s, RoutedEventArgs e)
        {
            string diagName = SanitizeFileName(DiagramNameBox.Text.Trim(), "network_diagram");

            var dlg = new SaveFileDialog
            {
                Filter   = "draw.io ファイル (*.drawio)|*.drawio|XML ファイル (*.xml)|*.xml",
                FileName = diagName + ".drawio",
                Title    = "ファイルに保存"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                File.WriteAllText(dlg.FileName, _xml, new System.Text.UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ファイルの書き込みに失敗しました。\n\n{ex.Message}",
                    "保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 保存後の選択ダイアログ
            ShowAfterSaveDialog(dlg.FileName);
        }

        private void ShowAfterSaveDialog(string savedPath)
        {
            var win = new Window
            {
                Title                 = "保存完了",
                Width                 = 340,
                Height                = 190,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner                 = this,
                ResizeMode            = ResizeMode.NoResize,
                Background            = C_Surface,
            };

            // BUG-33: 別 Window は MainWindow の DynamicResource を解決できないため
            // BtnSecondary/BtnGhost のブラシが null になりボタンが見えなくなる。
            // MainWindow のリソースをダイアログ Window にコピーして解決する。
            foreach (System.Collections.DictionaryEntry de in Resources)
                win.Resources[de.Key] = de.Value;

            var sp = new StackPanel { Margin = new Thickness(24, 20, 24, 20) };

            sp.Children.Add(new TextBlock
            {
                Text       = "✓  保存しました",
                Foreground = C_Green,
                FontSize   = 14,
                FontWeight = FontWeights.Bold,
                Margin     = new Thickness(0, 0, 0, 6)
            });
            sp.Children.Add(new TextBlock
            {
                Text         = System.IO.Path.GetFileName(savedPath),
                Foreground   = C_Muted,
                FontSize     = 11,
                FontFamily   = FF_Consolas,
                TextWrapping = TextWrapping.Wrap,
                Margin       = new Thickness(0, 0, 0, 18)
            });

            var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

            // フォルダを開く
            var btnFolder = MakeButton("📁  フォルダを開く", BtnKind.Secondary);
            btnFolder.Margin = new Thickness(0, 0, 8, 0);
            btnFolder.Click += (_, _) =>
            {
                System.Diagnostics.Process.Start("explorer.exe",
                    $"/select,\"{savedPath}\"");
                win.Close();
            };

            // ファイルを開く
            var btnOpen = MakeButton("📄  ファイルを開く", BtnKind.Secondary);
            btnOpen.Margin = new Thickness(0, 0, 8, 0);
            btnOpen.Click += (_, _) =>
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName        = savedPath,
                        UseShellExecute = true
                    });
                }
                catch
                {
                    MessageBox.Show("ファイルを開けませんでした。\ndraw.io がインストールされているか確認してください。",
                        "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                win.Close();
            };

            // 閉じる
            var btnClose = MakeButton("✕  閉じる", BtnKind.Ghost);
            btnClose.Click += (_, _) => win.Close();

            btnRow.Children.Add(btnFolder);
            btnRow.Children.Add(btnOpen);
            btnRow.Children.Add(btnClose);
            sp.Children.Add(btnRow);

            win.Content = sp;
            win.ShowDialog();
        }

        // ====================================================
        // ADD / EDIT DEVICE
        // ====================================================

        /// <summary>追加/編集フォームを閉じて状態をクリア（全リセット箇所から呼ぶ共通処理）</summary>
        private void CloseAddDevicePanel()
        {
            AddDevicePanel.Visibility   = Visibility.Collapsed;
            _editingDevice              = null;
            AddDevPanelTitle.Text       = "機器を追加";
            ConfirmAddDeviceBtn.Content = "✓ 追加";
            AddDevNameBox.Text          = "";
            AddDevModelBox.Text         = "";
            AddDevIPBox.Text            = "";
        }

        private void ToggleAddDevice_Click(object s, RoutedEventArgs e)
        {
            bool opening = AddDevicePanel.Visibility != Visibility.Visible;
            if (opening)
            {
                // 追加モードで開く
                CloseAddDevicePanel();          // フォームをクリア（編集残留リセット）
                AddDevicePanel.Visibility = Visibility.Visible;
                AddDevNameBox.Focus();
            }
            else
            {
                CloseAddDevicePanel();          // BUG-E: closing 時も _editingDevice をリセット
            }
        }

        // チップの ✎ ボタンから呼ばれる
        private void OpenEditDevice(Device dev)
        {
            _editingDevice              = dev;
            AddDevPanelTitle.Text       = $"機器を編集：{dev.Name}";
            ConfirmAddDeviceBtn.Content = "✓ 更新";
            AddDevNameBox.Text          = dev.Name;
            AddDevModelBox.Text         = dev.Model;
            AddDevIPBox.Text            = dev.IP;
            AddDevicePanel.Visibility   = Visibility.Visible;
            AddDevNameBox.Focus();
            AddDevNameBox.SelectAll();
        }

        private void CancelAddDevice_Click(object s, RoutedEventArgs e)
        {
            CloseAddDevicePanel();
        }

        private void ConfirmAddDevice_Click(object s, RoutedEventArgs e)
        {
            string name = AddDevNameBox.Text.Trim();
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("機器名を入力してください。", "確認",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                AddDevNameBox.Focus();
                return;
            }

            if (_editingDevice != null)
            {
                // ────────── 編集モード ──────────
                // BUG-1: 他の機器と同名への変更は警告する
                if (_editingDevice.Name != name && _devices.Any(d => d != _editingDevice && d.Name == name))
                {
                    var r = MessageBox.Show(
                        $"「{name}」という機器が既に存在します。\n同名で変更しますか？",
                        "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (r != MessageBoxResult.Yes) return;
                }
                // Device はクラス（参照型）なのでプロパティ変更で _blocks 内の参照も自動更新
                _editingDevice.Name  = name;
                _editingDevice.Model = AddDevModelBox.Text.Trim();
                _editingDevice.IP    = AddDevIPBox.Text.Trim();
                _editingDevice = null;
            }
            else
            {
                // ────────── 追加モード ──────────
                if (_devices.Any(d => d.Name == name))
                {
                    var r = MessageBox.Show(
                        $"「{name}」という機器が既に存在します。\n同名で追加しますか？",
                        "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (r != MessageBoxResult.Yes) return;
                }

                var dev = new Device
                {
                    Name  = name,
                    Model = AddDevModelBox.Text.Trim(),
                    IP    = AddDevIPBox.Text.Trim()
                };
                _deviceIndexMap[dev] = _devices.Count;
                _devices.Add(dev);
            }

            // 全コンボ・チップを更新（編集時は RenderChildren も必要 ─ DisplayText が変わるため）
            foreach (var b in _blocks) RenderChildren(b);
            RefreshAllCombosAndChips();

            CloseAddDevicePanel();
        }

        // ====================================================
        // HELPERS
        // ====================================================
        private ConnBlock? GetBlock(int id) => _blockDict.TryGetValue(id, out var b) ? b : null;

        private enum BtnKind { Primary, Secondary, Danger, Ghost }

        private Button MakeButton(string label, BtnKind kind)
        {
            var btn = new Button { Content = label };
            btn.Style = kind switch
            {
                BtnKind.Primary   => (Style)Resources["BtnPrimary"],
                BtnKind.Secondary => (Style)Resources["BtnSecondary"],
                BtnKind.Danger    => (Style)Resources["BtnDanger"],
                BtnKind.Ghost     => (Style)Resources["BtnGhost"],
                _                 => (Style)Resources["BtnSecondary"]
            };
            return btn;
        }

        private ComboBox MakeCombo()
        {
            var combo = new ComboBox
            {
                MinWidth                   = 200,
                HorizontalAlignment        = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
            };
            // Apply style directly (do not use Loaded event - timing is unreliable)
            if (Resources.Contains("DarkCombo"))
                combo.Style = (Style)Resources["DarkCombo"];
            return combo;
        }

        private ComboBoxItem MakeComboItem(string text, object? tag, SolidColorBrush fg,
            FontStyle? style = null, bool isHeader = false)
        {
            return new ComboBoxItem
            {
                Content     = text,
                Tag         = tag,
                Foreground  = fg,
                Background  = C_Surface2,
                FontFamily  = FF_Consolas,
                FontStyle   = style ?? FontStyles.Normal,
                IsEnabled   = !isHeader
            };
        }
    }

    // ====================================================
    // DRAW.IO XML BUILDER
    // ====================================================
    public static class DrawioXml
    {
        public static string Build(List<ConnBlock> blocks, string diagramName = "ネットワーク構成図", string memo = "")
        {
            // Collect all devices to draw
            var placed = new HashSet<Device>();
            foreach (var b in blocks)
            {
                if (b.Parent != null) placed.Add(b.Parent);
                foreach (var ch in b.Children)
                    if (ch.Device != null) placed.Add(ch.Device);
            }

            int nextId = 2;
            var nodeId = new Dictionary<Device, int>();
            foreach (var d in placed) nodeId[d] = nextId++;

            // Layout
            var pos = ComputeLayout(blocks, nodeId);

            var sb = new StringBuilder();
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            sb.AppendLine($"<mxfile host=\"NetworkDiagramBuilder\" modified=\"{now}\" agent=\"NetworkDiagramBuilder\" version=\"21.6.6\" type=\"device\">");
            sb.AppendLine($"  <diagram name=\"{XmlEsc(diagramName)}\" id=\"net-diagram\">");
            sb.AppendLine("    <mxGraphModel dx=\"1422\" dy=\"762\" grid=\"1\" gridSize=\"10\" guides=\"1\" tooltips=\"1\" connect=\"1\" arrows=\"1\" fold=\"1\" page=\"1\" pageScale=\"1\" pageWidth=\"1169\" pageHeight=\"827\" math=\"0\" shadow=\"0\">");
            sb.AppendLine("      <root>");
            sb.AppendLine("        <mxCell id=\"0\" />");
            sb.AppendLine("        <mxCell id=\"1\" parent=\"0\" />");

            // Nodes
            foreach (var kvp in nodeId)
            {
                var d   = kvp.Key;
                var nid = kvp.Value;
                var (x, y) = pos.TryGetValue(d, out var p) ? p : (40, 40);

                // value属性内のHTMLはXMLエスケープが必要
                // <b>名前</b><br>型番<br>IP → &lt;b&gt;名前&lt;/b&gt;&lt;br&gt;型番&lt;br&gt;IP
                var labelParts = new List<string>();
                if (!string.IsNullOrEmpty(d.Name))  labelParts.Add($"&lt;b&gt;{XmlEsc(d.Name)}&lt;/b&gt;");
                if (!string.IsNullOrEmpty(d.Model)) labelParts.Add(XmlEsc(d.Model));
                if (!string.IsNullOrEmpty(d.IP))    labelParts.Add(XmlEsc(d.IP));
                string label = string.Join("&lt;br&gt;", labelParts);

                string style = RectStyle(d);
                sb.AppendLine($"        <mxCell id=\"{nid}\" value=\"{label}\" style=\"{style}\" vertex=\"1\" parent=\"1\">");
                sb.AppendLine($"          <mxGeometry x=\"{x}\" y=\"{y}\" width=\"160\" height=\"70\" as=\"geometry\" />");
                sb.AppendLine($"        </mxCell>");
            }

            // Edges
            // ── Step1: 重複排除しながら全エッジを収集 ──
            var edgeList = new List<(int src, int tgt, ChildConn ch, Device parentDev)>();
            var drawnCheck = new HashSet<string>();
            foreach (var b in blocks)
            {
                if (b.Parent == null || !nodeId.ContainsKey(b.Parent)) continue;
                int src = nodeId[b.Parent];
                foreach (var ch in b.Children)
                {
                    if (ch.IsNone || ch.Device == null || !nodeId.ContainsKey(ch.Device)) continue;
                    int tgt = nodeId[ch.Device];
                    string key = $"{Math.Min(src, tgt)}-{Math.Max(src, tgt)}";
                    if (drawnCheck.Contains(key)) continue;
                    drawnCheck.Add(key);
                    edgeList.Add((src, tgt, ch, b.Parent));
                }
            }

            // ── Step2: ノードごとの exit/entry 本数を集計 ──
            var srcCount = new Dictionary<int, int>();
            var tgtCount = new Dictionary<int, int>();
            foreach (var (src, tgt, _, _) in edgeList)
            {
                srcCount[src] = srcCount.GetValueOrDefault(src) + 1;
                tgtCount[tgt] = tgtCount.GetValueOrDefault(tgt) + 1;
            }

            // ── Step3: 各ノードの割り当て済みインデックス ──
            var srcIdx = new Dictionary<int, int>();
            var tgtIdx = new Dictionary<int, int>();

            // ── Step4: エッジ生成 ──
            foreach (var (src, tgt, ch, parentDev) in edgeList)
            {
                int si = srcIdx.GetValueOrDefault(src); srcIdx[src] = si + 1;
                int ti = tgtIdx.GetValueOrDefault(tgt); tgtIdx[tgt] = ti + 1;

                // 均等分散: 1本→0.5, 2本→0.33/0.67, 3本→0.25/0.5/0.75 …
                double exitX  = DistX(si, srcCount[src]);
                double entryX = DistX(ti, tgtCount[tgt]);

                string edgeStyle =
                    $"edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;" +
                    $"jettySize=auto;" +
                    $"exitX={exitX:F4};exitY=1;exitDx=0;exitDy=0;" +
                    $"entryX={entryX:F4};entryY=0;entryDx=0;entryDy=0;";

                int edgeId = nextId++;
                sb.AppendLine($"        <mxCell id=\"{edgeId}\" value=\"\" style=\"{edgeStyle}\" edge=\"1\" source=\"{src}\" target=\"{tgt}\" parent=\"1\">");
                sb.AppendLine($"          <mxGeometry relative=\"1\" as=\"geometry\" />");
                sb.AppendLine($"        </mxCell>");

                // ポートラベルはノード座標から絶対位置で配置（edgeLabel相対位置だとエッジ長依存でずれるため）
                const int NodeW = 160, NodeH = 70;
                const string portLabelStyle =
                    "text;html=1;align=left;verticalAlign=middle;resizable=0;" +
                    "fontSize=10;fontFamily=Consolas;labelBackgroundColor=none;";

                if (!string.IsNullOrWhiteSpace(ch.PortSrc))
                {
                    // 接続元ノードのすぐ下: ノード下辺 + 4px、exitX位置から左に30px
                    var (sx, sy) = pos.TryGetValue(parentDev, out var sp) ? sp : (0, 0);
                    int lx = (int)(sx + exitX * NodeW) - 30;
                    int ly = sy + NodeH + 4;
                    sb.AppendLine($"        <mxCell id=\"{nextId++}\" value=\"{XmlEsc(ch.PortSrc.Trim())}\" style=\"{portLabelStyle}\" vertex=\"1\" parent=\"1\">");
                    sb.AppendLine($"          <mxGeometry x=\"{lx}\" y=\"{ly}\" width=\"60\" height=\"16\" as=\"geometry\" />");
                    sb.AppendLine($"        </mxCell>");
                }

                if (!string.IsNullOrWhiteSpace(ch.PortDst))
                {
                    // 接続先ノードのすぐ上: ノード上辺 - 20px、entryX位置から左に30px
                    var (tx, ty) = pos.TryGetValue(ch.Device!, out var tp) ? tp : (0, 0);
                    int lx = (int)(tx + entryX * NodeW) - 30;
                    int ly = ty - 20;
                    sb.AppendLine($"        <mxCell id=\"{nextId++}\" value=\"{XmlEsc(ch.PortDst.Trim())}\" style=\"{portLabelStyle}\" vertex=\"1\" parent=\"1\">");
                    sb.AppendLine($"          <mxGeometry x=\"{lx}\" y=\"{ly}\" width=\"60\" height=\"16\" as=\"geometry\" />");
                    sb.AppendLine($"        </mxCell>");
                }
            }

            // Memo textbox（図の右側に配置）
            if (!string.IsNullOrEmpty(memo))
            {
                // 図全体の右端X座標を取得
                int maxX = pos.Values.Count > 0 ? pos.Values.Max(p => p.x) : 40;
                int memoX = maxX + 200;  // ノード幅160 + 余白40

                // 図全体の上端Y座標（最小Y）
                int minY = pos.Values.Count > 0 ? pos.Values.Min(p => p.y) : 40;

                // 改行を draw.io HTML 形式に変換
                var memoLines = memo.Split('\n');
                string memoHtml = string.Join("&lt;br&gt;",
                    memoLines.Select(l => XmlEsc(l.TrimEnd('\r'))));

                // 高さ：1行あたり約18px + ヘッダー（メモ）分30px + 余白20px
                int memoHeight = Math.Max(80, memoLines.Length * 18 + 50);

                const string memoStyle =
                    "text;html=1;strokeColor=#d6b656;fillColor=#fffacd;" +
                    "align=left;verticalAlign=top;whiteSpace=wrap;" +
                    "rounded=1;arcSize=4;fontSize=11;";

                sb.AppendLine($"        <mxCell id=\"{nextId++}\" value=\"&lt;b&gt;メモ&lt;/b&gt;&lt;br&gt;{memoHtml}\" style=\"{memoStyle}\" vertex=\"1\" parent=\"1\">");
                sb.AppendLine($"          <mxGeometry x=\"{memoX}\" y=\"{minY}\" width=\"280\" height=\"{memoHeight}\" as=\"geometry\" />");
                sb.AppendLine($"        </mxCell>");
            }

            sb.AppendLine("      </root>");
            sb.AppendLine("    </mxGraphModel>");
            sb.AppendLine("  </diagram>");
            sb.Append("</mxfile>");
            return sb.ToString();
        }

        // ── Tree layout ──
        private static Dictionary<Device, (int x, int y)> ComputeLayout(
            List<ConnBlock> blocks, Dictionary<Device, int> nodeId)
        {
            var result   = new Dictionary<Device, (int, int)>();
            // PERF-01: List → HashSet で Contains を O(1) に改善
            var children = new Dictionary<Device, HashSet<Device>>();
            var allChild = new HashSet<Device>();

            foreach (var b in blocks)
            {
                if (b.Parent == null) continue;
                if (!children.ContainsKey(b.Parent)) children[b.Parent] = new HashSet<Device>();
                foreach (var ch in b.Children)
                    if (ch.Device != null)
                    {
                        // 二重接続(IsDouble)はツリー構造に含めない（位置は元の場所を維持）
                        if (ch.IsDouble) continue;
                        if (children[b.Parent].Add(ch.Device))  // HashSet.Add は追加成功時 true
                            allChild.Add(ch.Device);
                    }
            }

            var roots = children.Keys.Where(d => !allChild.Contains(d)).ToList();
            if (roots.Count == 0 && children.Count > 0) roots.Add(children.Keys.First());

            const int NodeW   = 160;  // ノード幅
            const int NodeH   = 70;   // ノード高さ
            const int XMargin = 40;   // 横マージン
            const int YGap    = 70;   // 縦ギャップ

            int rootX = 80;
            foreach (var root in roots)
            {
                int subtreeW = GetSubtreeWidth(root, children, NodeW, XMargin);
                int centerX  = rootX + subtreeW / 2;
                LayoutNode(root, centerX, 40, children, result, NodeW, NodeH + YGap, XMargin);
                rootX += subtreeW + XMargin;
            }

            // 二重接続ノード：接続元親の1段下（子と同じ高さ）に右側へ配置
            foreach (var b in blocks)
            {
                if (b.Parent == null || !result.ContainsKey(b.Parent)) continue;
                var (px, py) = result[b.Parent];
                foreach (var ch in b.Children)
                {
                    if (!ch.IsDouble || ch.Device == null) continue;
                    if (result.ContainsKey(ch.Device)) continue;  // 既に配置済みなら skip
                    // 接続元の親の右隣・同じ深さに配置
                    int sibY = py;  // 親と同じ高さ（既存ノードと同レベル）
                    // 既に使われていないX座標を右端から探す
                    int usedMaxX = result.Values.Count > 0 ? result.Values.Max(p => p.Item1) : 40;
                    result[ch.Device] = (usedMaxX + NodeW + XMargin, sibY);
                }
            }

            // Orphan devices（レイアウト外 → ツリーの最下段より下に並べる）
            int orphanMaxY = result.Values.Count > 0
                ? result.Values.Max(p => p.Item2) + NodeH + YGap
                : 40;
            int ox = 40, oy = orphanMaxY;
            foreach (var d in nodeId.Keys)
                if (!result.ContainsKey(d)) { result[d] = (ox, oy); ox += NodeW + XMargin; }

            return result;
        }

        // サブツリーが必要とする横幅を計算（重なり防止の核心）
        // BUG-A: visited ガードなしだと循環グラフ (A→B→A) で StackOverflowException
        private static int GetSubtreeWidth(Device node,
            Dictionary<Device, HashSet<Device>> ch, int nodeW, int xMargin,
            HashSet<Device>? visited = null)
        {
            if (!ch.ContainsKey(node) || ch[node].Count == 0)
                return nodeW + xMargin;
            // 循環検知: 既に展開中のノードは1スロット分として返す
            visited ??= new HashSet<Device>();
            if (!visited.Add(node)) return nodeW + xMargin;
            int width = ch[node].Sum(k => GetSubtreeWidth(k, ch, nodeW, xMargin, visited));
            visited.Remove(node);   // 兄弟サブツリーの展開のため除去
            return width;
        }

        private static void LayoutNode(Device node, int centerX, int y,
            Dictionary<Device, HashSet<Device>> ch,
            Dictionary<Device, (int, int)> result, int nodeW, int yStep, int xMargin)
        {
            if (result.ContainsKey(node)) return;
            result[node] = (centerX - nodeW / 2, y);
            if (!ch.ContainsKey(node)) return;

            // 既配置ノードは除外（二重接続などで既に位置が決まっているもの）
            var kids = ch[node].Where(k => !result.ContainsKey(k)).ToList();
            if (kids.Count == 0) return;

            int totalW  = kids.Sum(k => GetSubtreeWidth(k, ch, nodeW, xMargin));
            int startX  = centerX - totalW / 2;

            foreach (var kid in kids)
            {
                int kidW    = GetSubtreeWidth(kid, ch, nodeW, xMargin);
                int kidCenter = startX + kidW / 2;
                LayoutNode(kid, kidCenter, y + yStep, ch, result, nodeW, yStep, xMargin);
                startX += kidW;
            }
        }

        // ── Rectangle style (color-coded by device type) ──
        private static string RectStyle(Device d)
        {
            string n = (d.Name + d.Model).ToLower();

            // 共通ベース
            const string Base = "rounded=1;whiteSpace=wrap;html=1;arcSize=8;" +
                                 "fontFamily=Consolas;fontSize=11;align=center;verticalAlign=middle;";

            if (ContainsAny(n, "fortigate", "fg-", "fg6", "fg7", "fw", "firewall", "asa", "utm", "paloalto", "pa-"))
                // オレンジ系：ファイアウォール/UTM
                return Base + "fillColor=#FFE6CC;strokeColor=#d6b656;fontColor=#1a1a1a;";

            if (ContainsAny(n, "router", "rt", "rtx", "nec", "cisco", "c18", "c29"))
                // 緑系：ルーター（yamaha除外）
                return Base + "fillColor=#D5E8D4;strokeColor=#82b366;fontColor=#1a1a1a;";

            if (ContainsAny(n, "sw", "switch", "hub", "sg", "gs", "catalyst", "nexus"))
                // 青系：スイッチ/HUB
                return Base + "fillColor=#DAE8FC;strokeColor=#6c8ebf;fontColor=#1a1a1a;";

            if (ContainsAny(n, "ap", "wifi", "wap", "wireless", "wlx", "wax", "airo"))
                // 紫系：AP
                return Base + "fillColor=#E1D5E7;strokeColor=#9673a6;fontColor=#1a1a1a;";

            if (ContainsAny(n, "nas", "server", "srv", "ds", "rs", "esxi", "vmware"))
                // サーモン系：サーバー/NAS
                return Base + "fillColor=#FFD7CC;strokeColor=#d0472a;fontColor=#1a1a1a;";

            if (ContainsAny(n, "pc", "client", "ws", "desktop", "laptop", "note"))
                // グレー系：PC
                return Base + "fillColor=#F5F5F5;strokeColor=#666666;fontColor=#1a1a1a;";

            if (ContainsAny(n, "modem", "onu", "olt", "ont", "ctl"))
                // 黄系：ONU/モデム
                return Base + "fillColor=#FFFACD;strokeColor=#b8860b;fontColor=#1a1a1a;";

            if (ContainsAny(n, "mfp", "複合機", "printer", "プリンター", "bizhub", "ricoh", "xerox", "canon", "epson", "sharp", "konica"))
                // ティール系：複合機
                return Base + "fillColor=#D5EDF8;strokeColor=#0e7490;fontColor=#1a1a1a;";

            if (ContainsAny(n, "tel", "主装置", "pbx", "phone", "電話", "ひかり電話", "ip-pbx", "ippbx", "ntt", "nakayo", "saxa", "panasonic"))
                // ピンク系：主装置/TEL
                return Base + "fillColor=#FCE4EC;strokeColor=#c2185b;fontColor=#1a1a1a;";

            // デフォルト（白）
            return Base + "fillColor=#ffffff;strokeColor=#888888;fontColor=#1a1a1a;";
        }

        private static bool ContainsAny(string src, params string[] terms)
            => terms.Any(src.Contains);

        // 接続点のX座標を均等分散（1本→0.5、2本→0.33/0.67、3本→0.25/0.5/0.75…）
        private static double DistX(int idx, int total) =>
            total <= 1 ? 0.5 : (idx + 1.0) / (total + 1.0);

        private static string XmlEsc(string s) =>
            s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
             .Replace("\"", "&quot;").Replace("'", "&apos;");
    }
}
