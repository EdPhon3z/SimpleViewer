using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Threading;
using WinForms = System.Windows.Forms;

namespace SimpleViewer;

public partial class MainWindow : Window
{
    private enum MediaType
    {
        Image,
        Video
    }

    private const string DeleteFolderName = "_delete_";
    private const string AppVersion = "1.0.0";
    private const int FolderIconResourceId = -4;
    private const int HelpIconResourceId = -99;

    private sealed class MediaItem
    {
        public MediaItem(string path, MediaType type, DateTime createdAt)
        {
            Path = path;
            Type = type;
            CreatedAt = createdAt;
        }

        public string Path { get; }
        public MediaType Type { get; }
        public DateTime CreatedAt { get; }
    }

    private static readonly HashSet<string> ImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp"
    };

    private static readonly HashSet<string> VideoExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".mp4", ".mkv", ".avi", ".mov", ".wmv", ".m4v"
    };

    private readonly List<MediaItem> _allItems = new();
    private readonly List<MediaItem> _filteredItems = new();
    private readonly DispatcherTimer _slideshowTimer;
    private readonly HashSet<string> _hiddenPaths = new(StringComparer.OrdinalIgnoreCase);
    private static string? _iconBase64;
    private static readonly string IconLibraryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "imageres.dll");
    private static readonly ImageSource? FolderIconSource = LoadSystemIconImage(IconLibraryPath, FolderIconResourceId);
    private static readonly ImageSource? HelpIconSource = LoadSystemIconImage(IconLibraryPath, HelpIconResourceId);

    private double _slideshowIntervalSeconds = 4;
    private int _currentIndex = -1;
    private bool _isUpdatingInterval;
    private bool _isFullScreen;
    private WindowState _previousWindowState;
    private WindowStyle _previousWindowStyle;
    private ResizeMode _previousResizeMode;
    private string _currentFolder = string.Empty;

    public MainWindow()
    {
        InitializeComponent();
        Title = $"Simple Viewer v{AppVersion}";

        _slideshowTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(_slideshowIntervalSeconds)
        };
        _slideshowTimer.Tick += (_, _) => ShowNextItem(fromTimer: true);
        ApplySystemIcons();
    }

    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new WinForms.FolderBrowserDialog
        {
            Description = "Select a folder that contains media files.",
            SelectedPath = Directory.Exists(_currentFolder) ? _currentFolder : Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
            ShowNewFolderButton = false
        };

        if (dialog.ShowDialog() == WinForms.DialogResult.OK && Directory.Exists(dialog.SelectedPath))
        {
            if (IsDeleteFolder(dialog.SelectedPath))
            {
                System.Windows.MessageBox.Show("The _delete_ folder is reserved and cannot be viewed.", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            LoadFolder(dialog.SelectedPath);
            FolderPathText.Text = dialog.SelectedPath;
        }
    }

    private void LoadFolder(string folderPath, bool resetHidden = true, bool preserveSelection = false)
    {
        if (IsDeleteFolder(folderPath))
        {
            return;
        }

        _currentFolder = folderPath;
        _allItems.Clear();
        if (resetHidden)
        {
            _hiddenPaths.Clear();
        }

        if (!preserveSelection)
        {
            _currentIndex = -1;
        }

        var searchOption = RecursiveCheckBox?.IsChecked == true ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

        foreach (var file in Directory.EnumerateFiles(folderPath, "*.*", searchOption))
        {
            var extension = Path.GetExtension(file);
            if (extension is null)
            {
                continue;
            }

            var directoryOfFile = Path.GetDirectoryName(file);
            if (!string.IsNullOrEmpty(directoryOfFile) && IsDeleteFolder(directoryOfFile))
            {
                continue;
            }

            var createdAt = File.GetCreationTime(file);

            if (ImageExtensions.Contains(extension))
            {
                _allItems.Add(new MediaItem(file, MediaType.Image, createdAt));
            }
            else if (VideoExtensions.Contains(extension))
            {
                _allItems.Add(new MediaItem(file, MediaType.Video, createdAt));
            }
        }

        _allItems.Sort((a, b) =>
        {
            var comparison = DateTime.Compare(b.CreatedAt, a.CreatedAt);
            return comparison != 0
                ? comparison
                : string.Compare(a.Path, b.Path, StringComparison.OrdinalIgnoreCase);
        });
        ApplyFilters();
    }

    private void ApplyFilters()
    {
        if (ImagesCheckBox is null || VideosCheckBox is null)
        {
            return;
        }

        var currentPath = GetCurrentItem()?.Path;
        _filteredItems.Clear();

        var allowImages = ImagesCheckBox.IsChecked == true;
        var allowVideos = VideosCheckBox.IsChecked == true;

        if (allowImages || allowVideos)
        {
            foreach (var item in _allItems)
            {
                if ((_hiddenPaths.Contains(item.Path)))
                {
                    continue;
                }

                if ((item.Type == MediaType.Image && allowImages) ||
                    (item.Type == MediaType.Video && allowVideos))
                {
                    _filteredItems.Add(item);
                }
            }
        }

        if (_filteredItems.Count == 0)
        {
            _currentIndex = -1;
            ClearDisplay();
        }
        else
        {
            if (!string.IsNullOrEmpty(currentPath))
            {
                var index = _filteredItems.FindIndex(i => string.Equals(i.Path, currentPath, StringComparison.OrdinalIgnoreCase));
                _currentIndex = index >= 0 ? index : 0;
            }
            else
            {
                _currentIndex = 0;
            }

            DisplayCurrentItem();
        }

        UpdateSlideshowState();
    }

    private void DisplayCurrentItem()
    {
        var item = GetCurrentItem();
        if (item is null)
        {
            ClearDisplay();
            return;
        }

        switch (item.Type)
        {
            case MediaType.Image:
                ShowImage(item.Path);
                break;
            case MediaType.Video:
                ShowVideo(item.Path);
                break;
        }
    }

    private void ShowImage(string path)
    {
        VideoViewer.Stop();
        VideoViewer.Visibility = Visibility.Collapsed;
        VideoViewer.Source = null;
        var fileName = GetDisplayFileName(path);

        try
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.UriSource = new Uri(path);
            bitmap.EndInit();
            bitmap.Freeze();

            ImageViewer.Source = bitmap;
            ImageViewer.Visibility = Visibility.Visible;
            UpdateMediaInfo($"File: {fileName}", $"Size: {bitmap.PixelWidth} x {bitmap.PixelHeight}");
        }
        catch
        {
            ImageViewer.Source = null;
            UpdateMediaInfo($"File: {fileName}", "Size: --");
        }
    }

    private void ShowVideo(string path)
    {
        ImageViewer.Visibility = Visibility.Collapsed;
        ImageViewer.Source = null;
        var fileName = GetDisplayFileName(path);

        VideoViewer.Visibility = Visibility.Visible;
        VideoViewer.Source = new Uri(path);
        VideoViewer.Position = TimeSpan.Zero;
        VideoViewer.LoadedBehavior = MediaState.Manual;
        VideoViewer.UnloadedBehavior = MediaState.Manual;
        VideoViewer.Stop();
        VideoViewer.Play();
        UpdateMediaInfo($"File: {fileName}", "Size: --");
    }

    private void VideoViewer_MediaEnded(object? sender, RoutedEventArgs e)
    {
        if (VideoViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        VideoViewer.Position = TimeSpan.Zero;
        VideoViewer.Play();
    }

    private void UpdateMediaInfo(string fileNameText, string sizeText)
    {
        if (FileNameText is not null)
        {
            FileNameText.Text = fileNameText;
        }

        if (MediaInfoText is not null)
        {
            MediaInfoText.Text = sizeText;
        }
    }

    private void ClearDisplay()
    {
        if (ImageViewer is null || VideoViewer is null)
        {
            return;
        }

        ImageViewer.Source = null;
        ImageViewer.Visibility = Visibility.Collapsed;

        VideoViewer.Stop();
        VideoViewer.Source = null;
        VideoViewer.Visibility = Visibility.Collapsed;
        UpdateMediaInfo("File: --", "Size: --");
    }

    private MediaItem? GetCurrentItem()
    {
        if (_currentIndex < 0 || _currentIndex >= _filteredItems.Count)
        {
            return null;
        }

        return _filteredItems[_currentIndex];
    }

    private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Right:
            case Key.Down:
                ShowNextItem();
                e.Handled = true;
                break;
            case Key.Left:
            case Key.Up:
                ShowPreviousItem();
                e.Handled = true;
                break;
            case Key.F11:
                ToggleFullScreen();
                e.Handled = true;
                break;
            case Key.F5:
                RefreshCurrentFolder();
                e.Handled = true;
                break;
            case Key.Delete:
                RemoveCurrentItem();
                e.Handled = true;
                break;
        }
    }

    private void ShowNextItem(bool fromTimer = false)
    {
        if (_filteredItems.Count == 0)
        {
            return;
        }

        _currentIndex = (_currentIndex + 1) % _filteredItems.Count;
        DisplayCurrentItem();
        if (!fromTimer)
        {
            RestartSlideshowTimer();
        }
    }

    private void ShowPreviousItem()
    {
        if (_filteredItems.Count == 0)
        {
            return;
        }

        _currentIndex = (_currentIndex - 1 + _filteredItems.Count) % _filteredItems.Count;
        DisplayCurrentItem();
        RestartSlideshowTimer();
    }

    private void FilterCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        ApplyFilters();
    }

    private void SlideshowCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        UpdateSlideshowState();
    }

    private void RecursiveCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentFolder) || !Directory.Exists(_currentFolder))
        {
            return;
        }

        LoadFolder(_currentFolder, resetHidden: true, preserveSelection: false);
    }

    private void SlideshowIntervalBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (!IsLoaded || _isUpdatingInterval)
        {
            return;
        }

        var text = SlideshowIntervalBox.Text.Trim();
        if (!double.TryParse(text, out var seconds) || seconds <= 0)
        {
            _isUpdatingInterval = true;
            SlideshowIntervalBox.Text = _slideshowIntervalSeconds.ToString("0.##");
            SlideshowIntervalBox.CaretIndex = SlideshowIntervalBox.Text.Length;
            _isUpdatingInterval = false;
            return;
        }

        _slideshowIntervalSeconds = seconds;
        _slideshowTimer.Interval = TimeSpan.FromSeconds(_slideshowIntervalSeconds);
    }

    private void UpdateSlideshowState()
    {
        if (!IsLoaded || SlideshowCheckBox is null)
        {
            return;
        }

        if (SlideshowCheckBox.IsChecked == true && _filteredItems.Count > 0)
        {
            if (!_slideshowTimer.IsEnabled)
            {
                _slideshowTimer.Start();
            }
        }
        else
        {
            _slideshowTimer.Stop();
        }
    }

    private void RestartSlideshowTimer()
    {
        if (SlideshowCheckBox?.IsChecked == true && _slideshowTimer.IsEnabled)
        {
            _slideshowTimer.Stop();
            _slideshowTimer.Start();
        }
    }

    private void RefreshCurrentFolder()
    {
        if (string.IsNullOrWhiteSpace(_currentFolder) || !Directory.Exists(_currentFolder))
        {
            return;
        }

        LoadFolder(_currentFolder, resetHidden: true, preserveSelection: false);
    }

    private void Window_DragEnter(object sender, System.Windows.DragEventArgs e)
    {
        if (HasValidFolder(e.Data))
        {
            e.Effects = System.Windows.DragDropEffects.Copy;
            e.Handled = true;
        }
        else
        {
            e.Effects = System.Windows.DragDropEffects.None;
        }
    }

    private void Window_DragOver(object sender, System.Windows.DragEventArgs e)
    {
        Window_DragEnter(sender, e);
    }

    private void Window_Drop(object sender, System.Windows.DragEventArgs e)
    {
        if (!HasValidFolder(e.Data))
        {
            return;
        }

        if (e.Data.GetData(System.Windows.DataFormats.FileDrop) is not string[] paths)
        {
            return;
        }

        foreach (var path in paths)
        {
            if (Directory.Exists(path) && !IsDeleteFolder(path))
            {
                LoadFolder(path);
                FolderPathText.Text = path;
                break;
            }
        }
    }

    private void RemoveCurrentItem()
    {
        var item = GetCurrentItem();
        if (item is null)
        {
            return;
        }

        var deleteFolder = Path.Combine(Path.GetDirectoryName(item.Path) ?? _currentFolder, DeleteFolderName);
        var targetPath = GetUniqueDestinationPath(deleteFolder, Path.GetFileName(item.Path));

        try
        {
            Directory.CreateDirectory(deleteFolder);
            if (item.Type == MediaType.Video)
            {
                VideoViewer.Stop();
                VideoViewer.Source = null;
            }

            File.Move(item.Path, targetPath, overwrite: false);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Unable to move the file to '{DeleteFolderName}'.\n{ex.Message}", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _hiddenPaths.Add(item.Path);
        _allItems.RemoveAll(m => string.Equals(m.Path, item.Path, StringComparison.OrdinalIgnoreCase));

        if (_currentIndex >= 0 && _currentIndex < _filteredItems.Count)
        {
            _filteredItems.RemoveAt(_currentIndex);
        }

        if (_filteredItems.Count == 0)
        {
            _currentIndex = -1;
            ClearDisplay();
            UpdateSlideshowState();
            return;
        }

        if (_currentIndex >= _filteredItems.Count)
        {
            _currentIndex = 0;
        }

        DisplayCurrentItem();
        RestartSlideshowTimer();
    }

    private void HelpButton_Click(object sender, RoutedEventArgs e)
    {
        ShowHelpDialog();
    }

    private static bool IsDeleteFolder(string path)
    {
        return string.Equals(Path.GetFileName(Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)), DeleteFolderName, StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasValidFolder(System.Windows.IDataObject data)
    {
        if (!data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            return false;
        }

        if (data.GetData(System.Windows.DataFormats.FileDrop) is not string[] paths)
        {
            return false;
        }

        return paths.Any(path => Directory.Exists(path) && !IsDeleteFolder(path));
    }

    private static string GetUniqueDestinationPath(string folder, string originalFileName)
    {
        var name = Path.GetFileNameWithoutExtension(originalFileName);
        var extension = Path.GetExtension(originalFileName);
        var candidate = Path.Combine(folder, originalFileName);
        var index = 1;

        while (File.Exists(candidate))
        {
            var suffix = $"_{DateTime.Now:yyyyMMdd_HHmmssfff}_{index}";
            candidate = Path.Combine(folder, $"{name}{suffix}{extension}");
            index++;
        }

        return candidate;
    }

    private void ToggleFullScreen()
    {
        if (!_isFullScreen)
        {
            _previousWindowState = WindowState;
            _previousWindowStyle = WindowStyle;
            _previousResizeMode = ResizeMode;

            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            WindowState = WindowState.Maximized;
            ControlPanel.Visibility = Visibility.Collapsed;

            _isFullScreen = true;
        }
        else
        {
            WindowStyle = _previousWindowStyle;
            ResizeMode = _previousResizeMode;
            WindowState = _previousWindowState;
            ControlPanel.Visibility = Visibility.Visible;

            _isFullScreen = false;
        }
    }

    private void ShowHelpDialog()
    {
        var window = new Window
        {
            Title = $"Simple Viewer Help v{AppVersion}",
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Width = 540,
            Height = 640,
            Icon = Icon,
            Background = System.Windows.Media.Brushes.Black
        };

        var browser = new System.Windows.Controls.WebBrowser
        {
            HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
            VerticalAlignment = System.Windows.VerticalAlignment.Stretch
        };
        browser.Navigating += HelpBrowser_Navigating;
        browser.NavigateToString(BuildHelpHtml());

        window.Content = browser;
        window.ShowDialog();
    }

    private string GetDisplayFileName(string fullPath)
    {
        if (RecursiveCheckBox?.IsChecked == true && !string.IsNullOrEmpty(_currentFolder) && fullPath.StartsWith(_currentFolder, StringComparison.OrdinalIgnoreCase))
        {
            var relative = fullPath[_currentFolder.Length..].TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return relative;
        }

        return Path.GetFileName(fullPath);
    }

    private static void HelpBrowser_Navigating(object? sender, NavigatingCancelEventArgs e)
    {
        if (e.Uri is null)
        {
            return;
        }

        e.Cancel = true;
        try
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        }
        catch
        {
        }
    }

    private static string BuildHelpHtml()
    {
        var builder = new StringBuilder();
        builder.Append("""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    body {
        font-family: 'Segoe UI', sans-serif;
        background: #050505;
        color: #f5f5f5;
        margin: 0;
        padding: 24px;
    }
    .card {
        background: linear-gradient(145deg, #111827, #0b0f1a);
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 18px 35px rgba(0,0,0,0.5);
    }
    .header {
        display: flex;
        gap: 16px;
        align-items: center;
        margin-bottom: 18px;
    }
    .header img {
        width: 80px;
        height: 80px;
        border-radius: 20px;
        box-shadow: 0 6px 12px rgba(0,0,0,0.35);
    }
    h1 {
        margin: 0;
        font-size: 1.6rem;
    }
    p {
        margin: 0 0 12px 0;
        line-height: 1.5;
    }
    .links a {
        color: #5ad1ff;
        text-decoration: none;
    }
    .links a:hover {
        text-decoration: underline;
    }
    ul {
        padding-left: 20px;
        margin: 8px 0 0 0;
    }
</style>
</head>
<body>
<div class="card">
""");

        builder.Append("<div class=\"header\">");
        builder.Append($"<img src=\"data:image/jpeg;base64,{GetIconBase64()}\" alt=\"Simple Viewer Icon\"/>");
        builder.Append($"<div><h1>Simple Viewer v{AppVersion}</h1>");
        builder.Append("<p>Created by <strong>EdPhonez</strong>. Follow for updates and new builds.</p></div></div>");

        builder.Append("""
<div class="links">
    <p><strong>GitHub:</strong> <a href="https://github.com/EdPhon3z/SimpleViewer">github.com/EdPhon3z/SimpleViewer</a><br/>
    <strong>Website:</strong> <a href="https://www.edphonez.com/">www.edphonez.com</a><br/>
    <strong>Media:</strong> <a href="https://media.edphonez.com/">media.edphonez.com</a><br/>
    <strong>YouTube:</strong> <a href="https://www.youtube.com/@edphonez">@edphonez</a><br/>
    <strong>License:</strong> <a href="https://github.com/EdPhon3z/SimpleViewer/blob/main/LICENSE">EdPhonez Non-Commercial License</a><br/>
    <strong>Download:</strong> Grab <code>SimpleViewer_win-x64.zip</code> from the Releases tab, extract, and run <code>SimpleViewer.exe</code>.</p>
</div>
<div>
    <h2>Controls</h2>
    <ul>
        <li>F11 — Toggle full screen / hide controls</li>
        <li>Arrow Left/Up — Previous item</li>
        <li>Arrow Right/Down — Next item</li>
        <li>Delete — Move current item to the <code>_delete_</code> folder</li>
        <li>F5 — Refresh folder contents from disk</li>
        <li>Subfolders checkbox — Include media from child folders</li>
        <li>Slideshow checkbox — Auto-advance using the selected interval</li>
    </ul>
</div>
</div>
</body>
</html>
""");

        return builder.ToString();
    }

    private static string GetIconBase64()
    {
        if (!string.IsNullOrEmpty(_iconBase64))
        {
            return _iconBase64!;
        }

        try
        {
            var uri = new Uri("pack://application:,,,/assets/icon.jpg", UriKind.Absolute);
            using var stream = System.Windows.Application.GetResourceStream(uri)?.Stream;
            if (stream is null)
            {
                return _iconBase64 = string.Empty;
            }

            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            _iconBase64 = Convert.ToBase64String(ms.ToArray());
            return _iconBase64;
        }
        catch
        {
            return _iconBase64 = string.Empty;
        }
    }

    private void ApplySystemIcons()
    {
        BrowseButton.Content = CreateIconElement(FolderIconSource, "\uE8B7");
        HelpButton.Content = CreateIconElement(HelpIconSource, "\uE11B");
    }

    private static object CreateIconElement(ImageSource? source, string fallbackGlyph)
    {
        if (source is not null)
        {
            return new System.Windows.Controls.Image
            {
                Source = source,
                Width = 20,
                Height = 20,
                SnapsToDevicePixels = true
            };
        }

        return new TextBlock
        {
            Text = fallbackGlyph,
            FontFamily = new System.Windows.Media.FontFamily("Segoe MDL2 Assets"),
            FontSize = 18,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
    }

    private static ImageSource? LoadSystemIconImage(string libraryPath, int resourceIndex)
    {
        try
        {
            var handle = ExtractIconHandle(libraryPath, resourceIndex);
            if (handle == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                var source = Imaging.CreateBitmapSourceFromHIcon(handle, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(32, 32));
                source.Freeze();
                return source;
            }
            finally
            {
                DestroyIcon(handle);
            }
        }
        catch
        {
            return null;
        }
    }

    private static IntPtr ExtractIconHandle(string libraryPath, int resourceIndex)
    {
        var large = new IntPtr[1];
        var count = ExtractIconEx(libraryPath, resourceIndex, large, null, 1);
        return count > 0 ? large[0] : IntPtr.Zero;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern uint ExtractIconEx(string szFileName, int nIconIndex, IntPtr[]? phiconLarge, IntPtr[]? phiconSmall, uint nIcons);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);
}
