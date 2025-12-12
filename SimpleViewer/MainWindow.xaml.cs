using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Windows.Threading;
using System.Windows.Controls.Primitives;
using WinForms = System.Windows.Forms;

namespace SimpleViewer;

public partial class MainWindow : Window
{
    private enum MediaType
    {
        Image,
        Video
    }

    private enum SortMode
    {
        NewestCreated,
        NewestModified,
        Name,
        Random
    }

    private const string DeleteFolderName = "_delete_";
    private const string AppVersion = "1.1.0";
    private const int FolderIconResourceId = -4;
    private const int HelpIconResourceId = -99;
    private const int OptionsIconResourceId = -114;
    private const int RecentIconResourceId = -117;
    private const double MinImageZoom = 0.2;
    private const double MaxImageZoom = 6.0;
    private const double ImageZoomStep = 0.2;
    private const double VideoVolumeStep = 0.05;
    private const double MinGridZoom = 0.6;
    private const double MaxGridZoom = 2.5;
    private const double GridZoomStep = 0.15;
    private const int MaxRecentFolders = 5;

    private sealed class MediaItem
    {
        private BitmapSource? _thumbnail;

        public MediaItem(string path, MediaType type, DateTime createdAt, DateTime modifiedAt)
        {
            Path = path;
            Type = type;
            CreatedAt = createdAt;
            ModifiedAt = modifiedAt;
        }

        public string Path { get; }
        public MediaType Type { get; }
        public DateTime CreatedAt { get; }
        public DateTime ModifiedAt { get; }
        public string FileName => System.IO.Path.GetFileName(Path);
        public bool IsImage => Type == MediaType.Image;
        public bool IsVideo => Type == MediaType.Video;
        public BitmapSource? Thumbnail => _thumbnail ??= CreateThumbnail();

        private BitmapSource? CreateThumbnail()
        {
            if (!IsImage)
            {
                return null;
            }

            try
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(Path);
                bitmap.DecodePixelWidth = 320;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                return bitmap;
            }
            catch
            {
                return null;
            }
        }
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
    private readonly Dictionary<string, string> _metadataCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _recentFolders = new();
    private readonly ScaleTransform _imageScaleTransform = new(1.0, 1.0);
    private readonly TranslateTransform _imagePanTransform = new();
    private readonly TransformGroup _imageTransformGroup = new();
    private readonly string _recentFoldersPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SimpleViewer", "recentFolders.json");
    private static string? _iconBase64;
    private static readonly string IconLibraryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "imageres.dll");
    private static readonly ImageSource? FolderIconSource = LoadSystemIconImage(IconLibraryPath, FolderIconResourceId);
    private static readonly ImageSource? HelpIconSource = LoadSystemIconImage(IconLibraryPath, HelpIconResourceId);
    private static readonly ImageSource? OptionsIconSource = LoadSystemIconImage(IconLibraryPath, OptionsIconResourceId);
    private static readonly ImageSource? RecentIconSource = LoadSystemIconImage(IconLibraryPath, RecentIconResourceId);
    private static readonly (string Match, string Label)[] SummaryKeyMappings =
    {
        ("positive prompt", "Prompt"),
        ("prompt", "Prompt"),
        ("negative prompt", "Negative Prompt"),
        ("sampler", "Sampler"),
        ("model", "Model"),
        ("checkpoint", "Model"),
        ("steps", "Steps"),
        ("cfg scale", "CFG Scale"),
        ("cfg", "CFG Scale"),
        ("seed", "Seed"),
        ("size", "Size"),
        ("width", "Width"),
        ("height", "Height"),
        ("denoise", "Denoise"),
        ("clip skip", "Clip Skip")
    };

    private double _slideshowIntervalSeconds = 4;
    private int _currentIndex = -1;
    private bool _isUpdatingInterval;
    private bool _isFullScreen;
    private WindowState _previousWindowState;
    private WindowStyle _previousWindowStyle;
    private ResizeMode _previousResizeMode;
    private string _currentFolder = string.Empty;
    private double _imageZoom = 1.0;
    private double _videoVolume = 0.5;
    private bool _isVideoMuted;
    private bool _isVideoPaused;
    private SortMode _sortMode = SortMode.NewestCreated;
    private readonly Random _random = new();
    private bool _isGridView;
    private bool _ignoreGridSelectionChange;
    private bool _isLoading;
    private bool _isPanningImage;
    private System.Windows.Point _panStartPoint;
    private double _gridZoom = 1.0;

    public MainWindow()
    {
        InitializeComponent();
        if (ImageViewer is not null)
        {
            _imageTransformGroup.Children.Add(_imageScaleTransform);
            _imageTransformGroup.Children.Add(_imagePanTransform);
            ImageViewer.RenderTransform = _imageTransformGroup;
            ImageViewer.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
        }

        Title = $"Simple Viewer v{AppVersion}";

        _slideshowTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(_slideshowIntervalSeconds)
        };
        _slideshowTimer.Tick += (_, _) => ShowNextItem(fromTimer: true);
        ApplySystemIcons();
        UpdateSortMenuChecks();
        LoadRecentFolders();
    }

    private async void BrowseButton_Click(object sender, RoutedEventArgs e)
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

            await LoadFolderAsync(dialog.SelectedPath);
            FolderPathText.Text = dialog.SelectedPath;
        }
    }

    private void OptionsButton_Click(object sender, RoutedEventArgs e)
    {
        if (OptionsButton?.ContextMenu is ContextMenu menu)
        {
            menu.PlacementTarget = OptionsButton;
            menu.Placement = PlacementMode.Bottom;
            menu.IsOpen = true;
        }
    }

    private void RecentFoldersButton_Click(object sender, RoutedEventArgs e)
    {
        if (RecentFoldersButton?.ContextMenu is ContextMenu menu)
        {
            menu.PlacementTarget = RecentFoldersButton;
            menu.Placement = PlacementMode.Bottom;
            menu.IsOpen = true;
        }
    }

    private async Task LoadFolderAsync(string folderPath, bool resetHidden = true, bool preserveSelection = false)
    {
        if (_isLoading || !Directory.Exists(folderPath) || IsDeleteFolder(folderPath))
        {
            return;
        }

        _isLoading = true;
        SetLoadingState(true);

        if (IsDeleteFolder(folderPath))
        {
            _isLoading = false;
            Mouse.OverrideCursor = null;
            return;
        }

        try
        {
            var includeSubfolders = RecursiveCheckBox?.IsChecked == true;
            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var items = await Task.Run(() => BuildMediaItems(folderPath, searchOption));

            _currentFolder = folderPath;
            AddRecentFolder(folderPath);
            _allItems.Clear();
            if (resetHidden)
            {
                _hiddenPaths.Clear();
            }

            if (!preserveSelection)
            {
                _currentIndex = -1;
            }

            _allItems.AddRange(items);
            SortAllItems();
            ApplyFilters();
        }
        finally
        {
            _isLoading = false;
            SetLoadingState(false);
        }
    }

    private void SortAllItems()
    {
        switch (_sortMode)
        {
            case SortMode.Name:
                _allItems.Sort((a, b) =>
                {
                    var result = string.Compare(a.FileName, b.FileName, StringComparison.OrdinalIgnoreCase);
                    return result != 0 ? result : string.Compare(a.Path, b.Path, StringComparison.OrdinalIgnoreCase);
                });
                break;
            case SortMode.NewestModified:
                _allItems.Sort((a, b) =>
                {
                    var comparison = DateTime.Compare(b.ModifiedAt, a.ModifiedAt);
                    return comparison != 0
                        ? comparison
                        : string.Compare(a.Path, b.Path, StringComparison.OrdinalIgnoreCase);
                });
                break;
            case SortMode.Random:
                Shuffle(_allItems);
                break;
            case SortMode.NewestCreated:
            default:
                _allItems.Sort((a, b) =>
                {
                    var comparison = DateTime.Compare(b.CreatedAt, a.CreatedAt);
                    return comparison != 0
                        ? comparison
                        : string.Compare(a.Path, b.Path, StringComparison.OrdinalIgnoreCase);
                });
                break;
        }
    }

    private void Shuffle<T>(IList<T> list)
    {
        for (var i = list.Count - 1; i > 0; i--)
        {
            var j = _random.Next(i + 1);
            (list[i], list[j]) = (list[j], list[i]);
        }
    }

    private static List<MediaItem> BuildMediaItems(string folderPath, SearchOption searchOption)
    {
        var items = new List<MediaItem>();

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
            var modifiedAt = File.GetLastWriteTime(file);

            if (ImageExtensions.Contains(extension))
            {
                items.Add(new MediaItem(file, MediaType.Image, createdAt, modifiedAt));
            }
            else if (VideoExtensions.Contains(extension))
            {
                items.Add(new MediaItem(file, MediaType.Video, createdAt, modifiedAt));
            }
        }

        return items;
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

        OnFilteredItemsUpdated();
        UpdateSlideshowState();
    }

    private void DisplayCurrentItem()
    {
        if (_isGridView)
        {
            UpdateGridSelection();
            return;
        }

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
        _isVideoPaused = false;
        var fileName = GetDisplayFileName(path);
        ResetImageZoom();

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
            UpdateMetadataPanelIfVisible();
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
        VideoViewer.Volume = _videoVolume;
        VideoViewer.IsMuted = _isVideoMuted;
        _isVideoPaused = false;
        UpdateMediaInfo($"File: {fileName}", GetVideoInfoText());
        UpdateMetadataPanelIfVisible();
    }

    private void ResetImageZoom()
    {
        _imageZoom = 1.0;
        _imageScaleTransform.ScaleX = 1.0;
        _imageScaleTransform.ScaleY = 1.0;
        ResetImagePan();
    }

    private void AdjustImageZoom(bool zoomIn)
    {
        var delta = zoomIn ? ImageZoomStep : -ImageZoomStep;
        _imageZoom = Math.Clamp(_imageZoom + delta, MinImageZoom, MaxImageZoom);
        _imageScaleTransform.ScaleX = _imageZoom;
        _imageScaleTransform.ScaleY = _imageZoom;
        if (_imageZoom <= 1.0)
        {
            ResetImagePan();
        }
    }

    private void ResetImagePan()
    {
        _imagePanTransform.X = 0;
        _imagePanTransform.Y = 0;
    }

    private bool TryHandleVideoVolumeKey(Key key)
    {
        if (VideoViewer.Visibility != Visibility.Visible)
        {
            return false;
        }

        switch (key)
        {
            case Key.Up:
                AdjustVideoVolume(increase: true);
                return true;
            case Key.Down:
                AdjustVideoVolume(increase: false);
                return true;
            default:
                return false;
        }
    }

    private void AdjustGridZoom(bool zoomIn)
    {
        _gridZoom = Math.Clamp(_gridZoom + (zoomIn ? GridZoomStep : -GridZoomStep), MinGridZoom, MaxGridZoom);
        if (GridViewScaleTransform is not null)
        {
            GridViewScaleTransform.ScaleX = _gridZoom;
            GridViewScaleTransform.ScaleY = _gridZoom;
        }
    }

    private bool ToggleVideoPlayback()
    {
        if (VideoViewer.Visibility != Visibility.Visible)
        {
            return false;
        }

        if (_isVideoPaused)
        {
            VideoViewer.Play();
            _isVideoPaused = false;
        }
        else
        {
            VideoViewer.Pause();
            _isVideoPaused = true;
        }

        UpdateVideoVolumeDisplay();
        return true;
    }

    private void AdjustVideoVolume(bool increase)
    {
        _videoVolume = Math.Clamp(_videoVolume + (increase ? VideoVolumeStep : -VideoVolumeStep), 0.0, 1.0);
        _isVideoMuted = false;
        VideoViewer.Volume = _videoVolume;
        VideoViewer.IsMuted = false;
        UpdateVideoVolumeDisplay();
    }

    private void ToggleVideoMute()
    {
        _isVideoMuted = !_isVideoMuted;
        VideoViewer.IsMuted = _isVideoMuted;
        UpdateVideoVolumeDisplay();
    }

    private void UpdateVideoVolumeDisplay()
    {
        if (MediaInfoText is null || VideoViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        MediaInfoText.Text = GetVideoInfoText();
    }

    private string GetVideoInfoText()
    {
        string status = _isVideoMuted
            ? "Volume: Muted"
            : $"Volume: {Math.Round(_videoVolume * 100)}%";

        if (_isVideoPaused)
        {
            status += " (Paused)";
        }

        return status;
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
        _isVideoPaused = false;
        UpdateMediaInfo("File: --", "Size: --");
        if (GridViewControl is not null)
        {
            GridViewControl.SelectedIndex = -1;
        }
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
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
        {
            if (e.Key == Key.C)
            {
                if (TryCopyCurrentImageToClipboard())
                {
                    e.Handled = true;
                    return;
                }
            }
        }

        if ((e.Key == Key.Up || e.Key == Key.Down) && TryHandleVideoVolumeKey(e.Key))
        {
            e.Handled = true;
            return;
        }

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
            case Key.G:
                ToggleGridView();
                e.Handled = true;
                break;
            case Key.S:
                ToggleSlideshowShortcut();
                e.Handled = true;
                break;
            case Key.Space:
                if (ToggleVideoPlayback())
                {
                    e.Handled = true;
                }

                break;
            case Key.M:
                if (VideoViewer.Visibility == Visibility.Visible)
                {
                    ToggleVideoMute();
                    e.Handled = true;
                }

                break;
            case Key.D0:
            case Key.NumPad0:
                if (ImageViewer.Visibility == Visibility.Visible && GetCurrentItem()?.Type == MediaType.Image)
                {
                    ResetImageZoom();
                    e.Handled = true;
                }

                break;
            case Key.Escape:
                if (_isFullScreen)
                {
                    ToggleFullScreen();
                }
                else
                {
                    Close();
                }

                e.Handled = true;
                break;
            case Key.X:
                ToggleMetadataPanel();
                e.Handled = true;
                break;
        }
    }

    private void Window_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (e.Delta == 0)
        {
            return;
        }

        if (_isGridView)
        {
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                AdjustGridZoom(e.Delta > 0);
                e.Handled = true;
            }

            return;
        }

        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
        {
            if (ImageViewer.Visibility == Visibility.Visible && GetCurrentItem()?.Type == MediaType.Image)
            {
                AdjustImageZoom(e.Delta > 0);
                e.Handled = true;
            }

            return;
        }

        if (e.Delta > 0)
        {
            ShowPreviousItem();
            e.Handled = true;
        }
        else if (e.Delta < 0)
        {
            ShowNextItem();
            e.Handled = true;
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

    private void ToggleSlideshowShortcut()
    {
        if (SlideshowCheckBox is null)
        {
            return;
        }

        var newState = SlideshowCheckBox.IsChecked != true;
        SlideshowCheckBox.IsChecked = newState;
        UpdateSlideshowState();
        ShowSlideshowIndicator(newState);
    }

    private void LoadRecentFolders()
    {
        _recentFolders.Clear();

        try
        {
            if (File.Exists(_recentFoldersPath))
            {
                var json = File.ReadAllText(_recentFoldersPath);
                var restored = JsonSerializer.Deserialize<List<string>>(json);
                if (restored is not null)
                {
                    foreach (var entry in restored)
                    {
                        var normalized = NormalizeFolderPath(entry);
                        if (string.IsNullOrWhiteSpace(normalized))
                        {
                            continue;
                        }

                        if (!_recentFolders.Any(p => string.Equals(p, normalized, StringComparison.OrdinalIgnoreCase)))
                        {
                            _recentFolders.Add(normalized);
                        }
                    }
                }
            }
        }
        catch
        {
            // ignore persistence failures
        }

        TrimRecentFolders();
        UpdateRecentFoldersMenu();
    }

    private void SaveRecentFolders()
    {
        try
        {
            var directory = Path.GetDirectoryName(_recentFoldersPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(_recentFoldersPath, JsonSerializer.Serialize(_recentFolders));
        }
        catch
        {
            // ignore persistence failures
        }
    }

    private void AddRecentFolder(string folderPath)
    {
        var normalized = NormalizeFolderPath(folderPath);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        _recentFolders.RemoveAll(p => string.Equals(p, normalized, StringComparison.OrdinalIgnoreCase));
        _recentFolders.Insert(0, normalized);
        TrimRecentFolders();
        SaveRecentFolders();
        UpdateRecentFoldersMenu();
    }

    private void TrimRecentFolders()
    {
        if (_recentFolders.Count <= MaxRecentFolders)
        {
            return;
        }

        _recentFolders.RemoveRange(MaxRecentFolders, _recentFolders.Count - MaxRecentFolders);
    }

    private static string NormalizeFolderPath(string? folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath))
        {
            return string.Empty;
        }

        try
        {
            return Path.GetFullPath(folderPath.Trim());
        }
        catch
        {
            return folderPath.Trim();
        }
    }

    private void UpdateRecentFoldersMenu()
    {
        if (RecentFoldersContextMenu is null)
        {
            return;
        }

        RecentFoldersContextMenu.Items.Clear();

        if (_recentFolders.Count == 0)
        {
            RecentFoldersContextMenu.Items.Add(new MenuItem
            {
                Header = "(No recent folders)",
                IsEnabled = false
            });
            return;
        }

        foreach (var folder in _recentFolders)
        {
            var item = new MenuItem
            {
                Header = folder,
                Tag = folder
            };
            item.Click += RecentFolderMenuItem_Click;
            RecentFoldersContextMenu.Items.Add(item);
        }

        RecentFoldersContextMenu.Items.Add(new Separator());
        var clearItem = new MenuItem
        {
            Header = "Clear history"
        };
        clearItem.Click += ClearRecentFoldersMenuItem_Click;
        RecentFoldersContextMenu.Items.Add(clearItem);
    }

    private async void RecentFolderMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: string folderPath })
        {
            return;
        }

        if (!Directory.Exists(folderPath))
        {
            _recentFolders.RemoveAll(p => string.Equals(p, folderPath, StringComparison.OrdinalIgnoreCase));
            SaveRecentFolders();
            UpdateRecentFoldersMenu();
            return;
        }

        if (IsDeleteFolder(folderPath))
        {
            System.Windows.MessageBox.Show("The _delete_ folder is reserved and cannot be viewed.", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        await LoadFolderAsync(folderPath);
        FolderPathText.Text = folderPath;
    }

    private void ClearRecentFoldersMenuItem_Click(object sender, RoutedEventArgs e)
    {
        _recentFolders.Clear();
        SaveRecentFolders();
        UpdateRecentFoldersMenu();
    }

    private void ShowSlideshowIndicator(bool isPlaying)
    {
        if (SlideshowIndicator is null || SlideshowIndicatorText is null)
        {
            return;
        }

        SlideshowIndicator.BeginAnimation(UIElement.OpacityProperty, null);
        SlideshowIndicatorText.Text = isPlaying ? "▶" : "⏸";
        SlideshowIndicator.Visibility = Visibility.Visible;
        SlideshowIndicator.Opacity = 1.0;

        var animation = new DoubleAnimation
        {
            From = 1.0,
            To = 0.0,
            BeginTime = TimeSpan.FromMilliseconds(150),
            Duration = TimeSpan.FromSeconds(1.8),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut },
            FillBehavior = FillBehavior.Stop
        };

        animation.Completed += (_, _) =>
        {
            if (SlideshowIndicator is not null)
            {
                SlideshowIndicator.Visibility = Visibility.Collapsed;
                SlideshowIndicator.Opacity = 1.0;
                SlideshowIndicator.BeginAnimation(UIElement.OpacityProperty, null);
            }
        };

        SlideshowIndicator.BeginAnimation(UIElement.OpacityProperty, animation);
    }

    private bool TryCopyCurrentImageToClipboard()
    {
        var item = GetCurrentItem();
        if (item is null || !item.IsImage || !File.Exists(item.Path))
        {
            return false;
        }

        try
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.UriSource = new Uri(item.Path);
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.EndInit();
            bitmap.Freeze();
            System.Windows.Clipboard.SetImage(bitmap);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private async void RecursiveCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentFolder) || !Directory.Exists(_currentFolder))
        {
            return;
        }

        await LoadFolderAsync(_currentFolder, resetHidden: true, preserveSelection: false);
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

    private async void RefreshCurrentFolder()
    {
        if (string.IsNullOrWhiteSpace(_currentFolder) || !Directory.Exists(_currentFolder))
        {
            return;
        }

        await LoadFolderAsync(_currentFolder, resetHidden: true, preserveSelection: false);
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

    private async void Window_Drop(object sender, System.Windows.DragEventArgs e)
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
                await LoadFolderAsync(path);
                FolderPathText.Text = path;
                break;
            }
        }
    }

    private void ImageViewer_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_imageZoom <= 1.0 || ImageViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        _isPanningImage = true;
        _panStartPoint = e.GetPosition(this);
        ImageViewer.CaptureMouse();
        e.Handled = true;
    }

    private void ImageViewer_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isPanningImage || ImageViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        var current = e.GetPosition(this);
        var delta = current - _panStartPoint;
        _panStartPoint = current;
        _imagePanTransform.X += delta.X;
        _imagePanTransform.Y += delta.Y;
    }

    private void ImageViewer_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isPanningImage)
        {
            return;
        }

        _isPanningImage = false;
        ImageViewer.ReleaseMouseCapture();
        e.Handled = true;
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
            OnFilteredItemsUpdated();
            UpdateSlideshowState();
            return;
        }

        if (_currentIndex >= _filteredItems.Count)
        {
            _currentIndex = 0;
        }

        DisplayCurrentItem();
        OnFilteredItemsUpdated();
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

    private void ToggleGridView(bool? enable = null)
    {
        var target = enable ?? !_isGridView;
        if (_isGridView == target)
        {
            return;
        }

        _isGridView = target;
        if (_isGridView)
        {
            VideoViewer.Stop();
            VideoViewer.Source = null;
            VideoViewer.Visibility = Visibility.Collapsed;
            ImageViewer.Visibility = Visibility.Collapsed;
            if (GridViewControl is not null)
            {
                GridViewControl.Visibility = Visibility.Visible;
                RefreshGridViewItems();
                UpdateGridSelection();
            }
        }
        else
        {
            if (GridViewControl is not null)
            {
                GridViewControl.Visibility = Visibility.Collapsed;
            }

            DisplayCurrentItem();
        }

        if (SingleViewMenuItem is not null)
        {
            SingleViewMenuItem.IsChecked = !_isGridView;
        }

        if (GridViewMenuItem is not null)
        {
            GridViewMenuItem.IsChecked = _isGridView;
        }
    }

    private void RefreshGridViewItems()
    {
        if (GridViewControl is null)
        {
            return;
        }

        GridViewControl.ItemsSource = null;
        GridViewControl.ItemsSource = _filteredItems;
    }

    private void UpdateGridSelection()
    {
        if (GridViewControl is null)
        {
            return;
        }

        _ignoreGridSelectionChange = true;
        if (_currentIndex >= 0 && _currentIndex < _filteredItems.Count)
        {
            GridViewControl.SelectedIndex = _currentIndex;
            GridViewControl.ScrollIntoView(GridViewControl.SelectedItem);
        }
        else
        {
            GridViewControl.SelectedIndex = -1;
        }

        _ignoreGridSelectionChange = false;
    }

    private void OnFilteredItemsUpdated()
    {
        RefreshGridViewItems();
        if (_isGridView)
        {
            UpdateGridSelection();
        }
    }

    private void GridViewControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_isGridView || GridViewControl is null || _ignoreGridSelectionChange)
        {
            return;
        }

        if (GridViewControl.SelectedIndex >= 0 && GridViewControl.SelectedIndex < _filteredItems.Count)
        {
            _currentIndex = GridViewControl.SelectedIndex;
        }
    }

    private void GridViewControl_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (GridViewControl?.SelectedIndex is int index && index >= 0 && index < _filteredItems.Count)
        {
            _currentIndex = index;
            ToggleGridView(enable: false);
        }
    }

    private void ViewModeMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender == GridViewMenuItem)
        {
            ToggleGridView(enable: true);
        }
        else if (sender == SingleViewMenuItem)
        {
            ToggleGridView(enable: false);
        }
    }

    private void SortMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem item || item.Tag is null)
        {
            return;
        }

        if (!Enum.TryParse(item.Tag.ToString(), out SortMode newMode))
        {
            return;
        }

        if (_sortMode == newMode)
        {
            item.IsChecked = true;
            return;
        }

        _sortMode = newMode;
        UpdateSortMenuChecks();
        SortAllItems();
        ApplyFilters();
    }

    private void UpdateSortMenuChecks()
    {
        if (SortCreatedMenuItem is not null)
        {
            SortCreatedMenuItem.IsChecked = _sortMode == SortMode.NewestCreated;
        }

        if (SortModifiedMenuItem is not null)
        {
            SortModifiedMenuItem.IsChecked = _sortMode == SortMode.NewestModified;
        }

        if (SortNameMenuItem is not null)
        {
            SortNameMenuItem.IsChecked = _sortMode == SortMode.Name;
        }

        if (SortRandomMenuItem is not null)
        {
            SortRandomMenuItem.IsChecked = _sortMode == SortMode.Random;
        }
    }

    private void SetLoadingState(bool isLoading)
    {
        if (isLoading)
        {
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            if (LoadingOverlay is not null)
            {
                LoadingOverlay.Visibility = Visibility.Visible;
            }
        }
        else
        {
            Mouse.OverrideCursor = null;
            if (LoadingOverlay is not null)
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
            }
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
    <h2>Keyboard Shortcuts</h2>
    <h3>Navigation</h3>
    <ul>
        <li>Arrow Left / Right - Previous or next item</li>
        <li>Mouse Wheel - Previous (up) or next (down) item</li>
        <li>F11 - Toggle full screen / hide controls</li>
        <li>ESC - Exit full screen (or close the app)</li>
        <li>F5 - Refresh folder contents from disk</li>
        <li>Delete - Move current item to the <code>_delete_</code> folder</li>
    </ul>
    <h3>Recent folders</h3>
    <ul>
        <li>Recent button - Open one of the last 5 folders.</li>
        <li>"Clear history" inside that menu removes all remembered folders.</li>
    </ul>
    <h3>Slideshow</h3>
    <ul>
        <li>S - Play/Pause slideshow (uses selected interval)</li>
        <li>Slideshow checkbox - Enable auto-advance</li>
    </ul>
    <h3>Images</h3>
    <ul>
        <li>Ctrl + Mouse Wheel - Zoom in/out</li>
        <li>Drag (while zoomed) - Pan the image</li>
        <li>0 - Reset image zoom to 100%</li>
        <li>Ctrl + C - Copy the current image to the clipboard</li>
        <li>X - Open/close EXIF &amp; Comfy metadata panel</li>
    </ul>
    <h3>Grid View</h3>
    <ul>
        <li>G - Toggle grid view</li>
        <li>Ctrl + Mouse Wheel - Change thumbnail size</li>
    </ul>
    <h3>Videos</h3>
    <ul>
        <li>Arrow Up/Down - Adjust video volume while playing</li>
        <li>M - Toggle mute for the current video</li>
        <li>Space - Pause/Resume video playback</li>
    </ul>
</div>
<div>
    <h2>Release Highlights</h2>
    <h3>v1.1.0</h3>
    <ul>
        <li><strong>Toolbar & layout</strong>
            <ul>
                <li>New toolbar layout: Folder picker, Recent folders, Options, and Help buttons with system-themed icons.</li>
                <li>Options moved into a compact gear menu that hosts Filters, Slideshow, View (single/grid), and Sort-by settings.</li>
            </ul>
        </li>
        <li><strong>Recent folders</strong>
            <ul>
                <li>Added Recent folders button with persistent MRU list (last 5 paths).</li>
                <li>Quick reopen any recent path or clear the list from the same menu.</li>
            </ul>
        </li>
        <li><strong>Grid view</strong>
            <ul>
                <li>Added Grid view (G) with thumbnail browsing.</li>
                <li>Double-click a tile to return to single-image view.</li>
                <li>Ctrl + Mouse Wheel adjusts thumbnail size in grid view.</li>
            </ul>
        </li>
        <li><strong>Slideshow</strong>
            <ul>
                <li><code>S</code> toggles slideshow play/pause.</li>
                <li>Centered play/pause overlay appears briefly when the slideshow state changes.</li>
                <li>Timer behavior improved when manually navigating between items.</li>
            </ul>
        </li>
        <li><strong>Viewing & navigation</strong>
            <ul>
                <li>Mouse wheel now navigates between items (up = previous, down = next).</li>
                <li>Ctrl + Mouse Wheel zooms the current image in/out.</li>
                <li>Drag-to-pan is available when zoomed in; <code>0</code> resets zoom to 100%.</li>
                <li><code>ESC</code> exits full screen or closes the app when not in full screen.</li>
            </ul>
        </li>
        <li><strong>Media controls</strong>
            <ul>
                <li>Arrow Up/Down adjust video volume while playing.</li>
                <li><code>M</code> toggles mute for the current video.</li>
                <li><code>Space</code> pauses/resumes the current video.</li>
            </ul>
        </li>
        <li><strong>Metadata & clipboard</strong>
            <ul>
                <li>EXIF/Comfy metadata panel with Copy button for the full text.</li>
                <li>Ctrl + C copies the current image directly to the clipboard.</li>
            </ul>
        </li>
        <li><strong>Help & docs</strong>
            <ul>
                <li>Help window redesigned with grouped shortcut sections and an embedded change log.</li>
            </ul>
        </li>
    </ul>
    <h3>v1.0.0</h3>
    <ul>
        <li>Initial release: image/video viewer with slideshow, metadata viewer, sorting options, and basic zoom/pan.</li>
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
        RecentFoldersButton.Content = CreateIconElement(RecentIconSource, "\uE81C");
        if (OptionsButton is not null)
        {
            OptionsButton.Content = CreateIconElement(OptionsIconSource, "\uE713");
        }
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

    private void ToggleMetadataPanel()
    {
        if (MetadataPanel is null)
        {
            return;
        }

        if (MetadataPanel.Visibility == Visibility.Visible)
        {
            MetadataPanel.Visibility = Visibility.Collapsed;
            return;
        }

        ShowMetadataForCurrentItem();
    }

    private void ShowMetadataForCurrentItem()
    {
        if (MetadataPanel is null || MetadataTextBox is null)
        {
            return;
        }

        var item = GetCurrentItem();
        if (item is null)
        {
            MetadataTextBox.Text = "No item selected.";
            MetadataPanel.Visibility = Visibility.Visible;
            return;
        }

        if (item.Type != MediaType.Image)
        {
            MetadataTextBox.Text = "Metadata is currently available for images only.";
            MetadataPanel.Visibility = Visibility.Visible;
            return;
        }

        var text = GetOrLoadMetadata(item.Path);
        MetadataTextBox.Text = text;
        MetadataPanel.Visibility = Visibility.Visible;
    }

    private void UpdateMetadataPanelIfVisible()
    {
        if (MetadataPanel is null || MetadataPanel.Visibility != Visibility.Visible)
        {
            return;
        }

        ShowMetadataForCurrentItem();
    }

    private string GetOrLoadMetadata(string path)
    {
        if (_metadataCache.TryGetValue(path, out var cached))
        {
            return cached;
        }

        var metadata = ExtractMetadata(path);
        _metadataCache[path] = metadata;
        return metadata;
    }

    private static string ExtractMetadata(string path)
    {
        string? wpfMetadata = null;
        string? codecError = null;
        var summaryEntries = new List<(string Key, string Value)>();
        var seenSummaryKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var summaryNotes = new List<string>();

        // Try standard WPF metadata first.
        try
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
            if (decoder.Frames.Count > 0)
            {
                var frame = decoder.Frames[0];
                if (frame.Metadata is BitmapMetadata bitmapMetadata)
                {
                    var builder = new StringBuilder();

                    builder.AppendLine($"File: {Path.GetFileName(path)}");
                    builder.AppendLine($"Format: {bitmapMetadata.Format}");
                    builder.AppendLine();

                    AppendMetadataField(builder, "Title", bitmapMetadata.Title);
                    AppendMetadataField(builder, "Subject", bitmapMetadata.Subject);
                    AppendMetadataField(builder, "Comment", bitmapMetadata.Comment);
                    AppendMetadataField(builder, "Camera Manufacturer", bitmapMetadata.CameraManufacturer);
                    AppendMetadataField(builder, "Camera Model", bitmapMetadata.CameraModel);
                    AppendMetadataField(builder, "Date Taken", bitmapMetadata.DateTaken);

                    if (!string.IsNullOrWhiteSpace(bitmapMetadata.CameraModel))
                    {
                        AddSummaryEntry(summaryEntries, seenSummaryKeys, "Camera", bitmapMetadata.CameraModel);
                    }

                    if (!string.IsNullOrWhiteSpace(bitmapMetadata.DateTaken))
                    {
                        AddSummaryEntry(summaryEntries, seenSummaryKeys, "Date Taken", bitmapMetadata.DateTaken);
                    }

                    AddSummaryEntry(summaryEntries, seenSummaryKeys, "Pixel Size", $"{frame.PixelWidth} x {frame.PixelHeight}");

                    // Dump additional metadata (including potential ComfyUI text) where available.
                    AppendMetadataTree(builder, bitmapMetadata, string.Empty);

                    if (builder.Length > 0)
                    {
                        wpfMetadata = builder.ToString();
                    }
                }
            }
        }
        catch (Exception ex)
        {
            codecError = ex.Message;
        }

        // ComfyUI and similar tools often store their workflow text in PNG tEXt / zTXt / iTXt chunks.
        string? pngTextMetadata = null;
        if (string.Equals(Path.GetExtension(path), ".png", StringComparison.OrdinalIgnoreCase))
        {
            pngTextMetadata = ExtractPngTextMetadata(path, summaryEntries, seenSummaryKeys, summaryNotes);
        }

        var sections = new List<string>();
        if (summaryEntries.Count > 0 || summaryNotes.Count > 0)
        {
            var summaryBuilder = new StringBuilder();
            summaryBuilder.AppendLine("Metadata summary:");
            foreach (var entry in summaryEntries)
            {
                summaryBuilder.AppendLine($"{entry.Key}: {entry.Value}");
            }

            if (summaryNotes.Count > 0)
            {
                summaryBuilder.AppendLine();
                foreach (var note in summaryNotes)
                {
                    summaryBuilder.AppendLine($"- {note}");
                }
            }

            sections.Add(summaryBuilder.ToString().TrimEnd());
        }

        if (!string.IsNullOrWhiteSpace(wpfMetadata))
        {
            sections.Add($"Standard metadata:{Environment.NewLine}{wpfMetadata.Trim()}");
        }

        if (!string.IsNullOrWhiteSpace(pngTextMetadata))
        {
            sections.Add(pngTextMetadata.Trim());
        }

        if (sections.Count > 0)
        {
            return string.Join($"{Environment.NewLine}{Environment.NewLine}", sections);
        }

        if (!string.IsNullOrWhiteSpace(codecError))
        {
            return $"Unable to read metadata via image codec.\n{codecError}";
        }

        return "No metadata found.";
    }

    private static void AppendMetadataField(StringBuilder builder, string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        builder.AppendLine($"{label}: {value}");
    }

    private static string ExtractPngTextMetadata(
        string path,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys,
        List<string> summaryNotes)
    {
        try
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var reader = new BinaryReader(stream);

            // Validate PNG signature.
            var signature = reader.ReadBytes(8);
            var pngSignature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
            if (signature.Length != pngSignature.Length || !signature.SequenceEqual(pngSignature))
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            var hadAny = false;

            while (stream.Position + 8 <= stream.Length)
            {
                var lengthBytes = reader.ReadBytes(4);
                if (lengthBytes.Length < 4)
                {
                    break;
                }

                var length = BitConverter.ToInt32(lengthBytes.Reverse().ToArray(), 0);
                if (length < 0 || stream.Position + 4 + length + 4 > stream.Length)
                {
                    break;
                }

                var typeBytes = reader.ReadBytes(4);
                var chunkType = Encoding.ASCII.GetString(typeBytes);
                var data = reader.ReadBytes(length);

                // Skip CRC
                reader.ReadUInt32();

                switch (chunkType)
                {
                    case "tEXt":
                        AppendTextChunk(builder, data, ref hadAny, summaryEntries, seenSummaryKeys, summaryNotes);
                        break;
                    case "zTXt":
                        AppendZTextChunk(builder, data, ref hadAny, summaryEntries, seenSummaryKeys, summaryNotes);
                        break;
                    case "iTXt":
                        AppendInternationalTextChunk(builder, data, ref hadAny, summaryEntries, seenSummaryKeys, summaryNotes);
                        break;
                }

                if (chunkType == "IEND")
                {
                    break;
                }
            }

            if (!hadAny)
            {
                return string.Empty;
            }

            return builder.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void AppendTextChunk(
        StringBuilder builder,
        byte[] data,
        ref bool hadAny,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys,
        List<string> summaryNotes)
    {
        var separator = Array.IndexOf(data, (byte)0);
        if (separator <= 0 || separator >= data.Length - 1)
        {
            return;
        }

        var keyword = Encoding.ASCII.GetString(data, 0, separator);
        var text = Encoding.UTF8.GetString(data, separator + 1, data.Length - separator - 1);

        if (!hadAny)
        {
            builder.AppendLine("PNG text metadata (tEXt):");
            builder.AppendLine();
        }

        hadAny = true;
        builder.AppendLine($"[{keyword}]");
        builder.AppendLine(text);
        builder.AppendLine();

        CollectSummaryFromText(keyword, text, summaryEntries, seenSummaryKeys, summaryNotes);
    }

    private static void AppendZTextChunk(
        StringBuilder builder,
        byte[] data,
        ref bool hadAny,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys,
        List<string> summaryNotes)
    {
        var separator = Array.IndexOf(data, (byte)0);
        if (separator <= 0 || separator >= data.Length - 2)
        {
            return;
        }

        var keyword = Encoding.ASCII.GetString(data, 0, separator);
        var compressionMethod = data[separator + 1];
        var compressed = data.Skip(separator + 2).ToArray();
        var text = DecompressPngText(compressed);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        if (!hadAny)
        {
            builder.AppendLine("PNG text metadata (zTXt):");
            builder.AppendLine();
        }

        hadAny = true;
        builder.AppendLine($"[{keyword}] (compressed, method {compressionMethod})");
        builder.AppendLine(text);
        builder.AppendLine();

        CollectSummaryFromText(keyword, text, summaryEntries, seenSummaryKeys, summaryNotes);
    }

    private static void AppendInternationalTextChunk(
        StringBuilder builder,
        byte[] data,
        ref bool hadAny,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys,
        List<string> summaryNotes)
    {
        var offset = 0;

        int FindNull(byte[] buffer, int start)
        {
            for (var i = start; i < buffer.Length; i++)
            {
                if (buffer[i] == 0)
                {
                    return i;
                }
            }

            return -1;
        }

        var keywordEnd = FindNull(data, offset);
        if (keywordEnd <= 0)
        {
            return;
        }

        var keyword = Encoding.ASCII.GetString(data, offset, keywordEnd - offset);
        offset = keywordEnd + 1;
        if (offset + 2 > data.Length)
        {
            return;
        }

        var compressionFlag = data[offset++];
        var compressionMethod = data[offset++];

        var languageEnd = FindNull(data, offset);
        if (languageEnd < 0)
        {
            return;
        }

        var languageTag = Encoding.ASCII.GetString(data, offset, languageEnd - offset);
        offset = languageEnd + 1;

        var translatedEnd = FindNull(data, offset);
        if (translatedEnd < 0)
        {
            return;
        }

        var translatedKeyword = Encoding.UTF8.GetString(data, offset, translatedEnd - offset);
        offset = translatedEnd + 1;
        if (offset > data.Length)
        {
            return;
        }

        var textBytes = data.Skip(offset).ToArray();
        string text;
        if (compressionFlag == 1)
        {
            text = DecompressPngText(textBytes);
        }
        else
        {
            text = Encoding.UTF8.GetString(textBytes);
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        if (!hadAny)
        {
            builder.AppendLine("PNG text metadata (iTXt):");
            builder.AppendLine();
        }

        hadAny = true;
        builder.AppendLine($"[{keyword}] lang={languageTag}, translated=\"{translatedKeyword}\", compressed={compressionFlag == 1}, method={compressionMethod}");
        builder.AppendLine(text);
        builder.AppendLine();

        CollectSummaryFromText(keyword, text, summaryEntries, seenSummaryKeys, summaryNotes);
    }

    private static string DecompressPngText(byte[] compressed)
    {
        try
        {
            using var input = new MemoryStream(compressed);

            // Skip zlib header bytes if present to make DeflateStream happier.
            if (compressed.Length > 2)
            {
                input.ReadByte();
                input.ReadByte();
            }

            using var deflate = new DeflateStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream();
            deflate.CopyTo(output);
            return Encoding.UTF8.GetString(output.ToArray());
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void CollectSummaryFromText(
        string keyword,
        string text,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys,
        List<string> summaryNotes)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        var trimmed = text.Trim();
        if (LooksLikeLargeJson(trimmed))
        {
            summaryNotes.Add($"[{keyword}] contains a workflow JSON block (raw metadata shown below).");
            return;
        }

        var lines = trimmed.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var line in lines)
        {
            var separator = line.IndexOf(':');
            if (separator <= 0 || separator >= line.Length - 1)
            {
                continue;
            }

            var rawKey = line[..separator].Trim();
            var value = line[(separator + 1)..].Trim();
            if (string.IsNullOrWhiteSpace(rawKey) || string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            if (TryMapSummaryKey(rawKey, out var displayKey))
            {
                AddSummaryEntry(summaryEntries, seenSummaryKeys, displayKey, value);
            }
        }
    }

    private static void AddSummaryEntry(List<(string Key, string Value)> summaryEntries, HashSet<string> seenSummaryKeys, string key, string? value)
    {
        if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        if (seenSummaryKeys.Add(key))
        {
            summaryEntries.Add((key, value));
        }
    }

    private static bool TryMapSummaryKey(string rawKey, out string label)
    {
        foreach (var (match, mappedLabel) in SummaryKeyMappings)
        {
            if (rawKey.IndexOf(match, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                label = mappedLabel;
                return true;
            }
        }

        label = string.Empty;
        return false;
    }

    private static bool LooksLikeLargeJson(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var trimmed = text.TrimStart();
        if (!trimmed.StartsWith("{", StringComparison.OrdinalIgnoreCase) &&
            !trimmed.StartsWith("[", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return trimmed.Length > 400 || trimmed.Contains("\"nodes\"", StringComparison.OrdinalIgnoreCase);
    }

    private static void AppendMetadataTree(StringBuilder builder, BitmapMetadata metadata, string prefix)
    {
        try
        {
            foreach (var key in metadata)
            {
                if (key is not string query)
                {
                    continue;
                }

                var fullQuery = string.IsNullOrEmpty(prefix) ? query : $"{prefix}{query}";
                object? value;
                try
                {
                    value = metadata.GetQuery(fullQuery);
                }
                catch
                {
                    continue;
                }

                if (value is null)
                {
                    continue;
                }

                if (value is BitmapMetadata childMetadata)
                {
                    AppendMetadataTree(builder, childMetadata, fullQuery);
                }
                else if (value is System.Collections.IEnumerable enumerable && value is not string)
                {
                    var parts = new List<string>();
                    foreach (var item in enumerable)
                    {
                        if (item is null)
                        {
                            continue;
                        }

                        parts.Add(item.ToString() ?? string.Empty);
                    }

                    if (parts.Count > 0)
                    {
                        builder.AppendLine($"{fullQuery}: {string.Join(", ", parts)}");
                    }
                }
                else
                {
                    builder.AppendLine($"{fullQuery}: {value}");
                }
            }
        }
        catch
        {
            // Best-effort dump; ignore failures.
        }
    }

    private void CopyMetadataButton_Click(object sender, RoutedEventArgs e)
    {
        if (MetadataTextBox is null || string.IsNullOrWhiteSpace(MetadataTextBox.Text))
        {
            System.Windows.MessageBox.Show("There is no metadata to copy.", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            System.Windows.Clipboard.SetText(MetadataTextBox.Text);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Unable to copy metadata to the clipboard.\n{ex.Message}", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void CloseMetadataButton_Click(object sender, RoutedEventArgs e)
    {
        if (MetadataPanel is not null)
        {
            MetadataPanel.Visibility = Visibility.Collapsed;
        }
    }
}
