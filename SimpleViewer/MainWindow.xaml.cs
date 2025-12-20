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
using Microsoft.Win32;
using WinForms = System.Windows.Forms;
using TagLibFile = TagLib.File;
using TagLibMediaTypes = TagLib.MediaTypes;

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
    private const string AppVersion = "1.2.0";
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
    private const string ExplorerDirectoryOpenKey = @"Software\Classes\Directory\shell\SimpleViewer.Open";
    private const string ExplorerDirectoryAddKey = @"Software\Classes\Directory\shell\SimpleViewer.Add";

    private sealed record StartupRequest(bool Append, List<string> Paths);
    private sealed record ExplorerMenuEntry(string KeyPath, string DisplayName, bool AppendArgument);

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
    private static readonly string[] AllSupportedExtensions = ImageExtensions
        .Concat(VideoExtensions)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToArray();

    private readonly List<MediaItem> _allItems = new();
    private readonly List<MediaItem> _filteredItems = new();
    private readonly DispatcherTimer _slideshowTimer;
    private readonly HashSet<string> _hiddenPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _metadataCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> _recentFolders = new();
    private readonly ScaleTransform _imageScaleTransform = new(1.0, 1.0);
    private readonly TranslateTransform _imagePanTransform = new();
    private readonly TransformGroup _imageTransformGroup = new();
    private readonly ScaleTransform _videoScaleTransform = new(1.0, 1.0);
    private readonly TranslateTransform _videoPanTransform = new();
    private readonly TransformGroup _videoTransformGroup = new();
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
    private bool _isCustomSelection;
    private string _customSelectionLabel = string.Empty;
    private StartupRequest? _startupRequest;
    private readonly DispatcherTimer _videoProgressTimer;
    private double _imageZoom = 1.0;
    private double _videoZoom = 1.0;
    private double _videoVolume = 0.5;
    private bool _isVideoMuted;
    private bool _isVideoPaused;
    private SortMode _sortMode = SortMode.NewestCreated;
    private readonly Random _random = new();
    private bool _isGridView;
    private bool _ignoreGridSelectionChange;
    private bool _isLoading;
    private bool _isPanningImage;
    private bool _isPanningVideo;
    private System.Windows.Point _panStartPoint;
    private double _gridZoom = 1.0;
    private bool _isScrubbingVideo;
    private TimeSpan _knownVideoDuration = TimeSpan.Zero;
    private bool _slideshowWaitingForVideoEnd;

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

        if (VideoViewer is not null)
        {
            _videoTransformGroup.Children.Add(_videoScaleTransform);
            _videoTransformGroup.Children.Add(_videoPanTransform);
            VideoViewer.RenderTransform = _videoTransformGroup;
            VideoViewer.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
        }

        Title = $"Simple Viewer v{AppVersion}";

        _slideshowTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(_slideshowIntervalSeconds)
        };
        _slideshowTimer.Tick += (_, _) => ShowNextItem(fromTimer: true);
        _videoProgressTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(200)
        };
        _videoProgressTimer.Tick += (_, _) => UpdateVideoProgressDisplay();
        ApplySystemIcons();
        UpdateSortMenuChecks();
        LoadRecentFolders();
        _startupRequest = ParseStartupRequest(Environment.GetCommandLineArgs());
        Loaded += MainWindow_Loaded;
        UpdateSelectionLabel();
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

    private void InstallExplorerIntegrationMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (TryInstallExplorerIntegration())
            {
                System.Windows.MessageBox.Show("Explorer context menu installed. Right-click supported folders or media files to open or add them in Simple Viewer.", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Unable to install the Explorer context menu.\n{ex.Message}", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void RemoveExplorerIntegrationMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            foreach (var keyPath in EnumerateExplorerRegistryKeys())
            {
                Registry.CurrentUser.DeleteSubKeyTree(keyPath, throwOnMissingSubKey: false);
            }

            System.Windows.MessageBox.Show("Explorer context menu entries removed.", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Unable to remove the Explorer context menu.\n{ex.Message}", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private bool TryInstallExplorerIntegration()
    {
        var exePath = Process.GetCurrentProcess().MainModule?.FileName;
        if (string.IsNullOrWhiteSpace(exePath) || !File.Exists(exePath))
        {
            System.Windows.MessageBox.Show("Unable to locate SimpleViewer.exe for context menu registration.", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        foreach (var entry in EnumerateExplorerMenuEntries())
        {
            var command = entry.AppendArgument
                ? $"\"{exePath}\" --add \"%1\""
                : $"\"{exePath}\" \"%1\"";
            RegisterContextMenuKey(entry.KeyPath, entry.DisplayName, command, exePath);
        }

        return true;
    }

    private static void RegisterContextMenuKey(string keyPath, string displayName, string command, string? iconPath)
    {
        using var key = Registry.CurrentUser.CreateSubKey(keyPath);
        key?.SetValue(null, displayName);
        if (!string.IsNullOrWhiteSpace(iconPath))
        {
            key?.SetValue("Icon", iconPath);
        }

        using var commandKey = Registry.CurrentUser.CreateSubKey($"{keyPath}\\command");
        commandKey?.SetValue(null, command);
    }

    private static IEnumerable<ExplorerMenuEntry> EnumerateExplorerMenuEntries()
    {
        yield return new ExplorerMenuEntry(ExplorerDirectoryOpenKey, "Open in Simple Viewer", AppendArgument: false);
        yield return new ExplorerMenuEntry(ExplorerDirectoryAddKey, "Add to Simple Viewer selection", AppendArgument: true);

        foreach (var extension in AllSupportedExtensions)
        {
            var normalized = NormalizeExtensionForRegistry(extension);
            var basePath = $@"Software\Classes\SystemFileAssociations\{normalized}\shell\SimpleViewer";
            yield return new ExplorerMenuEntry($"{basePath}.Open", "Open in Simple Viewer", AppendArgument: false);
            yield return new ExplorerMenuEntry($"{basePath}.Add", "Add to Simple Viewer selection", AppendArgument: true);
        }
    }

    private static IEnumerable<string> EnumerateExplorerRegistryKeys()
    {
        return EnumerateExplorerMenuEntries()
            .Select(entry => entry.KeyPath)
            .Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static string NormalizeExtensionForRegistry(string extension)
    {
        if (string.IsNullOrWhiteSpace(extension))
        {
            return ".tmp";
        }

        return extension.StartsWith(".", StringComparison.Ordinal) ? extension : $".{extension}";
    }

    private StartupRequest? ParseStartupRequest(string[] args)
    {
        if (args is null || args.Length <= 1)
        {
            return null;
        }

        var paths = new List<string>();
        var append = false;
        for (int i = 1; i < args.Length; i++)
        {
            var arg = args[i];
            if (string.IsNullOrWhiteSpace(arg))
            {
                continue;
            }

            if (string.Equals(arg, "--add", StringComparison.OrdinalIgnoreCase))
            {
                append = true;
                continue;
            }

            paths.Add(arg.Trim('"'));
        }

        return paths.Count == 0 ? null : new StartupRequest(append, paths);
    }

    private async void MainWindow_Loaded(object? sender, RoutedEventArgs e)
    {
        Loaded -= MainWindow_Loaded;
        if (_startupRequest is null)
        {
            return;
        }

        await HandleStartupRequestAsync(_startupRequest);
    }

    private async Task HandleStartupRequestAsync(StartupRequest request)
    {
        var normalized = request.Paths
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Select(p => Environment.ExpandEnvironmentVariables(p.Trim('"')))
            .ToList();
        if (normalized.Count == 0)
        {
            return;
        }

        if (!request.Append && normalized.Count == 1 && Directory.Exists(normalized[0]) && !IsDeleteFolder(normalized[0]))
        {
            await LoadFolderAsync(normalized[0]);
            return;
        }

        var includeSubfolders = RecursiveCheckBox?.IsChecked == true;
        var newItems = await Task.Run(() => BuildMediaItemsFromSelection(normalized, includeSubfolders));
        if (newItems.Count == 0)
        {
            var folder = normalized.FirstOrDefault(Directory.Exists);
            if (!string.IsNullOrEmpty(folder) && !IsDeleteFolder(folder))
            {
                await LoadFolderAsync(folder);
            }

            return;
        }

        ApplyCustomItems(newItems, request.Append && _allItems.Count > 0);
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
            _isCustomSelection = false;
            _customSelectionLabel = string.Empty;
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
            UpdateSelectionLabel();
            SortAllItems();
            ApplyFilters();
        }
        finally
        {
            _isLoading = false;
            SetLoadingState(false);
        }
    }

    private void ApplyCustomItems(List<MediaItem> newItems, bool append, string? labelOverride = null)
    {
        if (!append)
        {
            _allItems.Clear();
            _hiddenPaths.Clear();
            _currentIndex = -1;
        }

        var knownPaths = new HashSet<string>(_allItems.Select(i => i.Path), StringComparer.OrdinalIgnoreCase);
        foreach (var item in newItems)
        {
            if (knownPaths.Add(item.Path))
            {
                _allItems.Add(item);
            }
        }

        if (_allItems.Count == 0)
        {
            System.Windows.MessageBox.Show("No supported images or videos were found.", "Simple Viewer", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _isCustomSelection = true;
        _currentFolder = string.Empty;
        UpdateSelectionLabel(labelOverride);
        SortAllItems();
        ApplyFilters();
    }

    private void UpdateSelectionLabel(string? customOverride = null)
    {
        if (FolderPathText is null)
        {
            return;
        }

        if (_isCustomSelection)
        {
            if (!string.IsNullOrWhiteSpace(customOverride))
            {
                _customSelectionLabel = customOverride!;
            }
            else
            {
                _customSelectionLabel = $"Custom selection ({_allItems.Count} items)";
            }

            FolderPathText.Text = _customSelectionLabel;
        }
        else
        {
            FolderPathText.Text = string.IsNullOrEmpty(_currentFolder) ? "No folder selected" : _currentFolder;
            _customSelectionLabel = string.Empty;
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

    private static bool TryCreateMediaItem(string filePath, out MediaItem? item)
    {
        item = null;
        if (!File.Exists(filePath))
        {
            return false;
        }

        var extension = Path.GetExtension(filePath);
        if (string.IsNullOrEmpty(extension))
        {
            return false;
        }

        var directoryOfFile = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directoryOfFile) && IsDeleteFolder(directoryOfFile))
        {
            return false;
        }

        MediaType type;
        if (ImageExtensions.Contains(extension))
        {
            type = MediaType.Image;
        }
        else if (VideoExtensions.Contains(extension))
        {
            type = MediaType.Video;
        }
        else
        {
            return false;
        }

        var createdAt = File.GetCreationTime(filePath);
        var modifiedAt = File.GetLastWriteTime(filePath);
        item = new MediaItem(filePath, type, createdAt, modifiedAt);
        return true;
    }

    private static List<MediaItem> BuildMediaItems(string folderPath, SearchOption searchOption)
    {
        var items = new List<MediaItem>();

        foreach (var file in Directory.EnumerateFiles(folderPath, "*.*", searchOption))
        {
            if (TryCreateMediaItem(file, out var mediaItem) && mediaItem is not null)
            {
                items.Add(mediaItem);
            }
        }

        return items;
    }

    private List<MediaItem> BuildMediaItemsFromSelection(IEnumerable<string> paths, bool includeSubfolders)
    {
        var items = new List<MediaItem>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

        foreach (var raw in paths)
        {
            var path = raw?.Trim();
            if (string.IsNullOrEmpty(path))
            {
                continue;
            }

            if (Directory.Exists(path) && !IsDeleteFolder(path))
            {
                foreach (var mediaItem in BuildMediaItems(path, searchOption))
                {
                    if (seen.Add(mediaItem.Path))
                    {
                        items.Add(mediaItem);
                    }
                }
            }
            else if (TryCreateMediaItem(path, out var item) && item is not null && seen.Add(item.Path))
            {
                items.Add(item);
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
        HideVideoProgressUi();
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
        finally
        {
            EnsureSlideshowTimerForCurrentMedia();
        }
    }

    private void ShowVideo(string path)
    {
        ImageViewer.Visibility = Visibility.Collapsed;
        ImageViewer.Source = null;
        var fileName = GetDisplayFileName(path);

        VideoViewer.Visibility = Visibility.Visible;
        ResetVideoZoom();
        VideoViewer.Source = new Uri(path);
        VideoViewer.Position = TimeSpan.Zero;
        VideoViewer.LoadedBehavior = MediaState.Manual;
        VideoViewer.UnloadedBehavior = MediaState.Manual;
        VideoViewer.Stop();
        VideoViewer.Play();
        VideoViewer.Volume = _videoVolume;
        VideoViewer.IsMuted = _isVideoMuted;
        _isVideoPaused = false;
        ShowVideoProgressUi();
        UpdateVideoProgressDisplay(TimeSpan.Zero);
        UpdateMediaInfo($"File: {fileName}", GetVideoInfoText());
        UpdateMetadataPanelIfVisible();
        EnsureSlideshowTimerForCurrentMedia();
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

    private void ResetVideoZoom()
    {
        _videoZoom = 1.0;
        _videoScaleTransform.ScaleX = 1.0;
        _videoScaleTransform.ScaleY = 1.0;
        ResetVideoPan();
    }

    private void AdjustVideoZoom(bool zoomIn)
    {
        var delta = zoomIn ? ImageZoomStep : -ImageZoomStep;
        _videoZoom = Math.Clamp(_videoZoom + delta, MinImageZoom, MaxImageZoom);
        _videoScaleTransform.ScaleX = _videoZoom;
        _videoScaleTransform.ScaleY = _videoZoom;
        if (_videoZoom <= 1.0)
        {
            ResetVideoPan();
        }
    }

    private void ResetImagePan()
    {
        _imagePanTransform.X = 0;
        _imagePanTransform.Y = 0;
    }

    private void ResetVideoPan()
    {
        _videoPanTransform.X = 0;
        _videoPanTransform.Y = 0;
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

        if (_slideshowWaitingForVideoEnd && SlideshowCheckBox?.IsChecked == true)
        {
            _slideshowWaitingForVideoEnd = false;
            ShowNextItem(fromTimer: true);
            return;
        }

        VideoViewer.Position = TimeSpan.Zero;
        _isVideoPaused = false;
        VideoViewer.Play();
        UpdateVideoProgressDisplay(TimeSpan.Zero);
    }

    private void VideoViewer_MediaOpened(object? sender, RoutedEventArgs e)
    {
        if (VideoViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        if (VideoViewer.NaturalDuration.HasTimeSpan)
        {
            _knownVideoDuration = VideoViewer.NaturalDuration.TimeSpan;
        }

        UpdateVideoProgressDisplay();
    }

    private void ShowVideoProgressUi()
    {
        _videoProgressTimer.Stop();
        _isScrubbingVideo = false;
        if (VideoProgressPanel is not null)
        {
            VideoProgressPanel.Visibility = Visibility.Visible;
        }

        if (VideoProgressSlider is not null)
        {
            VideoProgressSlider.Minimum = 0;
            VideoProgressSlider.Maximum = 1;
            VideoProgressSlider.Value = 0;
            VideoProgressSlider.IsEnabled = false;
        }

        if (VideoProgressText is not null)
        {
            VideoProgressText.Text = "00:00 / 00:00";
        }

        _knownVideoDuration = TimeSpan.Zero;
        _videoProgressTimer.Start();
    }

    private void HideVideoProgressUi()
    {
        _videoProgressTimer.Stop();
        _isScrubbingVideo = false;
        _knownVideoDuration = TimeSpan.Zero;
        if (VideoProgressPanel is not null)
        {
            VideoProgressPanel.Visibility = Visibility.Collapsed;
        }

        if (VideoProgressSlider is not null)
        {
            VideoProgressSlider.Minimum = 0;
            VideoProgressSlider.Maximum = 1;
            VideoProgressSlider.Value = 0;
            VideoProgressSlider.IsEnabled = false;
        }

        if (VideoProgressText is not null)
        {
            VideoProgressText.Text = "00:00 / 00:00";
        }
    }

    private void UpdateVideoProgressDisplay(TimeSpan? overridePosition = null)
    {
        if (VideoProgressSlider is null || VideoProgressText is null || VideoViewer is null)
        {
            return;
        }

        if (VideoViewer.NaturalDuration.HasTimeSpan)
        {
            _knownVideoDuration = VideoViewer.NaturalDuration.TimeSpan;
        }

        var duration = _knownVideoDuration;
        if (duration <= TimeSpan.Zero)
        {
            VideoProgressSlider.IsEnabled = false;
            VideoProgressSlider.Maximum = 1;
            if (!_isScrubbingVideo)
            {
                VideoProgressSlider.Value = 0;
            }

            VideoProgressText.Text = "00:00 / 00:00";
            return;
        }

        var durationSeconds = duration.TotalSeconds;
        VideoProgressSlider.IsEnabled = true;
        VideoProgressSlider.Maximum = durationSeconds;

        double displaySeconds;
        if (_isScrubbingVideo && !overridePosition.HasValue)
        {
            displaySeconds = Math.Clamp(VideoProgressSlider.Value, 0, durationSeconds);
        }
        else
        {
            var position = overridePosition ?? VideoViewer.Position;
            displaySeconds = Math.Clamp(position.TotalSeconds, 0, durationSeconds);
            if (!_isScrubbingVideo)
            {
                VideoProgressSlider.Value = displaySeconds;
            }
        }

        VideoProgressText.Text = $"{FormatTimestamp(displaySeconds)} / {FormatTimestamp(durationSeconds)}";
    }

    private static string FormatTimestamp(double totalSeconds)
    {
        if (double.IsNaN(totalSeconds) || double.IsInfinity(totalSeconds))
        {
            totalSeconds = 0;
        }

        var clamped = Math.Max(0, totalSeconds);
        var time = TimeSpan.FromSeconds(clamped);
        if (time.TotalHours >= 1)
        {
            return $"{(int)time.TotalHours:00}:{time.Minutes:00}:{time.Seconds:00}";
        }

        return $"{time.Minutes:00}:{time.Seconds:00}";
    }

    private void VideoProgressSlider_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (VideoViewer.Visibility != Visibility.Visible || VideoProgressSlider is null || !VideoProgressSlider.IsEnabled)
        {
            return;
        }

        _isScrubbingVideo = true;
    }

    private void VideoProgressSlider_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isScrubbingVideo)
        {
            return;
        }

        _isScrubbingVideo = false;
        SeekVideoToSlider();
    }

    private void VideoProgressSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (_isScrubbingVideo)
        {
            UpdateVideoProgressDisplay(TimeSpan.FromSeconds(e.NewValue));
        }
    }

    private void VideoProgressSlider_LostMouseCapture(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isScrubbingVideo)
        {
            return;
        }

        _isScrubbingVideo = false;
        SeekVideoToSlider();
    }

    private void VideoProgressSlider_DragStarted(object sender, DragStartedEventArgs e)
    {
        if (VideoProgressSlider is null || !VideoProgressSlider.IsEnabled || VideoViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        _isScrubbingVideo = true;
    }

    private void VideoProgressSlider_DragCompleted(object sender, DragCompletedEventArgs e)
    {
        if (!_isScrubbingVideo)
        {
            return;
        }

        _isScrubbingVideo = false;
        SeekVideoToSlider();
    }

    private void SeekVideoToSlider()
    {
        if (VideoProgressSlider is null || VideoViewer.Visibility != Visibility.Visible || !VideoProgressSlider.IsEnabled)
        {
            return;
        }

        var seconds = Math.Clamp(VideoProgressSlider.Value, 0, VideoProgressSlider.Maximum);
        var position = TimeSpan.FromSeconds(seconds);
        VideoViewer.Position = position;
        if (_isVideoPaused)
        {
            VideoViewer.Pause();
        }

        UpdateVideoProgressDisplay(position);
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
        HideVideoProgressUi();
        ResetVideoZoom();
        UpdateMediaInfo("File: --", "Size: --");
        if (GridViewControl is not null)
        {
            GridViewControl.SelectedIndex = -1;
        }
        UpdateSelectionLabel();
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
            else if (VideoViewer.Visibility == Visibility.Visible && GetCurrentItem()?.Type == MediaType.Video)
            {
                AdjustVideoZoom(e.Delta > 0);
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
            EnsureSlideshowTimerForCurrentMedia();
        }
        else
        {
            _slideshowTimer.Stop();
            _slideshowWaitingForVideoEnd = false;
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

    private void EnsureSlideshowTimerForCurrentMedia()
    {
        EnsureSlideshowTimerForMedia(GetCurrentItem());
    }

    private void EnsureSlideshowTimerForMedia(MediaItem? item)
    {
        if (SlideshowCheckBox?.IsChecked != true)
        {
            _slideshowWaitingForVideoEnd = false;
            return;
        }

        if (item is null)
        {
            _slideshowWaitingForVideoEnd = false;
            _slideshowTimer.Stop();
            return;
        }

        if (item.Type == MediaType.Video)
        {
            _slideshowWaitingForVideoEnd = true;
            if (_slideshowTimer.IsEnabled)
            {
                _slideshowTimer.Stop();
            }
        }
        else
        {
            _slideshowWaitingForVideoEnd = false;
            if (_slideshowTimer.IsEnabled)
            {
                RestartSlideshowTimer();
            }
            else
            {
                _slideshowTimer.Start();
            }
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
        if (TryGetDropPaths(e.Data, out _))
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
        if (!TryGetDropPaths(e.Data, out var droppedPaths))
        {
            return;
        }

        var includeSubfolders = RecursiveCheckBox?.IsChecked == true;
        var newItems = await Task.Run(() => BuildMediaItemsFromSelection(droppedPaths, includeSubfolders));
        if (newItems.Count == 0)
        {
            var folder = droppedPaths.FirstOrDefault(Directory.Exists);
            if (!string.IsNullOrEmpty(folder) && !IsDeleteFolder(folder))
            {
                await LoadFolderAsync(folder);
            }

            return;
        }

        var append = false;
        if (_allItems.Count > 0)
        {
            var result = System.Windows.MessageBox.Show(
                "Add dropped items to the current selection?\nYes = Add, No = Replace, Cancel = Ignore.",
                "Simple Viewer",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question,
                MessageBoxResult.Yes);

            if (result == MessageBoxResult.Cancel)
            {
                return;
            }

            append = result == MessageBoxResult.Yes;
        }

        ApplyCustomItems(newItems, append);
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

    private void VideoViewer_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_videoZoom <= 1.0 || VideoViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        _isPanningVideo = true;
        _panStartPoint = e.GetPosition(this);
        VideoViewer.CaptureMouse();
        e.Handled = true;
    }

    private void VideoViewer_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isPanningVideo || VideoViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        var current = e.GetPosition(this);
        var delta = current - _panStartPoint;
        _panStartPoint = current;
        _videoPanTransform.X += delta.X;
        _videoPanTransform.Y += delta.Y;
    }

    private void VideoViewer_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isPanningVideo)
        {
            return;
        }

        _isPanningVideo = false;
        VideoViewer.ReleaseMouseCapture();
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

    private static bool TryGetDropPaths(System.Windows.IDataObject data, out List<string> paths)
    {
        paths = new List<string>();
        if (!data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            return false;
        }

        if (data.GetData(System.Windows.DataFormats.FileDrop) is not string[] rawPaths)
        {
            return false;
        }

        foreach (var raw in rawPaths)
        {
            var path = raw?.Trim();
            if (string.IsNullOrEmpty(path))
            {
                continue;
            }

            if (Directory.Exists(path) && !IsDeleteFolder(path))
            {
                paths.Add(path);
            }
            else if (File.Exists(path) && IsSupportedFileType(path))
            {
                paths.Add(path);
            }
        }

        return paths.Count > 0;
    }

    private static bool IsSupportedFileType(string path)
    {
        var extension = Path.GetExtension(path);
        if (string.IsNullOrEmpty(extension))
        {
            return false;
        }

        return ImageExtensions.Contains(extension) || VideoExtensions.Contains(extension);
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
        <li>Timeline slider - Click or drag the progress bar to seek</li>
        <li>Ctrl + Mouse Wheel - Zoom the current video</li>
        <li>Drag (while zoomed) - Pan the video frame</li>
    </ul>
</div>
<div>
    <h2>Release Highlights</h2>
    <h3>v1.2.0</h3>
    <ul>
        <li><strong>Video timeline</strong>
            <ul>
                <li>New scrub bar shows elapsed/total time and supports precise seeking.</li>
                <li>Slider updates while looping videos and respects pause/mute states.</li>
            </ul>
        </li>
        <li><strong>Video zoom & pan</strong>
            <ul>
                <li>Ctrl + mouse wheel now zooms into videos, just like images.</li>
                <li>Click-drag pans the video when zoomed in for pixel-level inspection.</li>
            </ul>
        </li>
        <li><strong>Video metadata</strong>
            <ul>
                <li>Metadata panel now reads duration, codecs, bitrate, and tags from video files.</li>
                <li>Powered by TagLib# for broad format compatibility.</li>
            </ul>
        </li>
    </ul>
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
        <li><strong>Custom selections & Explorer</strong>
            <ul>
                <li>Drag folders/files into the window and choose Add vs Replace when a session is active.</li>
                <li>Command-line inputs accept folders/files with an optional <code>--add</code> switch to append selections.</li>
                <li>Options &rarr; Explorer integration installs/removes right-click entries for supported folders and media files.</li>
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

        var text = GetOrLoadMetadata(item);
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

    private string GetOrLoadMetadata(MediaItem item)
    {
        if (_metadataCache.TryGetValue(item.Path, out var cached))
        {
            return cached;
        }

        var metadata = item.Type switch
        {
            MediaType.Video => ExtractVideoMetadata(item.Path),
            MediaType.Image => ExtractImageMetadata(item.Path),
            _ => "No metadata available for this media type."
        };

        _metadataCache[item.Path] = metadata;
        return metadata;
    }

    private static string ExtractImageMetadata(string path)
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

    private static string ExtractVideoMetadata(string path)
    {
        var summaryEntries = new List<(string Key, string Value)>();
        var seenSummaryKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var summaryNotes = new List<string>();
        var sections = new List<string>();

        try
        {
            using var tagFile = TagLibFile.Create(path);
            var builder = new StringBuilder();
            builder.AppendLine($"File: {Path.GetFileName(path)}");
            if (File.Exists(path))
            {
                var info = new FileInfo(path);
                builder.AppendLine($"File Size: {FormatFileSize(info.Length)}");
            }

            var properties = tagFile.Properties;
            if (properties is not null)
            {
                if (properties.Duration > TimeSpan.Zero)
                {
                    var durationText = FormatTimestamp(properties.Duration.TotalSeconds);
                    builder.AppendLine($"Duration: {durationText}");
                    AddSummaryEntry(summaryEntries, seenSummaryKeys, "Duration", durationText);
                }

                if (properties.VideoWidth > 0 && properties.VideoHeight > 0)
                {
                    var videoSize = $"{properties.VideoWidth} x {properties.VideoHeight}";
                    builder.AppendLine($"Video Size: {videoSize}");
                    AddSummaryEntry(summaryEntries, seenSummaryKeys, "Size", videoSize);
                }

                if (properties.AudioBitrate > 0)
                {
                    builder.AppendLine($"Audio Bitrate: {properties.AudioBitrate} kbps");
                }

                if (properties.AudioSampleRate > 0)
                {
                    builder.AppendLine($"Audio Sample Rate: {properties.AudioSampleRate} Hz");
                }

                var videoCodecs = properties.Codecs?
                    .Where(codec => codec.MediaTypes.HasFlag(TagLibMediaTypes.Video))
                    .Select(codec => codec.Description)
                    .Where(desc => !string.IsNullOrWhiteSpace(desc))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (videoCodecs is { Count: > 0 })
                {
                    var codecText = string.Join(", ", videoCodecs);
                    builder.AppendLine($"Video Codec(s): {codecText}");
                    AddSummaryEntry(summaryEntries, seenSummaryKeys, "Video Codec", codecText);
                }

                var audioCodecs = properties.Codecs?
                    .Where(codec => codec.MediaTypes.HasFlag(TagLibMediaTypes.Audio))
                    .Select(codec => codec.Description)
                    .Where(desc => !string.IsNullOrWhiteSpace(desc))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (audioCodecs is { Count: > 0 })
                {
                    var codecText = string.Join(", ", audioCodecs);
                    builder.AppendLine($"Audio Codec(s): {codecText}");
                    AddSummaryEntry(summaryEntries, seenSummaryKeys, "Audio Codec", codecText);
                }
            }

            var tag = tagFile.Tag;
            var tagSection = new StringBuilder();
            if (tag is not null)
            {
                AppendMetadataField(tagSection, "Title", tag.Title);
                AppendMetadataField(tagSection, "Album", tag.Album);
                if (tag.Year > 0)
                {
                    AppendMetadataField(tagSection, "Year", tag.Year.ToString());
                }

                var performers = tag.Performers?.Where(p => !string.IsNullOrWhiteSpace(p)).ToArray();
                if (performers is { Length: > 0 })
                {
                    AppendMetadataField(tagSection, "Performers", string.Join(", ", performers));
                }

                var albumArtists = tag.AlbumArtists?.Where(a => !string.IsNullOrWhiteSpace(a)).ToArray();
                if (albumArtists is { Length: > 0 })
                {
                    AppendMetadataField(tagSection, "Album Artists", string.Join(", ", albumArtists));
                }

                var genres = tag.Genres?.Where(g => !string.IsNullOrWhiteSpace(g)).ToArray();
                if (genres is { Length: > 0 })
                {
                    AppendMetadataField(tagSection, "Genres", string.Join(", ", genres));
                }

                AppendMetadataField(tagSection, "Comment", tag.Comment);
                if (tag.Track > 0)
                {
                    var trackLabel = tag.TrackCount > 0 ? $"{tag.Track} / {tag.TrackCount}" : tag.Track.ToString();
                    AppendMetadataField(tagSection, "Track", trackLabel);
                }
            }

            if (tagSection.Length > 0)
            {
                builder.AppendLine();
                builder.AppendLine("Tag metadata:");
                builder.Append(tagSection.ToString().TrimEnd());
            }

            if (builder.Length > 0)
            {
                sections.Add(builder.ToString().Trim());
            }
        }
        catch (Exception ex)
        {
            return $"Unable to read metadata from the video file.\n{ex.Message}";
        }

        var comfyWorkflow = TryExtractComfyWorkflowFromVideo(path, summaryEntries, seenSummaryKeys, summaryNotes);
        if (!string.IsNullOrWhiteSpace(comfyWorkflow))
        {
            sections.Add($"Embedded workflow metadata:{Environment.NewLine}{comfyWorkflow}");
        }

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

            sections.Insert(0, summaryBuilder.ToString().TrimEnd());
        }

        if (sections.Count == 0)
        {
            return "No metadata found.";
        }

        return string.Join($"{Environment.NewLine}{Environment.NewLine}", sections);
    }

    private static string FormatFileSize(long bytes)
    {
        if (bytes < 0)
        {
            bytes = 0;
        }

        string[] units = { "B", "KB", "MB", "GB", "TB" };
        double value = bytes;
        var unitIndex = 0;
        while (value >= 1024 && unitIndex < units.Length - 1)
        {
            value /= 1024;
            unitIndex++;
        }

        return $"{value:0.##} {units[unitIndex]}";
    }

    private static void AppendMetadataField(StringBuilder builder, string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        builder.AppendLine($"{label}: {value}");
    }

    private static string? TryExtractComfyWorkflowFromVideo(
        string path,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys,
        List<string> summaryNotes)
    {
        const int MaxScanBytes = 4 * 1024 * 1024;
        try
        {
            if (!File.Exists(path))
            {
                return null;
            }

            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var length = stream.Length;
            var bytesToRead = (int)Math.Min(length, MaxScanBytes);
            var buffer = new byte[bytesToRead];
            if (length > bytesToRead)
            {
                stream.Seek(length - bytesToRead, SeekOrigin.Begin);
            }

            var read = stream.Read(buffer, 0, buffer.Length);
            if (read <= 0)
            {
                return null;
            }

            var text = Encoding.UTF8.GetString(buffer, 0, read);
            var keywords = new[] { "\"workflow\"", "\"Workflow\"", "\"prompt\"", "\"Prompt\"", "\"nodes\"" };
            foreach (var keyword in keywords)
            {
                var json = ExtractJsonBlockContainingKeyword(text, keyword);
                if (!string.IsNullOrWhiteSpace(json))
                {
                    TryCollectSummaryFromJson(json, summaryEntries, seenSummaryKeys);
                    return TryFormatJson(json);
                }
            }

            return null;
        }
        catch
        {
            summaryNotes.Add("Unable to scan for embedded workflow metadata.");
            return null;
        }
    }

    private static string? ExtractJsonBlockContainingKeyword(string text, string keyword)
    {
        var index = text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase);
        if (index < 0)
        {
            return null;
        }

        var start = text.LastIndexOf('{', index);
        if (start < 0)
        {
            return null;
        }

        var depth = 0;
        var inString = false;
        for (var i = start; i < text.Length; i++)
        {
            var ch = text[i];
            if (ch == '"' && (i == 0 || text[i - 1] != '\\'))
            {
                inString = !inString;
            }

            if (inString)
            {
                continue;
            }

            if (ch == '{')
            {
                depth++;
            }
            else if (ch == '}')
            {
                depth--;
                if (depth == 0)
                {
                    return text.Substring(start, i - start + 1);
                }
            }
        }

        return null;
    }

    private static string TryFormatJson(string json)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            return JsonSerializer.Serialize(document.RootElement, new JsonSerializerOptions
            {
                WriteIndented = true
            });
        }
        catch
        {
            return json;
        }
    }

    private static void TryCollectSummaryFromJson(
        string json,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys)
    {
        try
        {
            using var document = JsonDocument.Parse(json);
            CollectSummaryFromJsonElement(document.RootElement, string.Empty, summaryEntries, seenSummaryKeys);
        }
        catch
        {
            // Ignore JSON parsing failures for summary generation.
        }
    }

    private static void CollectSummaryFromJsonElement(
        JsonElement element,
        string path,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in element.EnumerateObject())
                {
                    var childPath = string.IsNullOrEmpty(path)
                        ? property.Name
                        : $"{path}.{property.Name}";
                    CollectSummaryFromJsonElement(property.Value, childPath, summaryEntries, seenSummaryKeys);
                }

                break;
            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in element.EnumerateArray())
                {
                    CollectSummaryFromJsonElement(item, $"{path}[{index}]", summaryEntries, seenSummaryKeys);
                    index++;
                }

                break;
            case JsonValueKind.String:
                AddSummaryEntryFromJsonPath(path, element.GetString(), summaryEntries, seenSummaryKeys);
                break;
            case JsonValueKind.Number:
                AddSummaryEntryFromJsonPath(path, element.GetRawText(), summaryEntries, seenSummaryKeys);
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
                AddSummaryEntryFromJsonPath(path, element.GetBoolean().ToString(), summaryEntries, seenSummaryKeys);
                break;
        }
    }

    private static void AddSummaryEntryFromJsonPath(
        string path,
        string? value,
        List<(string Key, string Value)> summaryEntries,
        HashSet<string> seenSummaryKeys)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        if (TryMapSummaryKey(path, out var mappedLabel))
        {
            AddSummaryEntry(summaryEntries, seenSummaryKeys, mappedLabel, value);
            return;
        }

        if (path.Contains("prompt", StringComparison.OrdinalIgnoreCase) ||
            path.Contains("workflow", StringComparison.OrdinalIgnoreCase))
        {
            AddSummaryEntry(summaryEntries, seenSummaryKeys, path, value);
        }
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
