using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace PptxViewer;

public partial class MainWindow : Window
{
    private List<SlideData> _slides = new();
    private int _currentIndex = -1;
    private double _zoom = 1.0;
    private bool _notesVisible = false;
    private string? _currentFile;

    // Thumbnail view model
    private record ThumbItem(int Index, BitmapSource Thumbnail);

    public MainWindow()
    {
        InitializeComponent();
        AllowDrop = true;
        Drop += Window_Drop;
        DragOver += Window_DragOver;
        UpdateNav();
    }

    // ══════════════ FILE OPEN ══════════════

    private void OpenFile_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "PowerPoint Files|*.pptx|All Files|*.*",
            Title = "Open PPTX File"
        };
        if (dlg.ShowDialog() == true)
            LoadFile(dlg.FileName);
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            LoadFile(files[0]);
    }

    private void LoadFile(string path)
    {
        if (!path.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show("Please open a .pptx file.", "Invalid File",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            StatusText.Text = "Loading…";
            _currentFile = path;
            FilePath.Text = Path.GetFileName(path);

            _slides = PptxParser.Parse(path);

            // Build thumbnails
            ThumbnailList.Items.Clear();
            foreach (var slide in _slides)
            {
                var thumb = PptxParser.RenderThumbnail(slide);
                ThumbnailList.Items.Add(new ThumbItem(slide.Index, thumb));
            }

            TotalSlides.Text = _slides.Count.ToString();
            Title = $"PPTX Viewer — {Path.GetFileName(path)}";

            if (_slides.Count > 0)
            {
                _currentIndex = 0;
                ShowSlide(0);
                ThumbnailList.SelectedIndex = 0;
                DropHint.Visibility = Visibility.Collapsed;
                SlideContainer.Visibility = Visibility.Visible;
            }

            StatusText.Text = $"Loaded {_slides.Count} slides";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Error loading file";
            MessageBox.Show($"Failed to load PPTX:\n{ex.Message}", "Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ══════════════ SLIDE DISPLAY ══════════════

    private void ShowSlide(int index)
    {
        if (index < 0 || index >= _slides.Count) return;
        _currentIndex = index;

        var slide = _slides[index];
        PptxParser.RenderSlide(slide, SlideCanvas);
        ApplyZoom();

        SlideNumberBox.Text = (index + 1).ToString();
        NotesText.Text = string.IsNullOrEmpty(slide.Notes)
            ? "(No speaker notes)"
            : slide.Notes;

        StatusText.Text = $"Slide {index + 1} of {_slides.Count}";
        UpdateNav();
    }

    private void ApplyZoom()
    {
        SlideCanvas.LayoutTransform = new ScaleTransform(_zoom, _zoom);
        ZoomLabel.Text = $"{(int)(_zoom * 100)}%";
    }

    // ══════════════ NAVIGATION ══════════════

    private void PrevSlide_Click(object sender, RoutedEventArgs e)
    {
        if (_currentIndex > 0)
        {
            _currentIndex--;
            ShowSlide(_currentIndex);
            ThumbnailList.SelectedIndex = _currentIndex;
            ThumbnailList.ScrollIntoView(ThumbnailList.Items[_currentIndex]);
        }
    }

    private void NextSlide_Click(object sender, RoutedEventArgs e)
    {
        if (_currentIndex < _slides.Count - 1)
        {
            _currentIndex++;
            ShowSlide(_currentIndex);
            ThumbnailList.SelectedIndex = _currentIndex;
            ThumbnailList.ScrollIntoView(ThumbnailList.Items[_currentIndex]);
        }
    }

    private void ThumbnailList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ThumbnailList.SelectedIndex >= 0 && ThumbnailList.SelectedIndex != _currentIndex)
            ShowSlide(ThumbnailList.SelectedIndex);
    }

    private void SlideNumberBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            if (int.TryParse(SlideNumberBox.Text, out int n))
            {
                n = Math.Clamp(n - 1, 0, _slides.Count - 1);
                ShowSlide(n);
                ThumbnailList.SelectedIndex = n;
                ThumbnailList.ScrollIntoView(ThumbnailList.Items[n]);
            }
        }
    }

    private void UpdateNav()
    {
        BtnPrev.IsEnabled = _currentIndex > 0;
        BtnNext.IsEnabled = _currentIndex < _slides.Count - 1;
    }

    // ══════════════ KEYBOARD SHORTCUTS ══════════════

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (_slides.Count == 0) return;

        switch (e.Key)
        {
            case Key.Right:
            case Key.Down:
            case Key.PageDown:
            case Key.Space:
                NextSlide_Click(this, new RoutedEventArgs());
                break;

            case Key.Left:
            case Key.Up:
            case Key.PageUp:
                PrevSlide_Click(this, new RoutedEventArgs());
                break;

            case Key.Home:
                ShowSlide(0);
                ThumbnailList.SelectedIndex = 0;
                break;

            case Key.End:
                ShowSlide(_slides.Count - 1);
                ThumbnailList.SelectedIndex = _slides.Count - 1;
                break;

            case Key.F5:
                Fullscreen_Click(this, new RoutedEventArgs());
                break;

            case Key.Escape:
                if (WindowStyle == WindowStyle.None)
                    ExitFullscreen();
                break;

            case Key.OemPlus:
            case Key.Add:
                ZoomIn_Click(this, new RoutedEventArgs());
                break;

            case Key.OemMinus:
            case Key.Subtract:
                ZoomOut_Click(this, new RoutedEventArgs());
                break;
        }
    }

    // ══════════════ ZOOM ══════════════

    private void ZoomIn_Click(object sender, RoutedEventArgs e)
    {
        _zoom = Math.Min(_zoom + 0.1, 3.0);
        ApplyZoom();
    }

    private void ZoomOut_Click(object sender, RoutedEventArgs e)
    {
        _zoom = Math.Max(_zoom - 0.1, 0.2);
        ApplyZoom();
    }

    private void ZoomReset_Click(object sender, RoutedEventArgs e)
    {
        _zoom = 1.0;
        ApplyZoom();
    }

    // ══════════════ NOTES PANEL ══════════════

    private void ToggleNotes_Click(object sender, RoutedEventArgs e)
    {
        _notesVisible = !_notesVisible;
        NotesPanel.Visibility = _notesVisible ? Visibility.Visible : Visibility.Collapsed;
    }

    // ══════════════ FULLSCREEN SLIDESHOW ══════════════

    private WindowStyle _prevStyle;
    private WindowState _prevState;

    private void Fullscreen_Click(object sender, RoutedEventArgs e)
    {
        if (_slides.Count == 0) return;

        _prevStyle = WindowStyle;
        _prevState = WindowState;

        WindowStyle = WindowStyle.None;
        WindowState = WindowState.Maximized;

        // Launch dedicated fullscreen window
        var fs = new FullscreenWindow(_slides, _currentIndex);
        fs.SlideChanged += idx =>
        {
            _currentIndex = idx;
            ShowSlide(idx);
            ThumbnailList.SelectedIndex = idx;
        };
        fs.ShowDialog();
    }

    private void ExitFullscreen()
    {
        WindowStyle = _prevStyle;
        WindowState = _prevState;
    }
}
