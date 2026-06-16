using System.Windows;
using System.Windows.Input;

namespace PptxViewer;

public partial class FullscreenWindow : Window
{
    private readonly List<SlideData> _slides;
    private int _current;

    public event Action<int>? SlideChanged;

    public FullscreenWindow(List<SlideData> slides, int startIndex)
    {
        InitializeComponent();
        _slides = slides;
        _current = Math.Clamp(startIndex, 0, slides.Count - 1);
        ShowSlide(_current);
    }

    private void ShowSlide(int index)
    {
        PptxParser.RenderSlide(_slides[index], FsCanvas);
        SlideCounter.Text = $"{index + 1} / {_slides.Count}";
        SlideChanged?.Invoke(index);
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Escape:
                Close();
                break;

            case Key.Right:
            case Key.Down:
            case Key.PageDown:
            case Key.Space:
                if (_current < _slides.Count - 1)
                    ShowSlide(++_current);
                break;

            case Key.Left:
            case Key.Up:
            case Key.PageUp:
                if (_current > 0)
                    ShowSlide(--_current);
                break;

            case Key.Home:
                ShowSlide(_current = 0);
                break;

            case Key.End:
                ShowSlide(_current = _slides.Count - 1);
                break;
        }
    }
}
