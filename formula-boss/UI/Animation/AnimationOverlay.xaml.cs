using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace FormulaBoss.UI.Animation;

public partial class AnimationOverlay
{
    private readonly List<SpriteFrame> _frames;
    private readonly List<WriteableBitmap> _bitmaps;
    private readonly DispatcherTimer _timer;
    private readonly double _baseIntervalMs;
    private int _currentFrame;
    private Storyboard? _shakeStoryboard;
    private bool _fading;

    public AnimationOverlay(List<SpriteFrame> frames, double baseIntervalMs = 120)
    {
        InitializeComponent();

        ShowActivated = false;

        _frames = frames;
        _baseIntervalMs = baseIntervalMs;
        _bitmaps = frames.Select(f => SpriteRenderer.Render(f.Grid)).ToList();

        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(baseIntervalMs) };
        _timer.Tick += OnTick;

        BuildShakeStoryboard();
    }

    /// <summary>
    ///     Plays the animation once, then closes the window.
    /// </summary>
    public void PlayOnce()
    {
        _currentFrame = 0;
        ShowFrame(0);
        Show();
        _timer.Start();
    }

    /// <summary>
    ///     Plays the animation in a loop until closed.
    /// </summary>
    public void PlayLoop()
    {
        _currentFrame = 0;
        ShowFrame(0);
        Show();
        _timer.Start();
    }

    /// <summary>
    ///     Set to true to close after one cycle, false for looping.
    /// </summary>
    public bool OneShot { get; set; }

    /// <summary>
    ///     Begins fading the overlay to transparent, then closes it.
    /// </summary>
    public void BeginFadeOut(double durationMs = 500)
    {
        if (_fading)
        {
            return;
        }

        _fading = true;

        var fade = new DoubleAnimation(Opacity, 0, TimeSpan.FromMilliseconds(durationMs))
        {
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
        };
        fade.Completed += (_, _) =>
        {
            _timer.Stop();
            _shakeStoryboard?.Stop();
            Close();
        };
        BeginAnimation(OpacityProperty, fade);
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);

        // Make window non-focusable so it never steals focus from Excel
        var hwnd = new WindowInteropHelper(this).Handle;
        var exStyle = NativeMethods.GetWindowLong(hwnd, NativeMethods.GwlExStyle);
        NativeMethods.SetWindowLong(hwnd, NativeMethods.GwlExStyle,
            exStyle | NativeMethods.WsExNoActivate | NativeMethods.WsExToolWindow);
    }

    private void OnTick(object? sender, EventArgs e)
    {
        _currentFrame++;

        if (_currentFrame >= _frames.Count)
        {
            if (OneShot)
            {
                _timer.Stop();
                _shakeStoryboard?.Stop();
                if (!_fading)
                {
                    Close();
                }

                return;
            }

            _currentFrame = 0;
        }

        ShowFrame(_currentFrame);
    }

    private void ShowFrame(int index)
    {
        SpriteImage.Source = _bitmaps[index];

        var frame = _frames[index];
        _timer.Interval = TimeSpan.FromMilliseconds(_baseIntervalMs * frame.DurationMultiplier);

        if (frame.Shake)
        {
            _shakeStoryboard?.Begin();
        }
        else
        {
            _shakeStoryboard?.Stop();
        }
    }

    private void BuildShakeStoryboard()
    {
        var translateTransform = new TranslateTransform();
        ShakeBorder.RenderTransform = translateTransform;

        var duration = TimeSpan.FromMilliseconds(60);

        var shakeX = new DoubleAnimationUsingKeyFrames
        {
            Duration = duration,
            RepeatBehavior = RepeatBehavior.Forever,
        };
        shakeX.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromPercent(0)));
        shakeX.KeyFrames.Add(new LinearDoubleKeyFrame(-3, KeyTime.FromPercent(0.25)));
        shakeX.KeyFrames.Add(new LinearDoubleKeyFrame(3, KeyTime.FromPercent(0.5)));
        shakeX.KeyFrames.Add(new LinearDoubleKeyFrame(-2, KeyTime.FromPercent(0.75)));
        shakeX.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromPercent(1)));

        var shakeY = new DoubleAnimationUsingKeyFrames
        {
            Duration = duration,
            RepeatBehavior = RepeatBehavior.Forever,
        };
        shakeY.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromPercent(0)));
        shakeY.KeyFrames.Add(new LinearDoubleKeyFrame(2, KeyTime.FromPercent(0.25)));
        shakeY.KeyFrames.Add(new LinearDoubleKeyFrame(-2, KeyTime.FromPercent(0.5)));
        shakeY.KeyFrames.Add(new LinearDoubleKeyFrame(-3, KeyTime.FromPercent(0.75)));
        shakeY.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromPercent(1)));

        Storyboard.SetTarget(shakeX, ShakeBorder);
        Storyboard.SetTargetProperty(shakeX, new PropertyPath("RenderTransform.(TranslateTransform.X)"));
        Storyboard.SetTarget(shakeY, ShakeBorder);
        Storyboard.SetTargetProperty(shakeY, new PropertyPath("RenderTransform.(TranslateTransform.Y)"));

        _shakeStoryboard = new Storyboard();
        _shakeStoryboard.Children.Add(shakeX);
        _shakeStoryboard.Children.Add(shakeY);
    }
}
