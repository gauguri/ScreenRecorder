using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Forms = System.Windows.Forms;

namespace ScreenRecorder
{
    public partial class MainWindow : Window
    {
        private const int FramesPerSecond = 30;

        private readonly SemaphoreSlim _stopSemaphore = new(1, 1);
        private CancellationTokenSource? _recordingCancellationSource;
        private Task? _recordingTask;
        private bool _isRecording;
        private string? _activeRecordingFilePath;
        private string? _lastCompletedRecording;

        public MainWindow()
        {
            InitializeComponent();
            Closing += MainWindow_OnClosing;
            UpdateUiState();
        }

        private void StartRecordingButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isRecording)
            {
                return;
            }

            try
            {
                var filePath = CreateRecordingFilePath();
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                _recordingCancellationSource = new CancellationTokenSource();
                _activeRecordingFilePath = filePath;
                _isRecording = true;

                StatusTextBlock.Text = "Recording in progress...";
                UpdateUiState();

                _recordingTask = Task.Run(() => RecordScreen(filePath, _recordingCancellationSource.Token), _recordingCancellationSource.Token);
                _recordingTask.ContinueWith(_ => _ = StopRecordingAsync(), CancellationToken.None, TaskContinuationOptions.OnlyOnFaulted, TaskScheduler.Default);
            }
            catch (Exception ex)
            {
                _isRecording = false;
                _activeRecordingFilePath = null;
                _recordingTask = null;
                _recordingCancellationSource?.Dispose();
                _recordingCancellationSource = null;

                UpdateUiState();
                StatusTextBlock.Text = "Recording failed to start.";
                MessageBox.Show(this, ex.Message, "Unable to start recording", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void StopRecordingButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isRecording)
            {
                return;
            }

            StatusTextBlock.Text = "Stopping recording...";
            StopButton.IsEnabled = false;
            await StopRecordingAsync();
        }

        private async Task StopRecordingAsync()
        {
            await _stopSemaphore.WaitAsync();
            try
            {
                if (!_isRecording)
                {
                    return;
                }

                _recordingCancellationSource?.Cancel();

                Exception? failure = null;
                var recordingTask = _recordingTask;
                if (recordingTask is not null)
                {
                    try
                    {
                        await recordingTask.ConfigureAwait(false);
                    }
                    catch (OperationCanceledException)
                    {
                    }
                    catch (Exception ex)
                    {
                        failure = ex;
                    }
                }

                _recordingTask = null;
                _recordingCancellationSource?.Dispose();
                _recordingCancellationSource = null;

                var finishedFile = _activeRecordingFilePath;
                _activeRecordingFilePath = null;
                _isRecording = false;

                if (Dispatcher.CheckAccess())
                {
                    FinalizeStop(failure, finishedFile);
                }
                else
                {
                    await Dispatcher.InvokeAsync(() => FinalizeStop(failure, finishedFile));
                }
            }
            finally
            {
                _stopSemaphore.Release();
            }
        }

        private void FinalizeStop(Exception? failure, string? finishedFile)
        {
            if (failure is not null)
            {
                StatusTextBlock.Text = $"Recording failed: {failure.Message}";
                MessageBox.Show(this, failure.Message, "Recording failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else if (!string.IsNullOrEmpty(finishedFile) && File.Exists(finishedFile))
            {
                _lastCompletedRecording = finishedFile;
                StatusTextBlock.Text = $"Recording saved to:\n{finishedFile}";
            }
            else
            {
                StatusTextBlock.Text = "Recording stopped.";
            }

            UpdateUiState();
        }

        private void RecordScreen(string filePath, CancellationToken token)
        {
            var virtualScreen = Forms.SystemInformation.VirtualScreen;
            var width = virtualScreen.Width - virtualScreen.Width % 2;
            var height = virtualScreen.Height - virtualScreen.Height % 2;

            if (width <= 0 || height <= 0)
            {
                throw new InvalidOperationException("Unable to determine a valid screen size for recording.");
            }

            var captureRect = new Rectangle(virtualScreen.Left, virtualScreen.Top, width, height);
            var captureSize = captureRect.Size;
            var videoSize = new OpenCvSharp.Size(captureSize.Width, captureSize.Height);

            using var bitmap = new Bitmap(captureSize.Width, captureSize.Height, PixelFormat.Format32bppArgb);
            using var graphics = Graphics.FromImage(bitmap);
            using var writer = new VideoWriter(filePath, FourCC.MP4V, FramesPerSecond, videoSize);

            if (!writer.IsOpened())
            {
                throw new InvalidOperationException("Failed to initialize the video writer. Ensure the necessary codecs are available.");
            }

            var frameInterval = TimeSpan.FromSeconds(1d / FramesPerSecond);
            var stopwatch = new Stopwatch();

            while (!token.IsCancellationRequested)
            {
                stopwatch.Restart();

                graphics.CopyFromScreen(captureRect.Left, captureRect.Top, 0, 0, captureSize, CopyPixelOperation.SourceCopy);
                using (var frame = BitmapConverter.ToMat(bitmap))
                {
                    Cv2.CvtColor(frame, frame, ColorConversionCodes.BGRA2BGR);
                    writer.Write(frame);
                }

                stopwatch.Stop();
                var remaining = frameInterval - stopwatch.Elapsed;
                if (remaining > TimeSpan.Zero)
                {
                    if (token.WaitHandle.WaitOne(remaining))
                    {
                        break;
                    }
                }
            }

            writer.Release();
        }

        private static string CreateRecordingFilePath()
        {
            var videosDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);
            var outputDirectory = Path.Combine(videosDirectory, "ScreenRecorder");
            var fileName = $"ScreenRecording_{DateTime.Now:yyyyMMdd_HHmmss}.mp4";
            return Path.Combine(outputDirectory, fileName);
        }

        private void UpdateUiState()
        {
            StartButton.IsEnabled = !_isRecording;
            StopButton.IsEnabled = _isRecording;

            if (_isRecording && !string.IsNullOrEmpty(_activeRecordingFilePath))
            {
                FilePathTextBlock.Text = _activeRecordingFilePath;
            }
            else if (!string.IsNullOrEmpty(_lastCompletedRecording) && File.Exists(_lastCompletedRecording))
            {
                FilePathTextBlock.Text = _lastCompletedRecording;
            }
            else
            {
                FilePathTextBlock.Text = "No recordings yet.";
            }

            OpenLocationButton.IsEnabled = !_isRecording && !string.IsNullOrEmpty(_lastCompletedRecording) && File.Exists(_lastCompletedRecording);
        }

        private void OpenLocationButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_lastCompletedRecording) || !File.Exists(_lastCompletedRecording))
            {
                return;
            }

            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{_lastCompletedRecording}\"",
                    UseShellExecute = true
                };
                Process.Start(processInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Unable to open folder", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void MainWindow_OnClosing(object? sender, CancelEventArgs e)
        {
            if (_isRecording)
            {
                e.Cancel = true;
                await StopRecordingAsync();
                Close();
            }
        }
    }
}
