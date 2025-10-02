using System;
using System.Collections.ObjectModel;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Media;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Win32;
using Qiosk.App.Configuration;
using Qiosk.App.Models;
using Qiosk.App.Services.Contracts;

namespace Qiosk.App.ViewModels;

public partial class ShellViewModel : ObservableObject, IDisposable
{
    private readonly IAttendeeRepository _attendeeRepository;
    private readonly IQrScannerService _qrScannerService;
    private readonly IKioskSettingsService _settingsService;
    private readonly IOptionsMonitor<KioskOptions> _optionsMonitor;
    private readonly IBadgePrinter _badgePrinter;
    private readonly ILogger<ShellViewModel> _logger;
    private readonly Dispatcher _dispatcher;
    private readonly DispatcherTimer _statusTimer;

    private KioskSettings? _settings;
    private bool _initialized;
    private bool _disposed;

    public ShellViewModel(
        IAttendeeRepository attendeeRepository,
        IQrScannerService qrScannerService,
        IKioskSettingsService settingsService,
        IOptionsMonitor<KioskOptions> optionsMonitor,
        IBadgePrinter badgePrinter,
        ILogger<ShellViewModel> logger)
    {
        _attendeeRepository = attendeeRepository;
        _qrScannerService = qrScannerService;
        _settingsService = settingsService;
        _optionsMonitor = optionsMonitor;
        _badgePrinter = badgePrinter;
        _logger = logger;
        _dispatcher = Dispatcher.CurrentDispatcher;

        Cameras = new ObservableCollection<CameraDevice>();
        Printers = new ObservableCollection<string>();

        _statusTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(_optionsMonitor.CurrentValue.PrintingMessageSeconds)
        };
        _statusTimer.Tick += OnStatusTimerTick;
    }

    public ObservableCollection<CameraDevice> Cameras { get; }
    public ObservableCollection<string> Printers { get; }

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private BitmapSource? _cameraFrame;

    [ObservableProperty]
    private Attendee? _currentAttendee;

    [ObservableProperty]
    private bool _isConfirmationVisible;

    [ObservableProperty]
    private bool _isPrintingVisible;

    [ObservableProperty]
    private bool _isStatusVisible;

    [ObservableProperty]
    private bool _isStatusError;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    [ObservableProperty]
    private CameraDevice? _selectedCamera;

    [ObservableProperty]
    private string? _selectedPrinter;

    [ObservableProperty]
    private bool _isLockMode;

    [ObservableProperty]
    private bool _isAdminVisible;

    [ObservableProperty]
    private bool _isAdminAuthenticated;

    [ObservableProperty]
    private string _adminUserInput = string.Empty;

    [ObservableProperty]
    private string _adminPasswordInput = string.Empty;

    [ObservableProperty]
    private string _adminErrorMessage = string.Empty;

    [ObservableProperty]
    private string _excelPathInput = string.Empty;

    [ObservableProperty]
    private string _templatePathInput = string.Empty;

    partial void OnSelectedCameraChanged(CameraDevice? value)
    {
        if (!_initialized || value is null || _disposed)
        {
            return;
        }

        _ = ApplyCameraSelectionAsync(value);
    }

    partial void OnSelectedPrinterChanged(string? value)
    {
        if (!_initialized || _settings is null)
        {
            return;
        }

        _settings.DefaultPrinter = value ?? string.Empty;
        _ = PersistSettingsAsync();
    }

    partial void OnIsLockModeChanged(bool value)
    {
        if (!_initialized || _settings is null)
        {
            return;
        }

        _settings.LockMode = value;
        _ = PersistSettingsAsync();
    }

    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        if (_initialized)
        {
            return;
        }

        _initialized = true;
        _qrScannerService.FrameReady += HandleFrameReady;
        _qrScannerService.CodeDetected += HandleCodeDetected;

        await LoadSettingsAsync(cancellationToken).ConfigureAwait(false);
        await LoadPrintersAsync().ConfigureAwait(false);
        await LoadAttendeesAsync(cancellationToken).ConfigureAwait(false);
        await StartScannerAsync(cancellationToken).ConfigureAwait(false);

        ShowStatus("Scanner pregatit", false, TimeSpan.FromSeconds(3));
    }

    [RelayCommand]
    private async Task ConfirmCurrentAsync()
    {
        if (CurrentAttendee is null)
        {
            return;
        }

        var attendee = CurrentAttendee;
        IsConfirmationVisible = false;
        IsPrintingVisible = true;
        _qrScannerService.RequestPause(TimeSpan.FromSeconds(_optionsMonitor.CurrentValue.PrintingMessageSeconds));
        ShowStatus("Printare in curs...", false, TimeSpan.FromSeconds(_optionsMonitor.CurrentValue.PrintingMessageSeconds));

        var templatePath = GetDisplayPath(_settings?.TemplatePathOverride, _optionsMonitor.CurrentValue.TemplatePath);
        if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
        {
            SystemSounds.Hand.Play();
            ShowStatus("Sablonul DOCX nu a fost gasit. Verifica setarile din panoul admin.", true, TimeSpan.FromSeconds(10));
            IsPrintingVisible = false;
            CurrentAttendee = null;
            return;
        }

        try
        {
            var printerName = SelectedPrinter ?? string.Empty;
            await _badgePrinter.PrintAsync(attendee, printerName, templatePath);
            SystemSounds.Asterisk.Play();
            ShowStatus("Ecusonul a fost trimis la imprimanta.", false, TimeSpan.FromSeconds(_optionsMonitor.CurrentValue.PrintingMessageSeconds));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Eroare la printarea ecusonului.");
            SystemSounds.Hand.Play();
            ShowStatus("Nu s-a putut imprima ecusonul.", true, TimeSpan.FromSeconds(10));
        }
        finally
        {
            await Task.Delay(TimeSpan.FromSeconds(_optionsMonitor.CurrentValue.PrintingMessageSeconds));
            IsPrintingVisible = false;
            CurrentAttendee = null;
        }
    }

    [RelayCommand]
    private void CancelCurrent()
    {
        CurrentAttendee = null;
        IsConfirmationVisible = false;
        IsPrintingVisible = false;
        ShowStatus("Pregatit pentru urmatorul cod.", false, TimeSpan.FromSeconds(3));
        _qrScannerService.RequestPause(TimeSpan.FromSeconds(0.5));
    }

    [RelayCommand]
    private void OpenAdmin()
    {
        if (_disposed)
        {
            return;
        }

        InitializeAdminFields();
        IsAdminAuthenticated = false;
        IsAdminVisible = true;
    }

    [RelayCommand]
    private void CloseAdmin()
    {
        IsAdminVisible = false;
        IsAdminAuthenticated = false;
        AdminUserInput = string.Empty;
        AdminPasswordInput = string.Empty;
        AdminErrorMessage = string.Empty;
    }

    [RelayCommand]
    private void AuthenticateAdmin()
    {
        var isValid = string.Equals(AdminUserInput, AdminCredentials.Username, StringComparison.OrdinalIgnoreCase)
                      && AdminPasswordInput == AdminCredentials.Password;
        if (isValid)
        {
            IsAdminAuthenticated = true;
            AdminErrorMessage = string.Empty;
        }
        else
        {
            SystemSounds.Hand.Play();
            AdminErrorMessage = "Credentiale invalide.";
        }
    }

    [RelayCommand]
    private void BrowseExcel()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            CheckFileExists = true,
            Title = "Selecteaza fisierul Excel"
        };

        if (dialog.ShowDialog() == true)
        {
            ExcelPathInput = dialog.FileName;
        }
    }

    [RelayCommand]
    private void BrowseTemplate()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Word template (*.docx)|*.docx|All files (*.*)|*.*",
            CheckFileExists = true,
            Title = "Selecteaza sablonul DOCX"
        };

        if (dialog.ShowDialog() == true)
        {
            TemplatePathInput = dialog.FileName;
        }
    }

    [RelayCommand]
    private async Task SaveAdminSettingsAsync()
    {
        if (_settings is null)
        {
            _settings = new KioskSettings();
        }

        if (!string.IsNullOrWhiteSpace(ExcelPathInput) && !File.Exists(ExcelPathInput))
        {
            AdminErrorMessage = "Fisierul Excel nu exista.";
            SystemSounds.Hand.Play();
            return;
        }

        if (!string.IsNullOrWhiteSpace(TemplatePathInput) && !File.Exists(TemplatePathInput))
        {
            AdminErrorMessage = "Sablonul DOCX nu exista.";
            SystemSounds.Hand.Play();
            return;
        }

        _settings.ExcelPathOverride = NormalizePathForStorage(ExcelPathInput);
        _settings.TemplatePathOverride = NormalizePathForStorage(TemplatePathInput);
        _settings.LockMode = IsLockMode;
        _settings.DefaultCameraMoniker = SelectedCamera?.MonikerString ?? string.Empty;
        _settings.DefaultPrinter = SelectedPrinter ?? string.Empty;

        await PersistSettingsAsync().ConfigureAwait(false);
        AdminErrorMessage = "Setari salvate.";
        ShowStatus("Setarile au fost salvate.", false, TimeSpan.FromSeconds(4));
    }

    [RelayCommand]
    private async Task ExitApplicationAsync()
    {
        try
        {
            await _qrScannerService.StopAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Nu am putut opri camera inainte de iesire.");
        }

        await _dispatcher.InvokeAsync(() =>
        {
            try
            {
                if (Application.Current?.MainWindow is Window window)
                {
                    window.Close();
                }
                else
                {
                    Application.Current?.Shutdown();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Nu am putut inchide aplicatia prin scurtatura.");
                Application.Current?.Shutdown();
            }
        }, DispatcherPriority.Normal);
    }
    [RelayCommand]
    private async Task ReloadExcelAsync()
    {
        await LoadAttendeesAsync().ConfigureAwait(false);
        ShowStatus("Lista Excel a fost reincarcata.", false, TimeSpan.FromSeconds(4));
    }

    private async Task LoadSettingsAsync(CancellationToken cancellationToken)
    {
        try
        {
            _settings = await _settingsService.GetAsync(cancellationToken).ConfigureAwait(false);
            await _dispatcher.InvokeAsync(() =>
            {
                IsLockMode = _settings!.LockMode;

                Cameras.Clear();
                foreach (var camera in _qrScannerService.GetAvailableCameras())
                {
                    Cameras.Add(camera);
                }

                var targetCamera = Cameras.FirstOrDefault(c => c.MonikerString == _settings.DefaultCameraMoniker)
                                   ?? Cameras.FirstOrDefault();
                SelectedCamera = targetCamera;
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Eroare la incarcare setari.");
            ShowStatus("Eroare la incarcare setari", true, TimeSpan.FromSeconds(10));
        }
    }

    private Task LoadPrintersAsync()
    {
        return _dispatcher.InvokeAsync(() =>
        {
            Printers.Clear();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                Printers.Add(printer);
            }

            if (_settings is not null && !string.IsNullOrWhiteSpace(_settings.DefaultPrinter))
            {
                SelectedPrinter = Printers.FirstOrDefault(p => string.Equals(p, _settings.DefaultPrinter, StringComparison.OrdinalIgnoreCase));
            }

            SelectedPrinter ??= Printers.FirstOrDefault();
        }).Task;
    }

    private async Task LoadAttendeesAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            await _dispatcher.InvokeAsync(() => IsBusy = true);
            await _attendeeRepository.LoadAsync(cancellationToken).ConfigureAwait(false);
            _logger.LogInformation("{Count} participanti incarcati din Excel.", _attendeeRepository.All.Count);
        }
        catch (FileNotFoundException fnf)
        {
            _logger.LogError(fnf, "Fisierul Excel lipseste.");
            ShowStatus("Fisierul Excel nu a fost gasit.", true, TimeSpan.FromSeconds(10));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Eroare la incarcarea participantilor.");
            ShowStatus("Eroare la incarcarea participantilor.", true, TimeSpan.FromSeconds(10));
        }
        finally
        {
            await _dispatcher.InvokeAsync(() => IsBusy = false);
        }
    }

    private async Task StartScannerAsync(CancellationToken cancellationToken)
    {
        try
        {
            await _qrScannerService.StartAsync(SelectedCamera?.MonikerString, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Nu am putut porni camera QR.");
            ShowStatus("Nu am putut porni camera.", true, TimeSpan.FromSeconds(10));
        }
    }

    private void HandleFrameReady(object? sender, BitmapSource frame)
    {
        if (_disposed)
        {
            return;
        }

        _dispatcher.Invoke(() => CameraFrame = frame);
    }

    private void HandleCodeDetected(object? sender, QrCodeDetection detection)
    {
        if (_disposed)
        {
            return;
        }

        _dispatcher.Invoke(() => ProcessDetection(detection));
    }

    private Attendee? FindAttendeeFromCode(string qrValue)
    {
        if (string.IsNullOrWhiteSpace(qrValue))
        {
            return null;
        }

        var direct = _attendeeRepository.FindById(qrValue);
        if (direct is not null)
        {
            return direct;
        }

        return _attendeeRepository.All.FirstOrDefault(attendee =>
            !string.IsNullOrWhiteSpace(attendee.Id) &&
            qrValue.Contains(attendee.Id, StringComparison.OrdinalIgnoreCase));
    }
    private void ProcessDetection(QrCodeDetection detection)
    {
        var cooldown = TimeSpan.FromSeconds(_optionsMonitor.CurrentValue.CooldownSeconds);

        if (IsConfirmationVisible || IsPrintingVisible || IsAdminVisible)
        {
            _qrScannerService.RequestPause(cooldown);
            return;
        }

        var code = detection.Value.Trim();
        if (string.IsNullOrWhiteSpace(code))
        {
            return;
        }

        var attendee = FindAttendeeFromCode(code);
        if (attendee is null)
        {
            SystemSounds.Hand.Play();
            ShowStatus($"Cod necunoscut: {code}", true, TimeSpan.FromSeconds(5));
            _qrScannerService.RequestPause(cooldown);
            return;
        }

        try
        {
            _attendeeRepository.MarkPresentAsync(attendee.Id).GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Nu am putut marca participantul {AttendeeId} ca prezent.", attendee.Id);
            SystemSounds.Hand.Play();
            ShowStatus("Nu am putut actualiza statusul din fisierul Excel.", true, TimeSpan.FromSeconds(8));
            _qrScannerService.RequestPause(cooldown);
            return;
        }

        SystemSounds.Asterisk.Play();
        CurrentAttendee = attendee;
        IsConfirmationVisible = true;
        _qrScannerService.RequestPause(cooldown);
    }

    private void ShowStatus(string message, bool isError, TimeSpan? hideAfter = null)
    {
        void ApplyStatus()
        {
            _statusTimer.Stop();
            StatusMessage = message;
            IsStatusError = isError;
            IsStatusVisible = true;

            if (hideAfter is { } duration && duration > TimeSpan.Zero)
            {
                _statusTimer.Interval = duration;
                _statusTimer.Start();
            }
        }

        if (_dispatcher.CheckAccess())
        {
            ApplyStatus();
        }
        else
        {
            _dispatcher.Invoke(ApplyStatus);
        }
    }

    private void HideStatus()
    {
        void ApplyHide()
        {
            _statusTimer.Stop();
            IsStatusVisible = false;
            StatusMessage = string.Empty;
            IsStatusError = false;
        }

        if (_dispatcher.CheckAccess())
        {
            ApplyHide();
        }
        else
        {
            _dispatcher.Invoke(ApplyHide);
        }
    }

    private void OnStatusTimerTick(object? sender, EventArgs e) => HideStatus();

    private async Task ApplyCameraSelectionAsync(CameraDevice camera)
    {
        try
        {
            await _dispatcher.InvokeAsync(() => IsBusy = true);
            await _qrScannerService.StartAsync(camera.MonikerString).ConfigureAwait(false);
            if (_settings is not null)
            {
                _settings.DefaultCameraMoniker = camera.MonikerString;
                await PersistSettingsAsync().ConfigureAwait(false);
            }
            ShowStatus($"Camera schimbata: {camera.Name}", false, TimeSpan.FromSeconds(3));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Nu am putut schimba camera.");
            ShowStatus("Eroare la schimbarea camerei.", true, TimeSpan.FromSeconds(6));
        }
        finally
        {
            await _dispatcher.InvokeAsync(() => IsBusy = false);
        }
    }

    private async Task PersistSettingsAsync()
    {
        if (_settings is null)
        {
            return;
        }

        try
        {
            await _settingsService.SaveAsync(_settings).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Nu am putut salva setarile.");
            ShowStatus("Eroare la salvarea setarilor.", true, TimeSpan.FromSeconds(6));
        }
    }

    private void InitializeAdminFields()
    {
        AdminUserInput = string.Empty;
        AdminPasswordInput = string.Empty;
        AdminErrorMessage = string.Empty;
        ExcelPathInput = GetDisplayPath(_settings?.ExcelPathOverride, _optionsMonitor.CurrentValue.ExcelPath);
        TemplatePathInput = GetDisplayPath(_settings?.TemplatePathOverride, _optionsMonitor.CurrentValue.TemplatePath);
    }

    private static string GetDisplayPath(string? overridePath, string fallback)
    {
        var candidate = string.IsNullOrWhiteSpace(overridePath) ? fallback : overridePath;
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return string.Empty;
        }

        if (!Path.IsPathRooted(candidate))
        {
            candidate = Path.Combine(AppContext.BaseDirectory, candidate);
        }

        return Path.GetFullPath(candidate);
    }

    private static string NormalizePathForStorage(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        var trimmed = path.Trim();
        var full = Path.GetFullPath(trimmed);
        var baseDir = Path.GetFullPath(AppContext.BaseDirectory);
        if (full.StartsWith(baseDir, StringComparison.OrdinalIgnoreCase))
        {
            var relative = Path.GetRelativePath(baseDir, full);
            return relative;
        }

        return full;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _statusTimer.Stop();
        _statusTimer.Tick -= OnStatusTimerTick;
        _qrScannerService.FrameReady -= HandleFrameReady;
        _qrScannerService.CodeDetected -= HandleCodeDetected;
        _qrScannerService.StopAsync().GetAwaiter().GetResult();
    }
}

