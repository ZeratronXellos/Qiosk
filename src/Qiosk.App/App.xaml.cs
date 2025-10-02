using System;
using System.Windows;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using NLog.Extensions.Logging;
using Qiosk.App.Configuration;
using Qiosk.App.Services;
using Qiosk.App.Services.Contracts;
using Qiosk.App.ViewModels;
using Qiosk.App.Views;

namespace Qiosk.App;

public partial class App : Application
{
    private readonly IHost _host;

    public App()
    {
        _host = Host.CreateDefaultBuilder()
            .UseContentRoot(AppContext.BaseDirectory)
            .ConfigureAppConfiguration((_, config) =>
            {
                config.SetBasePath(AppContext.BaseDirectory);
                config.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
                config.AddJsonFile("kiosk-settings.json", optional: true, reloadOnChange: true);
            })
            .ConfigureServices((context, services) =>
            {
                services.Configure<KioskOptions>(context.Configuration.GetSection(KioskOptions.SectionName));

                services.AddSingleton<IKioskSettingsService, KioskSettingsService>();
                services.AddSingleton<IAttendeeRepository, ExcelAttendeeRepository>();
                services.AddSingleton<IQrScannerService, QrScannerService>();
                services.AddSingleton<IBadgePrinter, DocxBadgePrinter>();

                services.AddSingleton<ShellViewModel>();
                services.AddSingleton<ShellWindow>();
            })
            .ConfigureLogging(logging =>
            {
                logging.ClearProviders();
                logging.SetMinimumLevel(LogLevel.Information);
                logging.AddNLog();
            })
            .Build();
    }

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        await _host.StartAsync();

        var window = _host.Services.GetRequiredService<ShellWindow>();
        window.DataContext = _host.Services.GetRequiredService<ShellViewModel>();
        window.Show();
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        await _host.StopAsync();
        _host.Dispose();
        base.OnExit(e);
    }
}
