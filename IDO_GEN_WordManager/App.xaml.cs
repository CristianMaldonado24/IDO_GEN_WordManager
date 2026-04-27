using System;
using System.Windows;
using System.Windows.Threading;

namespace IDO_GEN_WordManager;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Captura excepciones no controladas en el hilo UI
        DispatcherUnhandledException += (s, ex) =>
        {
            MessageBox.Show($"Error inesperado:\n\n{ex.Exception.Message}\n\n{ex.Exception.StackTrace}",
                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            ex.Handled = true;
        };

        // Captura excepciones en otros hilos
        AppDomain.CurrentDomain.UnhandledException += (s, ex) =>
        {
            if (ex.ExceptionObject is Exception exc)
                MessageBox.Show($"Error crítico:\n\n{exc.Message}",
                    "Error crítico", MessageBoxButton.OK, MessageBoxImage.Error);
        };
    }
}

