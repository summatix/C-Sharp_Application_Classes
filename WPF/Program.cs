namespace Wpf
{
    using System;
    using System.Windows;

    using Properties;

    /// <summary>
    /// Controls the application execution
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// The entry point to the application
        /// </summary>
        /// <returns>The application return code</returns>
        [STAThread]
        public static int Main()
        {
            Initialize();

            var app = new App();
            app.InitializeComponent();

            var wpfApp = new WpfApplication(app);
            SetupApplicationSettings();
            InitializeApplicationEvents();

            return wpfApp.Run();
        }

        /// <summary>
        /// The entry point to the application when the single instance application has already been started
        /// </summary>
        private static void NextMain()
        {
            WpfApplication.Current.MainWindow.Activate();
            MessageBox.Show(WpfApplication.Current.MainWindow, "Only one at a time please");
        }

        /// <summary>
        /// Initializes the application before the main code is executed
        /// </summary>
        private static void Initialize()
        {
            // Add initialization logic here
        }

        /// <summary>
        /// Loads the settings from the App.xaml file to initialize the WpfApplication instance
        /// </summary>
        private static void SetupApplicationSettings()
        {
            var wpfApp = WpfApplication.Current;
            var isSingleInstance = wpfApp.TryFindResource("IsSingleInstance");
            if (isSingleInstance is bool)
            {
                wpfApp.IsSingleInstance = (bool)isSingleInstance;
            }

            var handleUnresolvedAssembly = wpfApp.TryFindResource("HandleUnresolvedAssembly");
            if (handleUnresolvedAssembly is bool)
            {
                wpfApp.HandleUnresolvedAssembly = (bool)handleUnresolvedAssembly;
            }

            var handleUnhandledException = wpfApp.TryFindResource("HandleUnhandledException");
            if (handleUnhandledException is bool)
            {
                wpfApp.HandleUnhandledException = (bool)handleUnhandledException;
            }

            WpfApplication.UnhandledExceptionMessage = Resources.UnhandledExceptionMessage;
            WpfApplication.HandleMissingAssemblyMessage = Resources.HandleMissingAssemblyMessage;
        }

        /// <summary>
        /// Initializes the events of the WpfApplication instance
        /// </summary>
        private static void InitializeApplicationEvents()
        {
            WpfApplication.Current.StartupNextInstance += (s, e) => NextMain();
            WpfApplication.Current.Exit += (s, e) => CloseApplication();
        }

        /// <summary>
        /// Called when the application is closing. Does any final processing and saves the settings
        /// </summary>
        private static void CloseApplication()
        {
            // Add closing down logic here
        }
    }
}