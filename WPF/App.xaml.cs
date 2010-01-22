namespace WPF
{
    using System;

    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : WpfApplication
    {
        #region Constructors

        /// <summary>
        /// Loads default messages from application resources
        /// </summary>
        public App()
        {
            UnhandledExceptionMessage = WPF.Properties.Resources.UnhandledExceptionMessage;
            HandleMissingAssemblyMessage = WPF.Properties.Resources.HandleMissingAssemblyMessage;
        }

        #endregion Constructors

        #region Methods

        #region Protected Methods

        protected override void OnStartup(System.Windows.StartupEventArgs e)
        {
            base.OnStartup(e);

            string[] args = e.Args;
        }

        /// <summary>
        /// Shows an example dialog box when a new instance is started
        /// </summary>
        /// <param name="e">The arguments that describe this event</param>
        protected override void OnStartupNextInstance(StartupNextInstanceEventArgs e)
        {
            base.OnStartupNextInstance(e);

            string[] args = e.Args;

            MainWindow.Activate();
            System.Windows.MessageBox.Show(MainWindow, "Only one at a time please");
        }

        #endregion Protected Methods

        #endregion Methods
    }
}