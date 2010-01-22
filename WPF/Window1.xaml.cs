namespace WPF
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Threading;
    using SuperLibrary;

    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public Window1()
        {
            InitializeComponent();

            WpfApplication.Current.DispatcherUnhandledException += Application_DispatcherUnhandledException;
        }

        #endregion Constructors

        #region Methods

        #region Private Methods

        private void Application_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            if (ignore.IsChecked.Value)
            {
                e.Handled = true;
            }
        }

        private void ThrowException_Click(object sender, RoutedEventArgs e)
        {
            throw new InvalidOperationException
                ("This is an example of an exception thrown which is not caught within a try-catch block");
        }

        private void UseMissingDLL_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(System.IO.Path.Combine(Directory.GetCurrentDirectory(), "SuperLibrary.dll")))
            {
                MessageBox.Show("You need to delete SuperLibrary.dll first");
            }
            else
            {
                Functions.SayHello(); // This should cause an error since the SuperLibrary.dll does not exist
            }
        }

        #endregion Private Methods

        #endregion Methods
    }
}