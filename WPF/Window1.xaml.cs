namespace Wpf
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Threading;

    using SuperLibrary;

    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1
    {
        /// <summary>
        /// Initializes a new instance of the Window1 class
        /// </summary>
        public Window1()
        {
            InitializeComponent();

            WpfApplication.Current.DispatcherUnhandledException += Application_DispatcherUnhandledException;
        }

        private void Application_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            if (ignore.IsChecked != null && ignore.IsChecked.Value)
            {
                e.Handled = true;
            }
        }

        private void ThrowException_Click(object sender, RoutedEventArgs e)
        {
            throw new InvalidOperationException("This is an example of an exception thrown which is not caught within a try-catch block");
        }

        private void UseMissingDLL_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(Path.Combine(Directory.GetCurrentDirectory(), "SuperLibrary.dll")))
            {
                MessageBox.Show("You need to delete SuperLibrary.dll first");
            }
            else
            {
                Functions.SayHello(); // This should cause an error since the SuperLibrary.dll does not exist
            }
        }
    }
}