#region Header

/*
 * Copyright (c) Robert Roose 2009
 *
 *  All rights reserved.
 *
 *  Redistribution and use in source and binary forms, with or without
 *  modification, are permitted provided that the following conditions are
 *  met:
 *
 *      * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *      * Neither the name of the copyright holder nor the names of its
 *        contributors may be used to endorse or promote products derived from
 *        this software without specific prior written permission.
 *
 *  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 *  AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 *  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 *  ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
 *  LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 *  CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 *  SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 *  INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 *  CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 *  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
 *  THE POSSIBILITY OF SUCH DAMAGE.
 */

#endregion Header

namespace WPF
{
    using Microsoft.VisualBasic.ApplicationServices;
    using System;
    using System.Collections;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Reflection;
    using System.Security;
    using System.Windows;
    using System.Windows.Navigation;
    using System.Windows.Resources;
    using System.Windows.Threading;

    using ShutdownMode = System.Windows.ShutdownMode;
    using StartupEventArgs = System.Windows.StartupEventArgs;
    using StartupEventHandler = System.Windows.StartupEventHandler;

    #region Delegates

    /// <summary>
    /// Represents the method that handles the WpfApplication.StartupNextInstance event.
    /// </summary>
    /// <param name="sender">The object that raised the event.</param>
    /// <param name="e">The event data.</param>
    public delegate void StartupNextInstanceEventHandler(object sender, StartupNextInstanceEventArgs e);

    #endregion Delegates

    /// <summary>
    /// Contains the arguments for the WpfApplication.StartupNextInstance event
    /// </summary>
    public class StartupNextInstanceEventArgs : EventArgs
    {
        #region Constructors

        /// <summary>
        /// Initializes the instance.
        /// </summary>
        /// <param name="args">The command line arguments that were passed to the application.</param>
        public StartupNextInstanceEventArgs(string[] args)
        {
            Args = args;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets command line arguments that were passed to the application from either the command prompt or the
        /// desktop.
        /// </summary>
        /// <returns>A string array that contains the command line arguments that were passed to the application from
        /// either the command prompt or the desktop. If no command line arguments were passed, the string array as
        /// zero items.</returns>
        public string[] Args
        {
            get;
            private set;
        }

        #endregion Properties
    }

    /// <summary>
    /// Encapsulates a Windows Presentation Foundation (WPF) application and provides single instance support
    /// </summary>
    public class WpfApplication : DependencyObject
    {
        #region Fields

        private readonly Application _app = new Application();

        private static string _handleMissingAssemblyMessage = "{0} is needed to enable this application to run.";
        private static string _unhandledExceptionMessage =
            "A fatal error has occurred and the program has been unable to recover.";

        private bool _handleUnresolvedAssembly = true;
        private bool _isSingleInstance = true;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the WpfApplication class.
        /// </summary>
        /// <exception cref="InvalidOperationException">More than one instance of the System.Windows.Application class
        /// is created per System.AppDomain.</exception>
        [SecurityCritical]
        public WpfApplication()
        {
            Current = this;
            InitializeEvents();

            // TODO: Solve why this randomly fails
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        #endregion Constructors

        #region Events

        /// <summary>
        /// Occurs when an application becomes the foreground application.
        /// </summary>
        public event EventHandler Activated;

        /// <summary>
        /// Occurs when an application stops being the foreground application.
        /// </summary>
        public event EventHandler Deactivated;

        /// <summary>
        /// Occurs when an exception is thrown by an application but not handled.
        /// </summary>
        public event DispatcherUnhandledExceptionEventHandler DispatcherUnhandledException;

        /// <summary>
        /// Occurs just before an application shuts down, and cannot be canceled.
        /// </summary>
        public event ExitEventHandler Exit;

        /// <summary>
        /// Occurs when a navigator in the application begins navigation to a content fragment, Navigation occurs
        /// immediately if the desired fragment is in the current content, or after the source XAML content has been
        /// loaded if the desired fragment is in different content.
        /// </summary>
        public event FragmentNavigationEventHandler FragmentNavigation;

        /// <summary>
        /// Occurs when content that was navigated to by a navigator in the application has been loaded, parsed, and
        /// has begun rendering.
        /// </summary>
        public event LoadCompletedEventHandler LoadCompleted;

        /// <summary>
        /// Occurs when the content that is being navigated to by a navigator in the application has been found,
        /// although it may not have completed loading.
        /// </summary>
        public event NavigatedEventHandler Navigated;

        /// <summary>
        /// Occurs when a new navigation is requested by a navigator in the application.
        /// </summary>
        public event NavigatingCancelEventHandler Navigating;

        /// <summary>
        /// Occurs when an error occurs while a navigator in the application is navigating to the requested content.
        /// </summary>
        public event NavigationFailedEventHandler NavigationFailed;

        /// <summary>
        /// Occurs periodically during a download that is being managed by a navigator in the application to provide
        /// navigation progress information.
        /// </summary>
        public event NavigationProgressEventHandler NavigationProgress;

        /// <summary>
        /// Occurs when the StopLoading method of a navigator in the application is called, or when a new navigation is
        /// requested by a navigator while a current navigation is in progress.
        /// </summary>
        public event NavigationStoppedEventHandler NavigationStopped;

        /// <summary>
        /// Occurs when the user ends the Windows session by logging off or shutting down the operating system.
        /// </summary>
        public event SessionEndingCancelEventHandler SessionEnding;

        /// <summary>
        /// Occurs when the System.Windows.Application.Run() method of the System.Windows.Application object is called.
        /// </summary>
        public event StartupEventHandler Startup;

        /// <summary>
        /// Occurs when a new instance of the application is run and the application has been set to single instance
        /// </summary>
        public event StartupNextInstanceEventHandler StartupNextInstance;

        #endregion Events

        #region Properties

        /// <summary>
        /// Gets the WpfApplication object for the current System.AppDomain.
        /// </summary>
        /// <returns>The WpfApplication object for the current System.AppDomain.</returns>
        public static WpfApplication Current
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets the default message to display when the exception of a missing assembly reference occurs.
        /// </summary>
        public static string HandleMissingAssemblyMessage
        {
            get { return _handleMissingAssemblyMessage; }
            set { _handleMissingAssemblyMessage = value; }
        }

        /// <summary>
        /// Gets or sets the System.Reflection.Assembly that provides the pack uniform resource identifiers (URIs) for
        /// resources in a WPF application.
        /// </summary>
        /// <exception cref="InvalidOperationException">A WPF application has an entry assembly, or
        /// System.Windows.Application.ResourceAssembly has already been set.</exception>
        /// <returns>A reference to the System.Reflection.Assembly that provides the pack uniform resource identifiers
        /// (URIs) for resources in a WPF application.</returns>
        public static Assembly ResourceAssembly
        {
            get { return Application.ResourceAssembly; }
            set { Application.ResourceAssembly = value; }
        }

        /// <summary>
        /// Gets or sets the default message to display when an unhandled exception occurs.
        /// </summary>
        public static string UnhandledExceptionMessage
        {
            get { return _unhandledExceptionMessage; }
            set { _unhandledExceptionMessage = value; }
        }

        /// <summary>
        /// Gets or sets whether to handle the event of an unhandled exception.
        /// </summary>
        [DefaultValue(false)]
        public bool HandleUnhandledException
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets whether to handle the event of a missing library reference (a missing DLL).
        /// </summary>
        [DefaultValue(true)]
        public bool HandleUnresolvedAssembly
        {
            get { return _handleUnresolvedAssembly; }
            set { _handleUnresolvedAssembly = value; }
        }

        /// <summary>
        /// Gets or sets whether this application is a single instance application
        /// </summary>
        [DefaultValue(true)]
        public bool IsSingleInstance
        {
            get { return _isSingleInstance; }
            set { _isSingleInstance = value; }
        }

        /// <summary>
        /// Gets or sets the main window of the application.
        /// </summary>
        /// <exception cref="InvalidOperationException">System.Windows.Application.MainWindow is set from an
        /// application that's hosted n a browser, such as an XAML browser applications (XBAPs)</exception>
        /// <returns>A System.Windows.Window that is designated as the main application window.</returns>
        public Window MainWindow
        {
            get { return _app.MainWindow; }
            set { _app.MainWindow = value; }
        }

        /// <summary>
        /// Gets a collection of application-scope properties.
        /// </summary>
        /// <returns>An System.Collections.IDictionary that contains the application-scope properties.</returns>
        public IDictionary Properties
        {
            get { return _app.Properties; }
        }

        /// <summary>
        /// Gets or sets a collection of application-scope resources, such as styles and brushes.
        /// </summary>
        /// <returns>A System.Windows.ResourceDictionary object that contains zero or more application-scope
        /// resources.</returns>
        public ResourceDictionary Resources
        {
            get { return _app.Resources; }
            set { _app.Resources = value; }
        }

        /// <summary>
        /// Gets or sets the condition that causes the System.Windows.Application.Shutdown() method to be called.
        /// </summary>
        /// <returns>A System.Windows.ShutdownMode enumeration value. The default value is
        /// System.Windows.ShutdownMode.OnLastWindowClose.</returns>
        public ShutdownMode ShutdownMode
        {
            get { return _app.ShutdownMode; }
            set { _app.ShutdownMode = value; }
        }

        /// <summary>
        /// Gets or sets a UI that is automatically shown when an application starts.
        /// </summary>
        /// <exception cref="ArgumentNullException">System.Windows.Application.StartupUri is set with a value of
        /// null.</exception>
        /// <returns>A System.Uri that refers to the UI that automatically opens when an application starts.</returns>
        public Uri StartupUri
        {
            get { return _app.StartupUri; }
            set { _app.StartupUri = value; }
        }

        /// <summary>
        /// Gets the instantiated windows in an application.
        /// </summary>
        /// <returns>A System.Windows.WindowCollection that contains references to all window objects in the current
        /// System.AppDomain.</returns>
        public WindowCollection Windows
        {
            get { return _app.Windows; }
        }

        #endregion Properties

        #region Methods

        #region Public Static Methods

        /// <summary>
        /// Returns a resource stream for a content data file that is located at the specified System.Uri (see Windows
        /// Presentation Foundation Application Resource, Content, and Data Files).
        /// </summary>
        /// <exception cref="ArgumentNullException">The System.Uri that is passed to
        /// System.Windows.Application.GetContentStream(System.Uri) is null.</exception>
        /// <exception cref="ArgumentException">The System.Uri.OriginalString property of the System.Uri that is passed
        /// to System.Windows.Application.GetContentStream(System.Uri) is null.</exception>
        /// <exception cref="ArgumentException">The System.Uri that is passed to
        /// System.Windows.Application.GetContentStream(System.Uri) is an absolute System.Uri.</exception>
        /// <param name="uriContent">The relative System.Uri that maps to a loose resource.</param>
        /// <returns>A System.Windows.Resources.StreamResourceInfo that contains a content data file that is located at
        /// the specified System.Uri. If a loose resource is not found, null is returned.</returns>
        [SecurityCritical]
        public static StreamResourceInfo GetContentStream(Uri uriContent)
        {
            return Application.GetContentStream(uriContent);
        }

        /// <summary>
        /// Retrieves a cookie for the location specified by a System.Uri.
        /// </summary>
        /// <exception cref="Win32Exception">A Win32 error is raised by the InternetGetCookie function (called by
        /// System.Windows.Application.GetCookie(System.Uri)) if a problem occurs when attempting to retrieve the
        /// specified cookie.</exception>
        /// <param name="uri">The System.Uri that specifies the location for which a cookie was created.</param>
        /// <returns>A System.String value, if the cookie exists; otherwise, a System.ComponentModel.Win32Exception is
        /// thrown.</returns>
        public static string GetCookie(Uri uri)
        {
            return Application.GetCookie(uri);
        }

        /// <summary>
        /// Returns a resource stream for a site-of-origin data file that is located at the specified System.Uri (see
        /// Windows Presentation Foundation Application Resource, Content, and Data Files).
        /// </summary>
        /// <exception cref="ArgumentNullException">The System.Uri that is passed to
        /// System.Windows.Application.GetRemoteStream(System.Uri) is null.</exception>
        /// <exception cref="ArgumentException">The System.Uri.OriginalString property of the System.Uri that is passed
        /// to System.Windows.Application.GetRemoteStream(System.Uri) is null.</exception>
        /// <exception cref="ArgumentException">The System.Uri that is passed to
        /// System.Windows.Application.GetRemoteStream(System.Uri) is either not relative, or is absolute but not in
        /// the pack://siteoforigin:,,,/ form.</exception>
        /// <param name="uriRemote">The System.Uri that maps to a loose resource at the site of origin.</param>
        /// <returns>A System.Windows.Resources.StreamResourceInfo that contains a resource stream for a site-of-origin
        /// data file that is located at the specified System.Uri. If the loose resource is not found, null is
        /// returned.</returns>
        [SecurityCritical]
        public static StreamResourceInfo GetRemoteStream(Uri uriRemote)
        {
            return Application.GetRemoteStream(uriRemote);
        }

        /// <summary>
        /// Returns a resource stream for a resource data file that is located at the specified System.Uri (see Windows
        /// Presentation Foundation Application Resource, Content, and Data Files).
        /// </summary>
        /// <exception cref="ArgumentNullException">The System.Uri that is passed to
        /// System.Windows.Application.GetResourceStream(System.Uri) is null.</exception>
        /// <exception cref="ArgumentException">The System.Uri.OriginalString property of the System.Uri that is passed
        /// to System.Windows.Application.GetResourceStream(System.Uri) is null.</exception>
        /// <exception cref="ArgumentException">The System.Uri that is passed to
        /// System.Windows.Application.GetResourceStream(System.Uri) is either not relative, or is absolute but not in
        /// the pack://application:,,,/ form.</exception>
        /// <param name="uriResource">The System.Uri that maps to an embedded resource.</param>
        /// <returns>A System.Windows.Resources.StreamResourceInfo that contains a resource stream for resource data
        /// file that is located at the specified System.Uri. If the resource located at the specified System.Uri is
        /// not found, null is returned.</returns>
        [SecurityCritical]
        public static StreamResourceInfo GetResourceStream(Uri uriResource)
        {
            return Application.GetResourceStream(uriResource);
        }

        /// <summary>
        /// Loads a XAML file that is located at the specified uniform resource identifier (URI), and converts it to an
        /// instance of the object that is specified by the root element of the XAML file.
        /// </summary>
        /// <exception cref="ArgumentNullException">resourceLocator is null.</exception>
        /// <exception cref="ArgumentException">The System.Uri.OriginalString property of the resourceLocatorSystem.Uri
        /// parameter is null.</exception>
        /// <exception cref="ArgumentException">The resourceLocator is an absolute URI.</exception>
        /// <exception cref="Exception">The file is not a XAML file.</exception>
        /// <param name="resourceLocator">A System.Uri that maps to a relative XAML file.</param>
        /// <returns></returns>
        public static object LoadComponent(Uri resourceLocator)
        {
            return Application.LoadComponent(resourceLocator);
        }

        /// <summary>
        /// Loads a XAML file that is located at the specified uniform resource identifier (URI) and converts it to an
        /// instance of the object that is specified by the root element of the XAML file.
        /// </summary>
        /// <exception cref="ArgumentNullException">component is null.</exception>
        /// <exception cref="ArgumentNullException">resourceLocator is null.</exception>
        /// <exception cref="ArgumentException">The System.Uri.OriginalString property of the resourceLocatorSystem.Uri
        /// parameter is null.</exception>
        /// <exception cref="ArgumentException">The resourceLocator is an absolute URI.</exception>
        /// <exception cref="Exception">component is of a type that does not match the root element of the XAML
        /// file.</exception>
        /// <param name="component">An object of the same type as the root element of the XAML file.</param>
        /// <param name="resourceLocator">A System.Uri that maps to a relative XAML file.</param>
        [SecurityCritical]
        public static void LoadComponent(object component, Uri resourceLocator)
        {
            Application.LoadComponent(component, resourceLocator);
        }

        /// <summary>
        /// Creates a cookie for the location specified by a System.Uri.
        /// </summary>
        /// <exception cref="Win32Exception">A Win32 error is raised by the InternetSetCookie function (called by
        /// System.Windows.Application.SetCookie(System.Uri,System.String)) if a problem occurs when attempting to
        /// create the specified cookie.</exception>
        /// <param name="uri">The System.Uri that specifies the location for which the cookie should be
        /// created.</param>
        /// <param name="value">The System.String that contains the cookie data.</param>
        public static void SetCookie(Uri uri, string value)
        {
            Application.SetCookie(uri, value);
        }

        #endregion Public Static Methods

        #region Public Methods

        /// <summary>
        /// Searches for a user interface (UI) resource, such as a System.Windows.Style or System.Windows.Media.Brush,
        /// with the specified key, and throws an exception if the requested resource is not found (see Resources
        /// Overview).
        /// </summary>
        /// <exception cref="ResourceReferenceKeyNotFoundException">The resource cannot be found.</exception>
        /// <param name="resourceKey">The name of the resource to find.</param>
        /// <returns>The requested resource object. If the requested resource is not found, a
        /// System.Windows.ResourceReferenceKeyNotFoundException is thrown.</returns>
        public object FindResource(object resourceKey)
        {
            return _app.FindResource(resourceKey);
        }

        /// <summary>
        /// Starts a Windows Presentation Foundation (WPF) application.
        /// </summary>
        /// <exception cref="InvalidOperationException">System.Windows.Application.Run() is called from a
        /// browser-hosted application (for example, an XAML browser application (XBAP)).</exception>
        /// <returns>The System.Int32 application exit code that is returned to the operating system when the
        /// application shuts down. By default, the exit code value is 0.</returns>
        public int Run()
        {
            return Start(null);
        }

        /// <summary>
        /// Starts a Windows Presentation Foundation (WPF) application and opens the specified window.
        /// </summary>
        /// <exception cref="InvalidOperationException">System.Windows.Application.Run() is called from a
        /// browser-hosted application (for example, an XAML browser application (XBAP)).</exception>
        /// <param name="window">A System.Windows.Window that opens automatically when an application starts.</param>
        /// <returns>The System.Int32 application exit code that is returned to the operating system when the
        /// application shuts down. By default, the exit code value is 0.</returns>
        public int Run(Window window)
        {
            return Start(window);
        }

        /// <summary>
        /// Shuts down an application.
        /// </summary>
        public void Shutdown()
        {
            _app.Shutdown();
        }

        /// <summary>
        /// Shuts down an application that returns the specified exit code to the operating system.
        /// </summary>
        /// <param name="exitCode">An integer exit code for an application. The default exit code is 0.</param>
        [SecurityCritical]
        public void Shutdown(int exitCode)
        {
            _app.Shutdown(exitCode);
        }

        /// <summary>
        /// Searches for the specified resource.
        /// </summary>
        /// <param name="resourceKey">The name of the resource to find.</param>
        /// <returns>The requested resource object. If the requested resource is not found, a null reference is
        /// returned.</returns>
        public object TryFindResource(object resourceKey)
        {
            return _app.TryFindResource(resourceKey);
        }

        #endregion Public Methods

        #region Protected Methods

        /// <summary>
        /// Called when there is an unhandled exception. Shows a message to the user and exits the application
        /// </summary>
        /// <param name="e">The arguments that describe this event</param>
        protected virtual void HandleUnhandledExceptionHandler(DispatcherUnhandledExceptionEventArgs e)
        {
            if (_app.MainWindow != null)
            {
                MessageBox.Show(_app.MainWindow, UnhandledExceptionMessage);
            }
            else
            {
                MessageBox.Show(UnhandledExceptionMessage);
            }

            Environment.Exit(1);
        }

        /// <summary>
        /// Raises the Activated event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnActivated(EventArgs e)
        {
            if (Activated != null)
            {
                Activated(this, e);
            }
        }

        /// <summary>
        /// Raises the Deactivated event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnDeactivated(EventArgs e)
        {
            if (Deactivated != null)
            {
                Deactivated(this, e);
            }
        }

        /// <summary>
        /// Raises the DispatcherUnhandledException event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnDispatcherUnhandledException(DispatcherUnhandledExceptionEventArgs e)
        {
            if (DispatcherUnhandledException != null)
            {
                DispatcherUnhandledException(this, e);
            }

            if (!e.Handled && HandleUnhandledException)
            {
                HandleUnhandledExceptionHandler(e);
            }
        }

        /// <summary>
        /// Raises the Exit event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnExit(ExitEventArgs e)
        {
            if (Exit != null)
            {
                Exit(this, e);
            }
        }

        /// <summary>
        /// Raises the FragmentNavigation event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnFragmentNavigation(FragmentNavigationEventArgs e)
        {
            if (FragmentNavigation != null)
            {
                FragmentNavigation(this, e);
            }
        }

        /// <summary>
        /// Raises the LoadCompleted event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnLoadCompleted(NavigationEventArgs e)
        {
            if (LoadCompleted != null)
            {
                LoadCompleted(this, e);
            }
        }

        /// <summary>
        /// Raises the Navigated event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnNavigated(NavigationEventArgs e)
        {
            if (Navigated != null)
            {
                Navigated(this, e);
            }
        }

        /// <summary>
        /// Raises the Navigating event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnNavigating(NavigatingCancelEventArgs e)
        {
            if (Navigating != null)
            {
                Navigating(this, e);
            }
        }

        /// <summary>
        /// Raises the NavigationFailed event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnNavigationFailed(NavigationFailedEventArgs e)
        {
            if (NavigationFailed != null)
            {
                NavigationFailed(this, e);
            }
        }

        /// <summary>
        /// Raises the NavigationProgress event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnNavigationProgress(NavigationProgressEventArgs e)
        {
            if (NavigationProgress != null)
            {
                NavigationProgress(this, e);
            }
        }

        /// <summary>
        /// Raises the NavigationStopped event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnNavigationStopped(NavigationEventArgs e)
        {
            if (NavigationStopped != null)
            {
                NavigationStopped(this, e);
            }
        }

        /// <summary>
        /// Called when an assembly is missing. Displays a message to the user and exits the application
        /// </summary>
        /// <param name="args">The missing assembly event arguments</param>
        /// <returns>null to indicate that the assembly could not be found</returns>
        protected virtual Assembly OnResolveEventHandler(ResolveEventArgs args)
        {
            if (HandleUnresolvedAssembly)
            {
                HandleMissingAssembly(GetFileName(args.Name));
            }

            return null;
        }

        /// <summary>
        /// Raises the SessionEnding event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnSessionEnding(SessionEndingCancelEventArgs e)
        {
            if (SessionEnding != null)
            {
                SessionEnding(this, e);
            }
        }

        /// <summary>
        /// Raises the Startup event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnStartup(StartupEventArgs e)
        {
            if (Startup != null)
            {
                Startup(this, e);
            }
        }

        /// <summary>
        /// Raises the StartupNextInstance event.
        /// </summary>
        /// <param name="e">The arguments that describe this event.</param>
        protected virtual void OnStartupNextInstance(StartupNextInstanceEventArgs e)
        {
            if (Startup != null)
            {
                StartupNextInstance(this, e);
            }
        }

        #endregion Protected Methods

        #region Private Static Methods

        /// <summary>
        /// Get the missing assembly filename from the given missing item name
        /// </summary>
        /// <param name="itemName">The item to retreive the name for</param>
        /// <returns>The missing assembly filename</returns>
        private static string GetFileName(string itemName)
        {
            return string.Format("{0}.dll", new AssemblyName(itemName).Name);
        }

        /// <summary>
        /// Returns the formatted string message, or an empty string if there is an error.
        /// </summary>
        /// <param name="format">A composite format string.</param>
        /// <param name="arg">The value to format.</param>
        /// <returns>The formatted string or string.Empty if there was an error.</returns>
        private static string GetFormattedMessage(string format, string arg)
        {
            if (string.IsNullOrEmpty(format))
            {
                return string.Empty;
            }

            try
            {
                return string.Format(format, arg);
            }
            catch (FormatException)
            {
                return string.Empty;
            }
        }

        #endregion Private Static Methods

        #region Private Methods

        /// <summary>
        /// Called when an assembly is missing
        /// </summary>
        /// <param name="sender">Not used</param>
        /// <param name="args">The event handler arguemnts</param>
        /// <returns>The found assembly or null if it cannot be resolved</returns>
        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return OnResolveEventHandler(args);
        }

        /// <summary>
        /// Displays an error messge to the user about the missing assembly file and then exits the system.
        /// </summary>
        /// <param name="filename">The filename of the missing assembly</param>
        private void HandleMissingAssembly(string filename)
        {
            if (_app.MainWindow != null)
            {
                MessageBox.Show(_app.MainWindow, GetFormattedMessage(HandleMissingAssemblyMessage, filename));
            }
            else
            {
                MessageBox.Show(GetFormattedMessage(HandleMissingAssemblyMessage, filename));
            }

            Environment.Exit(1);
        }

        /// <summary>
        /// Initializes the Application event wrappers
        /// </summary>
        private void InitializeEvents()
        {
            _app.Activated += ((sender, e) => OnActivated(e));
            _app.Deactivated += ((sender, e) => OnDeactivated(e));
            _app.DispatcherUnhandledException += ((sender, e) => OnDispatcherUnhandledException(e));
            _app.Exit += ((sender, e) => OnExit(e));
            _app.FragmentNavigation += ((sender, e) => OnFragmentNavigation(e));
            _app.LoadCompleted += ((sender, e) => OnLoadCompleted(e));
            _app.Navigated += ((sender, e) => OnNavigated(e));
            _app.Navigating += ((sender, e) => OnNavigating(e));
            _app.NavigationFailed += ((sender, e) => OnNavigationFailed(e));
            _app.NavigationProgress += ((sender, e) => OnNavigationProgress(e));
            _app.NavigationStopped += ((sender, e) => OnNavigationStopped(e));
            _app.SessionEnding += ((sender, e) => OnSessionEnding(e));
            _app.Startup += ((sender, e) => OnStartup(e));
        }

        /// <summary>
        /// Starts the single instance application wrapper
        /// </summary>
        /// <param name="window">The Window which is to be started. This can be null</param>
        /// <returns>The return code returned from the WPF Application instance</returns>
        private int Start(Window window)
        {
            var wrapper = new SingleInstanceApplicationWrapper(this, IsSingleInstance) { Window = window };

            // Although an empty array of string can be given to the application wrapper, if done so then subsequent
            // instances will not be provided command line arguments in the
            // Microsoft.VisualBasic.ApplicationServices.StartupEventArgs argument
            wrapper.Run(Environment.GetCommandLineArgs());
            return wrapper.ReturnCode;
        }

        #endregion Private Methods

        #endregion Methods

        #region Nested Types

        /// <summary>
        /// An application wrapper that makes the WPF application a single instance application
        /// </summary>
        private class SingleInstanceApplicationWrapper : WindowsFormsApplicationBase
        {
            #region Fields

            private readonly WpfApplication _app;

            #endregion Fields

            #region Constructors

            /// <summary>
            /// Initializes the instance
            /// </summary>
            /// <param name="app">The reference to the WPF application instance</param>
            /// <param name="isSingleInstance">Set to true to make the application a single instance
            /// application</param>
            public SingleInstanceApplicationWrapper(WpfApplication app, bool isSingleInstance)
            {
                _app = app;
                IsSingleInstance = isSingleInstance;
            }

            #endregion Constructors

            #region Properties

            /// <summary>
            /// Gets the return code returned from the WpfApplication
            /// </summary>
            public int ReturnCode
            {
                get;
                private set;
            }

            /// <summary>
            /// Gets or sets the Window instance to provide to the WpfApplication.Run(Window) method. If null,
            /// WpfApplication.Run() will be called instead
            /// </summary>
            public Window Window
            {
                private get;
                set;
            }

            #endregion Properties

            #region Methods

            #region Protected Methods

            /// <summary>
            /// Runs the WpfApplication instance inside this wrapper, returning only when this application wrapper has
            /// completed
            /// </summary>
            /// <param name="eventArgs">Not used</param>
            /// <returns>false to indicated that this instance should not proceed in starting</returns>
            protected override bool OnStartup(Microsoft.VisualBasic.ApplicationServices.StartupEventArgs eventArgs)
            {
                if (Window == null)
                {
                    ReturnCode = _app._app.Run();
                }
                else
                {
                    ReturnCode = _app._app.Run(Window);
                    Window = null;
                }

                return false;
            }

            /// <summary>
            /// Calls the appropriate method of the WPF application when the next instance of the application is
            /// started
            /// </summary>
            /// <param name="eventArgs">The arguments that describe this event</param>
            protected override void OnStartupNextInstance(
                Microsoft.VisualBasic.ApplicationServices.StartupNextInstanceEventArgs eventArgs)
            {
                // Convert the ReadOnlyCollection<string> arguments to a string[] collection of arguments and remove
                // the first argument which will always be the application path
                ReadOnlyCollection<string> argsCollection = eventArgs.CommandLine;
                int count = argsCollection.Count - 1;
                string[] args;

                if (count > 0)
                {
                    args = new string[eventArgs.CommandLine.Count - 1];
                    for (int i = 0; i < count; ++i)
                    {
                        args[i] = argsCollection[i + 1];
                    }
                }
                else
                {
                    args = new string[0];
                }

                _app.OnStartupNextInstance(new StartupNextInstanceEventArgs(args));
            }

            #endregion Protected Methods

            #endregion Methods
        }

        #endregion Nested Types
    }
}