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

namespace Console
{
    using System;
    using System.ComponentModel;
    using System.Reflection;
    using System.Security;

    /// <summary>
    /// Encapsulates a Console application
    /// </summary>
    public class ConsoleApplication
    {
        #region Fields

        private static string _handleMissingAssemblyMessage = "{0} is needed to enable this application to run.";
        private static string _unhandledExceptionMessage = 
            "A fatal error has occurred and the program has been unable to recover.";

        private bool _handleUnhandledException = true;
        private bool _handleUnresolvedAssembly = true;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the ConsoleApplication class
        /// </summary>
        /// <exception cref="InvalidOperationException">More than one instance of the ConsoleApplication class is
        /// created per System.AppDomain.</exception>
        [SecurityCritical]
        public ConsoleApplication()
        {
            if (Current != null)
            {
                throw new InvalidOperationException("Only one ConsoleApplication can be created per System.AppDomain");
            }

            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            Current = this;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets the ConsoleApplication object for the current System.AppDomain
        /// </summary>
        public static ConsoleApplication Current
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
        /// Gets or sets the default message to display when an unhandled exception occurs.
        /// </summary>
        public static string UnhandledExceptionMessage
        {
            get { return _unhandledExceptionMessage; }
            set { _unhandledExceptionMessage = value; }
        }

        /// <summary>
        /// Gets the command line arguments provided to the application
        /// </summary>
        public string[] Args
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets or sets whether to handle the event of an unhandled exception.
        /// </summary>
        [DefaultValue(true)]
        public bool HandleUnhandledException
        {
            get { return _handleUnhandledException; }
            set { _handleUnhandledException = value; }
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

        #endregion Properties

        #region Methods

        #region Public Methods

        /// <summary>
        /// Starts the application
        /// </summary>
        /// <param name="args">The command line arguments provided to the application</param>
        public void Start(string[] args)
        {
            if (args == null)
            {
                throw new ArgumentException("args must not be null");
            }

            OnStartup();
        }

        #endregion Public Methods

        #region Protected Static Methods

        /// <summary>
        /// Get the missing assembly filename from the given missing item name
        /// </summary>
        /// <param name="itemName">The item to retreive the name for</param>
        /// <returns>The missing assembly filename</returns>
        protected static string GetFileName(string itemName)
        {
            return string.Format("{0}.dll", new AssemblyName(itemName).Name);
        }

        #endregion Protected Static Methods

        #region Protected Methods

        /// <summary>
        /// Called when an assembly is missing. Displays a message to the user and exits the application
        /// </summary>
        /// <param name="args">The missing assembly event arguments</param>
        /// <returns>null to indicate that the assembly could not be found</returns>
        protected virtual Assembly OnResolveAssembly(ResolveEventArgs args)
        {
            string filename = GetFileName(args.Name);

            if (HandleUnresolvedAssembly)
            {
                HandleMissingAssembly(filename);
            }

            return null;
        }

        /// <summary>
        /// Called when the application has started. Override this method to implement the application's functionality
        /// </summary>
        protected virtual void OnStartup()
        {
        }

        /// <summary>
        /// Called when an unhandled exception occur. Displays a message to the user before the application exits
        /// </summary>
        /// <param name="e">The arguments that describe this event</param>
        protected virtual void OnUnhandledException(UnhandledExceptionEventArgs e)
        {
            if (HandleUnhandledException)
            {
                System.Console.WriteLine(UnhandledExceptionMessage);
                Environment.Exit(1);
            }
        }

        #endregion Protected Methods

        #region Private Methods

        /// <summary>
        /// Called when an assembly is missing
        /// </summary>
        /// <param name="sender">Not used</param>
        /// <param name="args">The event handler arguemnts</param>
        /// <returns>The found assembly or null if it cannot be resolved</returns>
        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return OnResolveAssembly(args);
        }

        /// <summary>
        /// Called when an unhandled exception occurs
        /// </summary>
        /// <param name="sender">Not used</param>
        /// <param name="e">The event handler arguemnts</param>
        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            OnUnhandledException(e);
        }

        /// <summary>
        /// Returns the formatted string message, or an empty string if there is an error.
        /// </summary>
        /// <param name="format">A composite format string.</param>
        /// <param name="arg">The value to format.</param>
        /// <returns>The formatted string or string.Empty if there was an error.</returns>
        private string GetFormattedMessage(string format, string arg)
        {
            if (format == null || format == string.Empty)
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

        /// <summary>
        /// Displays an error messge to the user about the missing assembly file and then exits the system.
        /// </summary>
        /// <param name="filename">The filename of the missing assembly</param>
        private void HandleMissingAssembly(string filename)
        {
            System.Console.WriteLine(GetFormattedMessage(HandleMissingAssemblyMessage, filename));
            Environment.Exit(1);
        }

        #endregion Private Methods

        #endregion Methods
    }
}