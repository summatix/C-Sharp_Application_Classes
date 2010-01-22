namespace Console
{
    using SuperLibrary;
    using System;

    public class Program : ConsoleApplication
    {
        #region Methods

        #region Protected Methods

        protected override void OnStartup()
        {
            Console.WriteLine("Hello world!");

            Functions.SayHello();

            throw new InvalidOperationException("You should have deleted SuperLibrary.dll first");
        }

        #endregion Protected Methods

        #region Private Static Methods

        private static void Main(string[] args)
        {
            (new Program()).Start(args);
        }

        #endregion Private Static Methods

        #endregion Methods
    }
}