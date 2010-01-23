namespace Console
{
    using System;

    using SuperLibrary;

    public sealed class Program : ConsoleApplication
    {
        protected override void OnStartup()
        {
            Console.WriteLine("Hello world!");

            Functions.SayHello();

            throw new InvalidOperationException("You should have deleted SuperLibrary.dll first");
        }

        private static void Main(string[] args)
        {
            (new Program()).Start(args);
        }
    }
}