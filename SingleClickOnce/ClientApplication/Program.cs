using System;
using Common;

namespace ClientApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            string environment = AppSettings.Environment.ToString();
            Console.WriteLine("Hello! Yippee! It worked!! ");
            Console.WriteLine($"{nameof(environment)} is {environment}!");
            Console.WriteLine("Press any key to close ...");
            Console.ReadKey();
        }
    }
}

