namespace ComRega
{
    using System;

    public static class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "COM Startup App";
            try
            {
                if (args[0].EndsWith("exe"))
                {
                    Console.WriteLine("Select an option:");
                    Console.WriteLine("1 - [COM] Add to Startup");
                    Console.WriteLine("2 - [COM] Remove from Startup\n");

                    Console.Write("Your option: ");
                    int choice;
                    while (!int.TryParse(Console.ReadLine(), out choice))
                    {
                        Console.WriteLine("Please enter a number:");
                    }

                    switch (choice)
                    {
                        case 1: COMStartup.ComAddToStartup(true, args[0], "BlackJust", "Windows Service"); break;
                        case 2: COMStartup.ComAddToStartup(false, args[0], "BlackJust", "Windows Service"); break;
                        default: Console.WriteLine("Please select option 1 or 2.\n"); break;
                    }

                    Console.WriteLine("\nPress any key to close the program...");
                    Console.ReadKey();
                }
            }
            catch 
            {
                Console.WriteLine("Only extension .exe file");
                Console.ReadKey();
            }
        }
    }
}
