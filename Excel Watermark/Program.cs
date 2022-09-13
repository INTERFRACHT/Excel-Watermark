using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Watermark
{
    class Program
    {
        static void Main(string[] args)
        {
            ErrorHandler ErrorHandler = new ErrorHandler();

            try
            {
                FileHandler fileHandler = new FileHandler();
                
                if (args[0].Length == 0)
                    throw new Exception("Nebyl zadan argument!");

                if (args.Length == 1)
                    fileHandler.ProcessFiles(args[0]);
                else
                    fileHandler.ProcessFiles(args[0], args[1]);
            }
            catch (Exception ex)
            {
                ErrorHandler.SendError("Run", ex.ToString());
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
