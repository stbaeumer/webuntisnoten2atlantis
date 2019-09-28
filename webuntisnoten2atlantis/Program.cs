using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace webuntisnoten2atlantis
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> aktSj = new List<string>
            {
                (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString()
            };

            Console.WriteLine("Webuntisnoten2atlantis (Version 20190914)");
            Console.WriteLine("====================================");
            Console.WriteLine("");


        }
    }
}
