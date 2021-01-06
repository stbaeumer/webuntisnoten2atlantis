// Published under the terms of GPLv3 Stefan Bäumer 2020.

using System;
using System.Collections.Generic;
using System.IO;

namespace webuntisnoten2atlantis
{
    internal static class Global
    {
        public static List<string> Output { get; internal set; }

        internal static void PrintMessage(string message)
        {
            if (message.ToLower().Contains("zu"))
            {
                Output.Add("");
            }
            Output.Add("/* " + message.PadRight(97) + " */");
            if (message.ToLower().Contains("zu"))
            {
                Output.Add("");
            }
        }

        public static bool WaitForFile(string fullPath)
        {
            int anzahlVersuche = 0;
            while (true)
            {
                ++anzahlVersuche;
                try
                {
                    // Attempt to open the file exclusively.
                    using (FileStream fs = new FileStream(fullPath,
                        FileMode.Open, FileAccess.ReadWrite,
                        FileShare.None, 100))
                    {
                        fs.ReadByte();

                        // If we got this far the file is ready
                        break;
                    }
                }
                catch (FileNotFoundException ex)
                {
                    Console.WriteLine("Die Datei {0} soll jetzt eingelsen werden, existiert aber nicht.", fullPath);

                    if (anzahlVersuche > 10)
                        Console.WriteLine("Bitte Programm beenden.");

                    System.Threading.Thread.Sleep(4000);
                }
                catch (IOException ex)
                {
                    Console.WriteLine("Die Datei {0} ist gesperrt", fullPath);

                    if (anzahlVersuche > 10)
                        Console.WriteLine("Bitte Programm beenden.");

                    System.Threading.Thread.Sleep(4000);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            return true;
        }
    }
}