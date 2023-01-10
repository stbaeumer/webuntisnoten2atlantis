// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace webuntisnoten2atlantis
{
    internal static class Global
    {
        public static Leistungen Defizitleistungen { get; internal set; }

        public static List<string> SqlZeilen { get; internal set; }
        public static string HzJz { get; internal set; }
        public static List<string> VerschiedeneKlassenAusMarkPerLesson { get; set; }

        internal static void PrintMessage(int index, string message)
        {
            if (message.ToLower().Contains("zu"))
            {
                SqlZeilen.Insert(index, "");
            }

            SqlZeilen.Insert(index, "/* " + message.PadRight(97) + " */");
            
            if (message.ToLower().Contains("zu"))
            {
                SqlZeilen.Insert(index, "");
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
                catch (FileNotFoundException)
                {
                    Console.WriteLine("Die Datei {0} soll jetzt eingelsen werden, existiert aber nicht.", fullPath);

                    if (anzahlVersuche > 10)
                        Console.WriteLine("Bitte Programm beenden.");

                    System.Threading.Thread.Sleep(4000);
                }
                catch (IOException)
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

        public static void AusgabeSchreiben(string text, List<string> klassen)
        {
            Global.SqlZeilen.Add("");
            int z = 0;

            do
            {
                var zeile = "";

                try
                {
                    while ((zeile + text.Split(' ')[z] + ", ").Length <= 96)
                    {
                        zeile += text.Split(' ')[z] + " ";
                        z++;
                    }
                }
                catch (Exception)
                {
                    z++;
                    zeile.TrimEnd(',');
                }

                zeile = zeile.TrimEnd(' ');

                string o = "/* " + zeile.TrimEnd(' ');
                Global.SqlZeilen.Add((o.Substring(0, Math.Min(101, o.Length))).PadRight(101) + "*/");

            } while (z < text.Split(' ').Count());



            z = 0;

            do
            {
                var zeile = " ";

                try
                {
                    if (klassen[z].Length >= 95)
                    {
                        klassen[z] = klassen[z].Substring(0, Math.Min(klassen[z].Length, 95));
                        zeile += klassen[z];
                        throw new Exception();
                    }

                    while ((zeile + klassen[z] + ", ").Length <= 97)
                    {
                        zeile += klassen[z] + ", ";
                        z++;
                    }
                }
                catch (Exception)
                {
                    z++;
                    zeile.TrimEnd(',');
                }

                zeile = zeile.TrimEnd(' ');
                int s = zeile.Length;
                string o = "/* " + zeile.TrimEnd(',');
                Global.SqlZeilen.Add((o.Substring(0, Math.Min(101, o.Length))).PadRight(101) + "*/");

            } while (z < klassen.Count);
        }
    }
}