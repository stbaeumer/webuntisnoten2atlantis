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
        public static List<string> VerschiedeneKlassenAusMarkPerLesson { get; set; }
        public static int PadRight { get; internal set; }
        public static Leistungen LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum { get; internal set; }
        public static Leistungen Notenblatt { get; internal set; }
        public static Abwesenheiten AtlantisAbwesenheiten { get; internal set; }        
        public static bool BlaueBriefe { get; internal set; }
        public static List<string> Reihenfolge { get; set; }

        internal static void PrintMessage(int index, string message)
        {
            SqlZeilen.Insert(index, "/* " + message.PadRight(Global.PadRight + 61) + " */");
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
                    Global.WriteLine("Die Datei " + fullPath + " soll jetzt eingelsen werden, existiert aber nicht.");

                    if (anzahlVersuche > 10)
                        Global.WriteLine("Bitte Programm beenden.");

                    System.Threading.Thread.Sleep(4000);
                }
                catch (IOException)
                {
                    Global.WriteLine("Die Datei " + fullPath + " ist gesperrt");

                    if (anzahlVersuche > 10)
                        Global.WriteLine("Bitte Programm beenden.");

                    System.Threading.Thread.Sleep(4000);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            return true;
        }

        internal static void WriteLine(string zeile)
        {
            Console.WriteLine(zeile);            
            Global.PrintMessage(Global.SqlZeilen.Count(), zeile);
        }

        internal static void Write(string zeile)
        {
            Console.Write(zeile);
            Global.PrintMessage(Global.SqlZeilen.Count(), zeile);
        }

        internal static string List2String(List<string> interessierendeKlassen, Char delimiter)
        {
            var s = "";
            foreach (var item in interessierendeKlassen)
            {
                s += item + delimiter;
            }
            return s.TrimEnd(delimiter);
        }
        internal static string List2String(List<int> interessierendeKlassen, Char delimiter)
        {
            var s = "";
            foreach (var item in interessierendeKlassen)
            {
                s += item.ToString() + delimiter;
            }
            return s.TrimEnd(delimiter);
        }

        internal static void SetzeReihenfolgeDerFächer(Leistungen aleistungen)
        {
            Reihenfolge = new List<string>();

            foreach (var al in aleistungen.OrderBy(x=>x.Reihenfolge))
            {
                if (!Reihenfolge.Contains(al.Fach))
                {
                    Reihenfolge.Add(al.Fach);
                }
            }
        }
    }
}