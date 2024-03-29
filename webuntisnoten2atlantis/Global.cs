﻿// Published under the terms of GPLv3 Stefan Bäumer 2023.

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
        /// <summary>
        /// Die Rückmeldung ist gedacht, um an Ende Hinweise (über Teams) an die Lehrer zu geben.
        /// </summary>
        public static List<string> Rückmeldung { get; internal set; }
        public static List<string> Tabelle { get; internal set; }
        public static Unterrichte AlleVerschiedenenUnterrichteInDieserKlasseAktuellUnsortiert { get; internal set; }
        public static Unterrichte AlleVerschiedenenUnterrichteInDieserKlasseAktuell { get; internal set; }
        public static Rückmeldungen Rückmeldungen { get; internal set; }

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

        internal static void WriteLineTabelle(string zeile)
        {
            Console.WriteLine(zeile);
            Global.Tabelle.Add(zeile);
        }

        internal static void Write(string zeile)
        {
            Console.Write(zeile);
            Global.PrintMessage(Global.SqlZeilen.Count(), zeile);
        }

        internal static string List2String(List<string> interessierendeKlassen, string delimiter)
        {
            try
            {
                var s = "";

                if (interessierendeKlassen.Count > 0)
                {
                    foreach (var item in interessierendeKlassen)
                    {
                        s += item + delimiter;

                        if (s.Length > Global.PadRight - 5 && s.Length <= Global.PadRight || s.Length > Global.PadRight * 2 - 5 && s.Length <= Global.PadRight * 2 || s.Length > Global.PadRight * 3 - 5 && s.Length <= Global.PadRight * 3 || s.Length > Global.PadRight * 4 - 5 && s.Length <= Global.PadRight * 4)
                        {
                            s += "\n";
                        }
                    }
                }

                return s.Substring(0, s.Length - delimiter.Length);
            }
            catch (Exception)
            {
                return "";
            }
        }

        internal static string List2String90(List<string> interessierendeKlassen, string delimiter)
        {
            var s = "";

            if (interessierendeKlassen.Count > 0)
            {
                foreach (var item in interessierendeKlassen)
                {
                    s += item + delimiter;

                    if (s.Length > 90 - 7 && s.Length <= 90 || s.Length > 90 * 2 - 7 && s.Length <= 90 * 2 || s.Length > 90 * 3 - 7 && s.Length <= 90 * 3 || s.Length > 60 * 4 - 7 && s.Length <= 60 * 4)
                    {
                        s += "\n                  ";
                    }
                }
            }

            return s.Substring(0, s.Length - delimiter.Length);
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