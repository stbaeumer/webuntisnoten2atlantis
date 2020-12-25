﻿// Published under the terms of GPLv3 Stefan Bäumer 2020.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace webuntisnoten2atlantis
{
    class Program
    {
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=DBA";

        static void Main(string[] args)
        {
            Global.Output = new List<string>();

            try
            {
                string inputNotenCsv = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MarksPerLesson.csv";
                string inputAbwesenheitenCsv = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\AbsenceTimesTotal.csv";
                string zeitstempel = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                List<string> aktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };

                Console.WriteLine(" Webuntisnoten2Atlantis | Published under the terms of GPLv3 | Stefan Bäumer " + DateTime.Now.Year + " | Version 20201222");
                Console.WriteLine("=====================================================================================================");
                Console.WriteLine(" *Webuntisnoten2Atlantis* erstellt eine SQL-Datei mit entsprechenden Befehlen zum Import in Atlantis.");
                Console.WriteLine(" ACHTUNG: Wenn der Lehrer es versäumt hat, mindestens 1 Teilleistung zu dokumentieren, wird keine Ge- ");
                Console.WriteLine(" samtnote von Webuntis übergeben!");
                Console.WriteLine("=====================================================================================================");

                CheckCsv(inputAbwesenheitenCsv, inputNotenCsv);

                string pfad = SetPfad(zeitstempel, inputAbwesenheitenCsv, inputNotenCsv);

                string outputSql = pfad + "\\webuntisnoten2atlantis_" + zeitstempel + ".SQL";

                Leistungen alleWebuntisLeistungen = new Leistungen(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\MarksPerLesson.csv");
                Abwesenheiten alleWebuntisAbwesenheiten = new Abwesenheiten(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\AbsenceTimesTotal.csv");
                Abwesenheiten alleAtlantisAbwesenheiten = new Abwesenheiten(ConnectionStringAtlantis, aktSj[0] + "/" + aktSj[1]);
                Leistungen alleAtlantisLeistungen = new Leistungen(ConnectionStringAtlantis, aktSj);
                                
                do
                {                    
                    Leistungen webuntisLeistungen = new Leistungen();
                    Abwesenheiten webuntisAbwesenheiten = new Abwesenheiten();
                    Leistungen atlantisLeistungen = new Leistungen();
                    Abwesenheiten atlantisAbwesenheiten = new Abwesenheiten();

                    var interessierendeKlassen = new List<string>();

                    do
                    {
                        interessierendeKlassen = alleAtlantisLeistungen.GetIntessierendeKlassen(alleWebuntisLeistungen, aktSj);
                        
                        webuntisLeistungen.AddRange((from a in alleWebuntisLeistungen
                                                     where interessierendeKlassen.Contains(a.Klasse)
                                                     select a));
                        webuntisAbwesenheiten.AddRange((from a in alleWebuntisAbwesenheiten                                                        
                                                        where interessierendeKlassen.Contains(a.Klasse)
                                                        select a));
                        atlantisLeistungen.AddRange((from a in alleAtlantisLeistungen                                                     
                                                     where interessierendeKlassen.Contains(a.Klasse)
                                                     select a));
                        atlantisAbwesenheiten.AddRange((from a in alleAtlantisAbwesenheiten                                                        
                                                        where interessierendeKlassen.Contains(a.Klasse)
                                                        select a));

                        if (webuntisLeistungen.Count == 0)
                        {
                            Console.WriteLine("[!] Es liegt kein einziger Leistungsdatensatz für Ihre Auswahl vor. Ist evtl. die Auswahl in Webuntis eingeschränkt? ");
                        }

                    } while (webuntisLeistungen.Count <= 0);

                    // Alte Noten holen
                                        
                    webuntisLeistungen.AddRange(alleAtlantisLeistungen.HoleAlteNoten(webuntisLeistungen, interessierendeKlassen, aktSj));

                    // Korrekturen
                                        
                    webuntisLeistungen.ReligionKorrigieren();
                    webuntisLeistungen.Religionsabwähler(atlantisLeistungen);
                    webuntisLeistungen.BindestrichfächerZuordnen(atlantisLeistungen);
                    webuntisLeistungen.SprachenZuordnen(atlantisLeistungen);
                    webuntisLeistungen.WeitereFächerZuordnen(atlantisLeistungen); // außer REL, ER, KR, Bindestrich-Fächer                    
                    // Sortieren

                    webuntisLeistungen.OrderBy(x => x.Klasse).ThenBy(x => x.Name);
                    atlantisLeistungen.OrderBy(x => x.Klasse).ThenBy(x => x.Name);

                    // Add-Delete-Update

                    atlantisLeistungen.Add(webuntisLeistungen, interessierendeKlassen);
                    atlantisLeistungen.Delete(webuntisLeistungen, interessierendeKlassen, aktSj);
                    atlantisLeistungen.Update(webuntisLeistungen, interessierendeKlassen);

                    atlantisAbwesenheiten.Add(webuntisAbwesenheiten);
                    atlantisAbwesenheiten.Delete(webuntisAbwesenheiten);
                    atlantisAbwesenheiten.Update(webuntisAbwesenheiten);

                    alleAtlantisLeistungen.ErzeugeSqlDatei(outputSql);

                    Console.WriteLine("");
                    Console.WriteLine("");
                    Console.WriteLine("  Programm beenden mit Enter.");

                } while (Console.ReadKey().Key == ConsoleKey.Escape);                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ooops! Es ist etwas schiefgelaufen! Die Verarbeitung wird gestoppt.");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private static string SetPfad(string zeitstempel, string inputNotenCsv, string inputAbwesenheitenCsv)
        {
            var pfad = @"\\fs01\Schulverwaltung\webuntisnoten2atlantis\Dateien";

            if (Properties.Settings.Default.Pfad != "")
            {
                pfad = Properties.Settings.Default.Pfad;
            }

            if (!Directory.Exists(pfad))
            {
                do
                {
                    Console.WriteLine(" Wo sollen die Dateien gespeichert werden? [ " + pfad + " ]");
                    pfad = Console.ReadLine();
                    if (pfad == "")
                    {
                        pfad = @"\\fs01\Schulverwaltung\webuntisnoten2atlantis\Dateien";
                    }
                    try
                    {
                        Directory.CreateDirectory(pfad);
                        Properties.Settings.Default.Pfad = pfad;
                    }
                    catch (Exception)
                    {

                        Console.WriteLine("Der Pfad " + pfad + " kann nicht angelegt werden.");
                    }

                } while (!Directory.Exists(pfad));
            }

#if DEBUG
            
            File.Copy(inputNotenCsv, pfad + @"\" + zeitstempel + "-MarksPerLesson.csv");
            File.Copy(inputAbwesenheitenCsv, pfad + @"\" + zeitstempel + "-AbsenceTimesTotal.csv");
#else
            File.Copy(inputNotenCsv, pfad + @"\" + zeitstempel + "-MarksPerLesson.csv");
            File.Copy(inputAbwesenheitenCsv, pfad + @"\" + zeitstempel + "-AbsenceTimesTotal.csv");
#endif




            return pfad;
        }

        private static void CheckCsv(string inputAbwesenheitenCsv, string inputNotenCsv)
        {
            if (!File.Exists(inputAbwesenheitenCsv))
            {
                RenderInputAbwesenheitenCsv(inputAbwesenheitenCsv, "existiert nicht");
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(inputAbwesenheitenCsv).Date != DateTime.Now.Date)
                {
                    RenderInputAbwesenheitenCsv(inputAbwesenheitenCsv, "ist nicht aktuell von heute");
                }
            }

            if (!File.Exists(inputNotenCsv))
            {
                RenderNotenexportCsv(inputNotenCsv, "existiert nicht");
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(inputNotenCsv).Date != DateTime.Now.Date)
                {
                    RenderNotenexportCsv(inputNotenCsv, "ist nicht aktuell von heute");
                }
            }

            Console.WriteLine("");
        }

        private static void RenderInputAbwesenheitenCsv(string inputAbwesenheitenCsv, string meldung)
        {
            Console.WriteLine("");
            Console.WriteLine("  Die Datei " + inputAbwesenheitenCsv + " " + meldung + ".");
            Console.WriteLine("  Exportieren Sie die Datei frisch aus dem Digitalen Klassenbuch, indem Sie als Administrator:");
            Console.WriteLine("   1. Administration > Export klicken");
            Console.WriteLine("   2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
            Console.WriteLine("   4. !!! Zeitraum begrenzen (z.B. bis drei Tage vor der) Zeugniskonferenz !!!");
            Console.WriteLine("   5. Die Datei \"AbsenceTimesTotal.csv\" auf dem Desktop speichern");
            Console.WriteLine("");
            Console.WriteLine(" ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }

        private static void RenderNotenexportCsv(string inputNotenCsv, string meldung)
        {
            Console.WriteLine("");
            Console.WriteLine("  Die Datei " + inputNotenCsv + " " + meldung + ".");
            Console.WriteLine("  Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine("   1. Klassenbuch > Berichte klicken");
            Console.WriteLine("   2. Alle Klassen auswählen");
            Console.WriteLine("   3. Unter \"Noten\" die Prüfungsart (-Alle-) auswählen");
            Console.WriteLine("   4. Unter \"Noten\" den Haken bei Notennamen ausgeben _NICHT_ setzen");
            Console.WriteLine("   5. Hinter \"Noten pro Schüler\" auf CSV klicken");
            Console.WriteLine("   6. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern");
            Console.WriteLine("");
            Console.WriteLine("  ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}
