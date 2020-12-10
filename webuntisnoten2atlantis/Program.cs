﻿// Published under the terms of GPLv3 Stefan Bäumer 2019.

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

                List<string> aktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };

                Console.WriteLine(" Webuntisnoten2atlantis | Published under the terms of GPLv3 | Stefan Bäumer 2020 | Version 20201208");
                Console.WriteLine("=====================================================================================================");
                Console.WriteLine(" Webuntisnoten2atlantis erstellt eine SQL-Datei mit den entsprechenden Befehlen zum Import in      ");
                Console.WriteLine(" Atlantis. ACHTUNG: Wenn der Lehrer es versäumt hat, mindestens 1 Teilleistung zu dokumentieren,   ");
                Console.WriteLine(" wird auch keine Gesamtnote ausgegeben!                                                            ");
                Console.WriteLine("=====================================================================================================");

                CsvDateienVorhanden(inputAbwesenheitenCsv, inputNotenCsv);

                Leistungen alleWebuntisLeistungen = new Leistungen(inputNotenCsv);
                Abwesenheiten alleWebuntisAbwesenheiten = new Abwesenheiten(inputAbwesenheitenCsv);
                Abwesenheiten alleAtlantisAbwesenheiten = new Abwesenheiten(ConnectionStringAtlantis, aktSj[0] + "/" + aktSj[1]);
                Leistungen alleAtlantisLeistungen = new Leistungen(ConnectionStringAtlantis, aktSj[0] + "/" + aktSj[1]);
                Abwesenheiten alleAtlantisAbwesenheitenGefiltert = new Abwesenheiten();
                Leistungen alleAtlantisLeistungenGefiltert = new Leistungen();
                Abwesenheiten alleWebuntisAbwesenheitenGefiltert = new Abwesenheiten();
                Leistungen alleWebuntisLeistungenGefiltert = new Leistungen();

                do
                {

                    List<string> zeugnisart = new List<string>();

                    do
                    {
                        Console.WriteLine("");
                        Console.WriteLine("Geben Sie die gewünschte(n) Zeugnisart(en) (kommagetrennt) ein. Die Zeugnisart steht im Noten-Kopf.");
                        Console.WriteLine("Beispiel: A01HZ oder A01HZ,C03HJ) " + (Properties.Settings.Default.Zeugnisarten == "" ? "" : "[" + Properties.Settings.Default.Zeugnisarten + "] "));

                        var z = Console.ReadLine();

                        if (z.Split(',')[0] == "")
                        {
                            foreach (var item in (Properties.Settings.Default.Zeugnisarten).Split(','))
                            {
                                zeugnisart.Add(item.Trim());
                            }
                        }
                        else
                        {
                            foreach (var item in z.Split(','))
                            {
                                zeugnisart.Add(item.Trim());
                            }
                            Properties.Settings.Default.Zeugnisarten = z;
                            Properties.Settings.Default.Save();
                        }

                        if (zeugnisart.Count > 0)
                        {
                            alleAtlantisAbwesenheitenGefiltert = new Abwesenheiten();
                            alleAtlantisAbwesenheitenGefiltert.AddRange(alleAtlantisAbwesenheiten.Filter(zeugnisart));

                            alleAtlantisLeistungenGefiltert = new Leistungen();
                            alleAtlantisLeistungenGefiltert.AddRange((from a in alleAtlantisLeistungen where zeugnisart.Contains(a.Zeugnisart) select a).ToList());

                            alleWebuntisAbwesenheitenGefiltert = new Abwesenheiten();
                            alleWebuntisAbwesenheitenGefiltert.AddRange(alleWebuntisAbwesenheiten.Filter(zeugnisart, alleAtlantisAbwesenheitenGefiltert));

                            alleWebuntisLeistungenGefiltert = new Leistungen();
                            alleWebuntisLeistungenGefiltert.AddRange(alleWebuntisLeistungen.Filter(zeugnisart, alleAtlantisLeistungenGefiltert));

                            if (alleAtlantisAbwesenheitenGefiltert.Count() == 0 && alleAtlantisLeistungenGefiltert.Count == 0)
                            {
                                Console.WriteLine("");
                                Console.WriteLine("! Es konnte nichts aus Atlantis ausgelesen werden. Es geht nicht weiter. Sind Ihre Angaben richtig?");
                                Console.WriteLine("");
                                Properties.Settings.Default.Zeugnisarten = "";
                                Properties.Settings.Default.Save();
                            }
                        }
                    } while (zeugnisart.Count == 0 || (alleAtlantisAbwesenheitenGefiltert.Count() == 0 && alleAtlantisLeistungenGefiltert.Count == 0));
                    
                    Console.WriteLine("");
                    
                    Leistungen webuntisLeistungen = new Leistungen();
                    Abwesenheiten webuntisAbwesenheiten = new Abwesenheiten();
                    Leistungen atlantisLeistungen = new Leistungen();
                    Abwesenheiten atlantisAbwesenheiten = new Abwesenheiten();

                    var interessierendeKlassen = new List<string>();

                    string outputSql = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\webuntisnoten2atlantis_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".SQL";

                    do
                    {
                        interessierendeKlassen = alleAtlantisLeistungenGefiltert.GetIntessierendeKlassen(alleWebuntisLeistungenGefiltert);
                        string o = "/* " + Properties.Settings.Default.Zeugnisarten + " | " + Properties.Settings.Default.Klassenwahl;
                        Global.Output.Add((o.Substring(0, Math.Min(82, o.Length))).PadRight(82) + "*/");
                        webuntisLeistungen = alleWebuntisLeistungenGefiltert.Filter(interessierendeKlassen);
                        webuntisAbwesenheiten.AddRange((from a in alleWebuntisAbwesenheiten where interessierendeKlassen.Contains(a.Klasse) select a));
                        atlantisLeistungen.AddRange((from a in alleAtlantisLeistungenGefiltert where interessierendeKlassen.Contains(a.Klasse) select a));
                        atlantisAbwesenheiten.AddRange((from a in alleAtlantisAbwesenheitenGefiltert where interessierendeKlassen.Contains(a.Klasse) select a));

                    } while (webuntisLeistungen.Count > 0 ? false : true);

                    // Korrekturen:

                    webuntisLeistungen.Punkte2NoteInAnlageD(atlantisLeistungen);
                    webuntisLeistungen.ReligionKorrigieren();
                    webuntisLeistungen.Religionsabwähler(atlantisLeistungen);
                    webuntisLeistungen.BindestrichfächerZuordnen(atlantisLeistungen);
                    webuntisLeistungen.WeitereFächerZuordnen(atlantisLeistungen); // außer REL, ER, KR

                    atlantisLeistungen.Add(webuntisLeistungen, interessierendeKlassen);
                    atlantisLeistungen.Delete(webuntisLeistungen, interessierendeKlassen);
                    atlantisLeistungen.Update(webuntisLeistungen, interessierendeKlassen);

                    atlantisAbwesenheiten.Add(webuntisAbwesenheiten);
                    atlantisAbwesenheiten.Delete(webuntisAbwesenheiten);
                    atlantisAbwesenheiten.Update(webuntisAbwesenheiten);

                    alleAtlantisLeistungenGefiltert.ErzeugeSqlDatei(outputSql);

                    Console.WriteLine("Weitere Zeugnisarten und Klassen auswählen mit ENTER. Programm beenden mit ESC.");

                } while (Console.ReadKey().Key == ConsoleKey.Enter ? true : false);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Heiliger Bimbam! Es ist etwas schiefgelaufen! Die Verarbeitung wird gestoppt.");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }
        
        private static void CsvDateienVorhanden(string inputAbwesenheitenCsv, string inputNotenCsv)
        {
            if (!File.Exists(inputAbwesenheitenCsv))
            {
                RenderInputAbwesenheitenCsv(inputAbwesenheitenCsv);
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(inputAbwesenheitenCsv).Date != DateTime.Now.Date)
                {
                    RenderInputAbwesenheitenCsv(inputAbwesenheitenCsv);
                }
            }

            if (!File.Exists(inputNotenCsv))
            {
                RenderNotenexportCsv(inputNotenCsv);
            }
            else
            {
                if (System.IO.File.GetLastWriteTime(inputNotenCsv).Date != DateTime.Now.Date)
                {
                    RenderNotenexportCsv(inputNotenCsv);
                }
            }

            Console.WriteLine("");
        }

        private static void RenderInputAbwesenheitenCsv(string inputAbwesenheitenCsv)
        {
            Console.WriteLine("Die Datei " + inputAbwesenheitenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine(" 1. Administration > Export klicken");
            Console.WriteLine(" 2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
            Console.WriteLine(" 4. !!! Zeitraum begrenzen bis zur Zeugniskonferenz !!!");
            Console.WriteLine(" 5. Die Datei \"AbsenceTimesTotal.csv\" auf dem Desktop speichern.");
            Console.WriteLine("ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }

        private static void RenderNotenexportCsv(string inputNotenCsv)
        {
            Console.WriteLine("Die Datei " + inputNotenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine(" 1. Klassenbuch > Berichte klicken");
            Console.WriteLine(" 2. Alle Klassen auswählen");
            Console.WriteLine(" 3. Unter \"Noten\" die Prüfungsart (-Alle-) auswählen");
            Console.WriteLine(" 4. Unter \"Noten\" den Haken bei Notennamen ausgeben _NICHT_ setzen");
            Console.WriteLine(" 5. Hinter \"Noten pro Schüler\" auf CSV klicken.");
            Console.WriteLine(" 6. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern.");
            Console.WriteLine("ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}
