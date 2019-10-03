// Published under the terms of GPLv3 Stefan Bäumer 2019.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

                Console.WriteLine(" Webuntisnoten2atlantis | Published under the terms of GPLv3 | Stefan Bäumer 2019 | Version 20191003");
                Console.WriteLine("====================================================================================================");
                

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

                Console.WriteLine("");
                Schlüssels schlüssels = new Schlüssels(ConnectionStringAtlantis);
                Leistungen alleWebuntisLeistungen = new Leistungen(inputNotenCsv);
                Abwesenheiten alleWebuntisAbwesenheiten = new Abwesenheiten(inputAbwesenheitenCsv);
                Abwesenheiten atlantisAbwesenheiten = new Abwesenheiten(ConnectionStringAtlantis, aktSj[0] + "/" + aktSj[1], alleWebuntisLeistungen[0].Prüfungsart, schlüssels);
                Leistungen atlantisLeistungen = new Leistungen(ConnectionStringAtlantis, aktSj[0] + "/" + aktSj[1], alleWebuntisLeistungen[0].Prüfungsart, schlüssels);
                                
                Console.WriteLine("");
                
                do
                {
                    List<Leistung> webuntisLeistungen;
                    List<Abwesenheit> webuntisAbwesenheiten;

                    var interessierendeKlassen = new List<string>();

                    string outputSql = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\webuntisnoten2atlantis_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".SQL";

                    do
                    {                        
                        interessierendeKlassen = atlantisLeistungen.GetIntessierendeKlassen(alleWebuntisLeistungen);
                        webuntisLeistungen = (from a in alleWebuntisLeistungen where interessierendeKlassen.Contains(a.Klasse) select a).ToList();
                        webuntisAbwesenheiten = (from a in alleWebuntisAbwesenheiten where interessierendeKlassen.Contains(a.Klasse) select a).ToList();

                    } while (webuntisLeistungen.Count > 0 ? false : true);

                    atlantisLeistungen.Add(webuntisLeistungen);
                    atlantisLeistungen.Delete(webuntisLeistungen);
                    atlantisLeistungen.Update(webuntisLeistungen);
                    
                    atlantisAbwesenheiten.Add(webuntisAbwesenheiten);
                    atlantisAbwesenheiten.Delete(webuntisAbwesenheiten);
                    atlantisAbwesenheiten.Update(webuntisAbwesenheiten);

                    atlantisLeistungen.ErzeugeSqlDatei(outputSql);
                    
                    Console.WriteLine("Weitere Klassen auswählen dann . Beenden mit ESC.");

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

        private static void RenderInputAbwesenheitenCsv(string inputAbwesenheitenCsv)
        {
            Console.WriteLine("Die Datei " + inputAbwesenheitenCsv + " existiert nicht.");
            Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
            Console.WriteLine(" 1. Administration > Export klicken");
            Console.WriteLine(" 2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
            Console.WriteLine(" 3. Die Datei \"AbsenceTimesTotal.csv\" auf dem Desktop speichern.");
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
            Console.WriteLine(" 3. Unter \"Noten\" die Prüfungsart (z.B. Halbjahreszeugnis) auswählen");
            Console.WriteLine(" 4. Hinter \"Noten pro Schüler\" auf CSV klicken.");
            Console.WriteLine(" 5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern.");
            Console.WriteLine("ENTER beendet das Programm.");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}
