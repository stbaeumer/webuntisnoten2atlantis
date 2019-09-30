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
            try
            {
                string inputCsv = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MarksPerLesson.csv";
                                
                List<string> aktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };

                Console.WriteLine("Webuntisnoten2atlantis (Version " + DateTime.Now.ToString("yyyyMMdd") + ")");
                Console.WriteLine("=========================================");
                Console.WriteLine("");
                Console.WriteLine("Voraussetzungen:");
                Console.WriteLine("1. Fächer- und Klassenbezeichnungen sind identisch in Untis und Atlantis.");
                Console.WriteLine("2. Ein Zeugnisdatensatz für die jeweilige Klasse ist angelegt:");
                Console.WriteLine("   1. Zeugnisse>Sammelbearbeitung");
                Console.WriteLine("   2. Klasse wählen und dann mit 'Alle auswählen' die SuS wählen.");
                Console.WriteLine("   3. Zeugnissätze N (Notenblatt) und HZ (Halbjahreszeugnis) anklicken.");
                Console.WriteLine("   4. 'Zeugnissätze anlegen (N, HZ)' klicken. Dann 'Funktion starten' klicken.");                
                Console.WriteLine("3. Noten aus Webuntis exportieren:");
                Console.WriteLine("   1. Klasenbuch > Berichte");
                Console.WriteLine("   2. Alle Klassen auswählen");
                Console.WriteLine("   3. Unter \"Noten\" Prüfungsart HZ auswählen");
                Console.WriteLine("   4. Hinter \"Noten pro Schüler\" auf CSV klicken.");
                Console.WriteLine("   5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern.");

                if (!File.Exists(inputCsv))
                {
                    Console.WriteLine("Die Datei " + inputCsv + " existiert nicht.");
                    Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
                    Console.WriteLine(" 1. Klasenbuch > Berichte");
                    Console.WriteLine(" 2. Alle Klassen auswählen");
                    Console.WriteLine(" 3. Unter \"Noten\" Prüfungsart HZ auswählen");
                    Console.WriteLine(" 4. Hinter \"Noten pro Schüler\" auf CSV klicken.");
                    Console.WriteLine(" 5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern.");
                    Console.WriteLine("ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
                else
                {
                    if (System.IO.File.GetLastWriteTime(inputCsv).Date != DateTime.Now.Date)
                    {
                        Console.WriteLine("Die Datei " + inputCsv + " ist nicht von heute.");
                        Console.WriteLine("Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie");
                        Console.WriteLine(" 1. Klasenbuch > Berichte");
                        Console.WriteLine(" 2. Alle Klassen auswählen");
                        Console.WriteLine(" 3. Unter \"Noten\" Prüfungsart HZ auswählen");
                        Console.WriteLine(" 4. Hinter \"Noten pro Schüler\" auf CSV klicken.");
                        Console.WriteLine(" 5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern.");
                        Console.WriteLine("ENTER beendet das Programm.");
                        Console.ReadKey();
                        Environment.Exit(0);
                    }
                }

                Console.WriteLine("");
                Leistungen alleWebuntisLeistungen = new Leistungen(inputCsv);                
                Leistungen atlantisLeistungen = new Leistungen(ConnectionStringAtlantis, aktSj[0] + "/" + aktSj[1], alleWebuntisLeistungen[0].Prüfungsart);
                Leistungen webuntisLeistungen;

                Console.WriteLine("");

                var interessierendeKlassen = new List<string>();

                do
                {
                    string outputSql = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\webuntisnoten2atlantis_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".SQL";

                    do
                    {
                        webuntisLeistungen = new Leistungen();

                        try
                        {
                            Console.WriteLine("********************************************************************************************************************");
                            Console.WriteLine("*                                                                                                                  *");
                            Console.WriteLine("*  Geben Sie die interessierende(n) Klasse(n) ein (z. B. HH oder HHU oder HHU1 oder HHU1,HHU2). Dann ENTER.        *");
                            Console.WriteLine("*  Oder 10 Sekunden warten, um alle " + (from a in atlantisLeistungen select a.Klasse).Distinct().Count().ToString().PadLeft(3) + " Klassen mit angelegtem Notenblatt und                                      *");
                            Console.WriteLine("*  zugewiesenem " + alleWebuntisLeistungen[0].Prüfungsart + "-Zeugnisformular zu wählen:                                                       *");
                            Console.WriteLine("*                                                                                                                  *");
                            Console.WriteLine("********************************************************************************************************************");
                            Console.WriteLine("");
                            Console.Write("Ihre Wahl: ");
                            string eingabe = Reader.ReadLine(15000).ToUpper();
                            Console.WriteLine("");
                            // Wenn die Eingabe nicht leer ist, wird die Lehrerliste geleert.

                            if (eingabe != "")
                            {
                                List<string> alleKlassen = (from k in atlantisLeistungen select k.Klasse).Distinct().ToList();

                                var interessierendeKlassenString = "";
                                var klassenOhneNotenblattString = "";
                                var klassenOhneLeistungsdatensätzeString = "";

                                foreach (var klasse in alleKlassen)
                                {
                                    var x = (from k in eingabe.Split(',') where klasse.StartsWith(k) select k).FirstOrDefault();

                                    if (x != null)
                                    {
                                        // Klassen, die ins Suchmuster passen, aber ohne Noten in Webuntis

                                        if (!(from w in alleWebuntisLeistungen where w.Klasse == klasse select w).Any())
                                        {
                                            klassenOhneLeistungsdatensätzeString += klasse + ",";
                                        }

                                        var z = (from w in alleWebuntisLeistungen where w.Klasse == klasse select w).ToList();

                                        if (z.Count > 0)
                                        {
                                            webuntisLeistungen.AddRange(z);
                                            interessierendeKlassen.Add(x);
                                            interessierendeKlassenString += klasse + ",";
                                        }
                                    }
                                }

                                if (klassenOhneLeistungsdatensätzeString != "")
                                {
                                    Console.WriteLine("[!] Folgende Klassen passen in Ihr Suchmuster, allerdings liegen in Webuntis keine Leistungsdatensätze vor:\n    " + klassenOhneLeistungsdatensätzeString.TrimEnd(',') + "\n");
                                }

                                foreach (var w in alleWebuntisLeistungen)
                                {
                                    if (!(from x in atlantisLeistungen where x.Klasse == w.Klasse select x).Any())
                                    {
                                        klassenOhneNotenblattString += w.Klasse + ",";
                                    }
                                }

                                if (klassenOhneNotenblattString != "")
                                {
                                    Console.WriteLine("[!] Folgende Klassen passen in Ihr Suchmuster, allerdings ist kein Notenblatt in Atlantis angelegt:\n    " + klassenOhneNotenblattString.TrimEnd(',') + "\n");

                                }

                                if (interessierendeKlassenString == "")
                                {
                                    Console.WriteLine("Es ist keine einzige Klasse bereit zur Verarbeitung.");
                                }
                                else
                                {
                                    Console.WriteLine("Klassen bereit zur Verarbeitung: \n    " + interessierendeKlassenString.TrimEnd(','));
                                }
                            }

                            Console.WriteLine("");
                        }
                        catch (TimeoutException)
                        {
                            Console.WriteLine("");
                            Console.WriteLine("Ihre Auswahl: Alle " + (from a in atlantisLeistungen select a.Klasse).Distinct().Count().ToString().PadLeft(2) + " Klassen, in denen ein Notenblatt angelegt ist.");
                            webuntisLeistungen = alleWebuntisLeistungen;
                        }
                    } while (webuntisLeistungen.Count > 0 ? false : true);

                    Console.WriteLine("");
                    Console.WriteLine("Weiter mit ENTER");
                    Console.ReadKey();

                    atlantisLeistungen.NeuZuSetzendeNoten(webuntisLeistungen);
                    atlantisLeistungen.ZuLöschendeNoten(webuntisLeistungen);
                    atlantisLeistungen.ZuÄnderendeNoten(webuntisLeistungen);

                    atlantisLeistungen.ErzeugeSqlDatei(outputSql);

                    Console.WriteLine("Weitere Klassen mit ENTER auswählen. Beenden mit ANYKEY.");
                    

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
            
    }
}
