// Published under the terms of GPLv3 Stefan Bäumer 2020.

using Microsoft.Exchange.WebServices.Data;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.IO;
using System.Linq;

namespace webuntisnoten2atlantis
{
    class Program
    {
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=";

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

                Console.WriteLine(" Webuntisnoten2Atlantis | Published under the terms of GPLv3 | Stefan Bäumer " + DateTime.Now.Year + " | Version 20201230");
                Console.WriteLine("=====================================================================================================");
                Console.WriteLine(" *Webuntisnoten2Atlantis* erstellt eine SQL-Datei mit entsprechenden Befehlen zum Import in Atlantis.");
                Console.WriteLine(" ACHTUNG: Wenn der Lehrer es versäumt hat, mindestens 1 Teilleistung zu dokumentieren, wird keine Ge-");
                Console.WriteLine(" samtnote von Webuntis übergeben!");
                Console.WriteLine("=====================================================================================================");

                if (Properties.Settings.Default.DBUser == "" || Properties.Settings.Default.Klassenart == "")
                {
                    Settings();
                }
                
                CheckCsv(inputAbwesenheitenCsv, inputNotenCsv);

                string pfad = SetPfad(zeitstempel, inputAbwesenheitenCsv, inputNotenCsv);

                string outputSql = pfad + "\\webuntisnoten2atlantis_" + zeitstempel + ".SQL";

                Leistungen alleAtlantisLeistungen = new Leistungen(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, aktSj);
                Leistungen alleWebuntisLeistungen = new Leistungen(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\MarksPerLesson.csv");                
                Abwesenheiten alleAtlantisAbwesenheiten = new Abwesenheiten(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, aktSj[0] + "/" + aktSj[1]);
                Abwesenheiten alleWebuntisAbwesenheiten = new Abwesenheiten(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\AbsenceTimesTotal.csv");

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
                        
                        webuntisLeistungen.AddRange((from a in alleWebuntisLeistungen where interessierendeKlassen.Contains(a.Klasse) select a));
                        webuntisAbwesenheiten.AddRange((from a in alleWebuntisAbwesenheiten where interessierendeKlassen.Contains(a.Klasse) select a));
                        atlantisLeistungen.AddRange((from a in alleAtlantisLeistungen where interessierendeKlassen.Contains(a.Klasse) select a));
                        atlantisAbwesenheiten.AddRange((from a in alleAtlantisAbwesenheiten where interessierendeKlassen.Contains(a.Klasse) select a));

                        if (webuntisLeistungen.Count == 0)
                        {
                            Console.WriteLine("[!] Es liegt kein einziger Leistungsdatensatz für Ihre Auswahl vor. Ist evtl. die Auswahl in Webuntis eingeschränkt? ");
                        }

                    } while (webuntisLeistungen.Count <= 0);

                    // Alte Noten holen
                                        
                    webuntisLeistungen.AddRange(alleAtlantisLeistungen.HoleAlteNoten(webuntisLeistungen, interessierendeKlassen, aktSj));

                    // Korrekturen
                                        
                    webuntisLeistungen.ReligionZuordnen();
                    webuntisLeistungen.ReligionsabwählerBehandeln(atlantisLeistungen);
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

                    Console.Write("Sollen Meetings in Outlook angelegt werden? (j/n) " + (Properties.Settings.Default.Meeting.ToLower() == "j" ? "[j]" : "[n]"));
                                        
                    var meeting = Console.ReadKey();

                    if (meeting.Key != ConsoleKey.J && Properties.Settings.Default.Meeting != "j")
                    {
                        Properties.Settings.Default.Meeting = "n";
                        Properties.Settings.Default.Save();
                        
                    }
                    if (meeting.Key == ConsoleKey.J || Properties.Settings.Default.Meeting == "j")
                    {
                        Properties.Settings.Default.Meeting = "j";
                        Properties.Settings.Default.Save();
                        // Die Exceldatei des anstehenden Abschnitts (HZ/JZ) wird durchlaufen ...
                        Console.WriteLine("");
                        Lehrers lehrers = new Lehrers(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, aktSj);

                        var konferenzen = ReadExcel(@"C:\Users\bm\Berufskolleg Borken\Kollegium - General\03 Schulleitung\3.04 Termine\2020-21\14 Zeugniskonferenzen HZ\\Zeugniskonferenzen HZ.xlsx");

                        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

                        Console.Write("Bitte das O365-Kennwort für " + System.Security.Principal.WindowsIdentity.GetCurrent().Name + " eingeben:");
                        string pw = Console.ReadLine();
                        Console.WriteLine("");
                        string mail = (from l in lehrers where System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1] == l.Kuerzel select l.Mail).FirstOrDefault();
                        service.Credentials = new WebCredentials(mail, pw);
                        service.UseDefaultCredentials = false;
                        service.AutodiscoverUrl(mail, RedirectionUrlValidationCallback);

                        foreach (var konferenz in konferenzen)
                        {                                         
                            var le = (from l in lehrers where System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1] == l.Kuerzel select l).FirstOrDefault();
                            konferenz.Lehrers = alleWebuntisLeistungen.LehrerDieserKlasse(konferenz, lehrers);
                            le.ToOutlook(service, konferenz);
                        }

                        // Fehlende Klassen in der Exceldatei werden ausgegeben



                        // Überzählige Klassen in der Excelliste werden ausgegeben



                        // Alle Klassen mit anstehenden Zeugniskonferenzen werden ermittelt:



                        // Alle Klassen mit anstehenden Zeugniskonferenzen werden durchlaufen ...



                        // ... und alle verschiedenen Lehrer werden ermittelt.


                        // ... und jeweils ein Appointment eingetragen


                    }

                    alleAtlantisLeistungen.ErzeugeSqlDatei(outputSql);

                    Console.WriteLine("");
                    Console.WriteLine("  -----------------------------------------------------------------");
                    Console.WriteLine("  Verarbeitung abgeschlossen. Programm beenden mit Enter.");
                    Environment.Exit(0);
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

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        public static Termine ReadExcel(string pfad)
        {
            Termine termine = new Termine();

            if (File.Exists(pfad))
            {
                Global.WaitForFile(pfad);
            }
            else
            {
                Console.WriteLine("Die Datei " + pfad + ".xlsx soll jetzt eingelesen werden, kann aber nicht gefunden werden.");
            }

            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(pfad);
            Worksheet worksheet = (Worksheet)workbook.Worksheets.get_Item(1);
            Range xlRange = worksheet.UsedRange;

            try
            {                
                Console.Write("Lese Exceldatei " + pfad + " ... ");

                // InSpalte D und K stehen die verschiednenen Klassen

                int rowCount = xlRange.Rows.Count;

                for (int spalte = 1; spalte < 9; spalte += 7)
                {
                    DateTime zeit = new DateTime();
                    string raum = "";
                    string bildungsgang = "";
                    var datum = Convert.ToString(xlRange.Cells[2, spalte].Value2);

                    for (int zeile = 4; zeile < rowCount + 1; zeile++)
                    {
                        var termin = new Termin();

                        bildungsgang = Convert.ToString(xlRange.Cells[zeile, spalte + 1].Value2) ?? bildungsgang;
                        termin.Bildungsgang = bildungsgang;
                        termin.Klasse = Convert.ToString(xlRange.Cells[zeile, spalte + 2].Value2);
                        zeit = GetZeit(zeit, datum, Convert.ToString(xlRange.Cells[zeile, spalte + 4].Value2));
                        termin.Uhrzeit = zeit;
                        raum = Convert.ToString(xlRange.Cells[zeile, spalte + 5].Value2) ?? raum;
                        termin.Raum = raum;
                        if (termin.Klasse != null)
                        {
                            termine.Add(termin);
                        }
                    }
                }
                
                workbook.Close(0);
                excel.Quit();
                Console.WriteLine(" ... ok.");
                return termine;

            }
            catch (Exception ex)
            {                
                Console.WriteLine(ex.ToString());
                workbook.Close(0);
                excel.Quit();
                Console.ReadKey();                
                return termine;
            }
        }

        private static DateTime GetZeit(DateTime uhrzeit, dynamic datum, dynamic dynamic)
        {
            try
            {
                return DateTime.FromOADate(double.Parse(datum) + double.Parse(dynamic));
            }
            catch (Exception)
            {
                return dynamic ?? uhrzeit.AddMinutes(10);
            }
        }

        private static void Settings()
        {   
            do
            {
                Console.Write(" 1. Wie heißt der Datenbankbenutzer? " + (Properties.Settings.Default.DBUser == "" ? "" : "[ " + Properties.Settings.Default.DBUser + " ]  "));
                var dbuser = Console.ReadLine();

                if (dbuser == "")
                {
                    _ = Properties.Settings.Default.DBUser;
                }
                else
                {
                    Properties.Settings.Default.DBUser = dbuser.ToLower();
                }

                try
                {
                    using (OdbcConnection connection = new OdbcConnection(ConnectionStringAtlantis + Properties.Settings.Default.DBUser))
                    {
                        connection.Open();
                        connection.Close();
                    }
                }
                catch (Exception)
                {
                    Properties.Settings.Default.DBUser = "";
                    Console.WriteLine("     Der Datenbankuser scheint nicht zu funktionieren. Flasche Eingabe? Ist Atlantis auf diesem Rechner nicht installiert?");
                }

                Properties.Settings.Default.Save();

            } while (Properties.Settings.Default.DBUser == "");
            Console.WriteLine("");
            do
            {
                Console.WriteLine(" 2. Teil- und Vollzeitklassen lassen sich in Atlantis über die Organisationsform (=Anlage) unterscheiden.");
                Console.WriteLine("    Typischerweise beginnt die Organisationsform der Teilzeitklassen in Atlantis mit 'A'.");
                Console.Write("    Wie lautet der Anfangsbuchstabe der Organisationsform Ihrer Teilzeitklassen? " + (Properties.Settings.Default.Klassenart == "" ? "" : "[ " + Properties.Settings.Default.Klassenart + " ]  "));
                var dbuser = Console.ReadLine();

                if (dbuser == "")
                {
                    dbuser = Properties.Settings.Default.Klassenart;
                }
                else
                {
                    Properties.Settings.Default.Klassenart = dbuser.Substring(0, 1).ToUpper();
                }

                Properties.Settings.Default.Save();

            } while (Properties.Settings.Default.Klassenart == "");

            Console.WriteLine("");

            do
            {
                Console.WriteLine(" 3. Bei 3,5-jährigen Teilzeit-Bildungsgängen müssen zum Halbjahr im 4.Jahrgang die alten Noten geholt werden.");
                Console.WriteLine("    Für alle anderen Teilzeitklassen werden die Noten am Ende des 3.Jahrgangs geholt.");                
                Console.Write("    Wie lauten die Anfangsbuchstaben oder Klassennamen der 3,5-Jährigen? " + (Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben == "" ? "" : "[ " + Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben + " ]  "));

                var aks = Console.ReadLine();
                List<string> aksl = new List<string>();

                if (aks == "")
                {
                    aks = Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben;
                }
                else
                {
                    foreach (var item in aks.Trim(' ').Split(','))
                    {
                        aksl.Add(item.ToUpper());
                    }

                    Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben = "";

                    foreach (var item in aksl)
                    {
                        Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben = Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben + item + ",";
                    }

                    Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben = Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben.TrimEnd(',');
                    Properties.Settings.Default.Save();
                    var ddd = Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben;
                }
            } while (Properties.Settings.Default.AbschlussklassenAnfangsbuchstaben == "");
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
