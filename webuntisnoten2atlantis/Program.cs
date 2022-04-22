// Published under the terms of GPLv3 Stefan Bäumer 2021.

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace webuntisnoten2atlantis
{
    class Program
    {
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=";
        public static string Passwort = "";
        public static string User = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1];
        public static string Zeitstempel = DateTime.Now.ToString("yyMMdd-HHmmss");
        public static List<string> AktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };
        public static string ExcelPath = @"C:\Users\" + User + @"\Berufskolleg Borken\Kollegium - General\03 Schulleitung\3.04 Termine\2020-21\14 Zeugniskonferenzen HZ\\Zeugniskonferenzen HZ.xlsx";

        static void Main(string[] args)
        {
            Global.Output = new List<string>();

            try
            {
                Console.WriteLine(" Webuntisnoten2Atlantis | Published under the terms of GPLv3 | Stefan Bäumer " + DateTime.Now.Year + " | Version 20220419");
                Console.WriteLine("=====================================================================================================");
                Console.WriteLine(" *Webuntisnoten2Atlantis* erstellt eine SQL-Datei mit entsprechenden Befehlen zum Import in Atlantis.");
                Console.WriteLine(" ACHTUNG: Wenn der Lehrer es versäumt hat, mindestens 1 Teilleistung zu dokumentieren, wird keine Ge-");
                Console.WriteLine(" samtnote von Webuntis nach Atlantis übergeben!");
                Console.WriteLine("=====================================================================================================");

                if (Properties.Settings.Default.DBUser == "" || Properties.Settings.Default.Klassenart == null || Properties.Settings.Default.Klassenart == "")
                {
                    Settings();
                }

                Global.HzJz = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

                string targetPath = SetTargetPath();
                Process notepadPlus = new Process();
                notepadPlus.StartInfo.FileName = "notepad++.exe";
                notepadPlus.Start();
                Thread.Sleep(1500);
                string targetAbsenceTimesTotal = CheckFile(targetPath, User, "AbsenceTimesTotal");
                string targetMarksPerLesson = CheckFile(targetPath, User, "MarksPerLesson");
                string targetSql = Path.Combine(targetPath, Zeitstempel + "_webuntisnoten2atlantis_" + User + ".SQL");

                Lehrers alleAtlantisLehrer = new Lehrers(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj);
                Leistungen alleAtlantisLeistungen = new Leistungen(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj, User);
                Leistungen alleWebuntisLeistungen = new Leistungen(targetMarksPerLesson, alleAtlantisLehrer);

                Abwesenheiten alleAtlantisAbwesenheiten = targetAbsenceTimesTotal == null ? null : new Abwesenheiten(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj[0] + "/" + AktSj[1]);
                Abwesenheiten alleWebuntisAbwesenheiten = targetAbsenceTimesTotal == null ? null : new Abwesenheiten(targetAbsenceTimesTotal);

                Leistungen webuntisLeistungen = new Leistungen();
                Abwesenheiten webuntisAbwesenheiten = new Abwesenheiten();
                Leistungen atlantisLeistungen = new Leistungen();
                Abwesenheiten atlantisAbwesenheiten = new Abwesenheiten();

                //do
                //{
                var interessierendeKlassen = new List<string>();
                interessierendeKlassen = alleAtlantisLeistungen.GetIntessierendeKlassen(alleWebuntisLeistungen, AktSj);

                webuntisLeistungen.AddRange((from a in alleWebuntisLeistungen where interessierendeKlassen.Contains(a.Klasse) select a).OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name));

                if (targetAbsenceTimesTotal != null)
                {
                    webuntisAbwesenheiten.AddRange((from a in alleWebuntisAbwesenheiten where interessierendeKlassen.Contains(a.Klasse) select a));
                    atlantisAbwesenheiten.AddRange((from a in alleAtlantisAbwesenheiten where interessierendeKlassen.Contains(a.Klasse) select a));
                }

                atlantisLeistungen.AddRange((from a in alleAtlantisLeistungen where interessierendeKlassen.Contains(a.Klasse) select a).OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name));

                atlantisLeistungen.ErzeugeSerienbriefquelleFehlendeTeilleistungen(webuntisLeistungen);

                if (webuntisLeistungen.Count == 0)
                {
                    throw new Exception("[!] Es liegt kein einziger Leistungsdatensatz für Ihre Auswahl vor. Ist evtl. die Auswahl in Webuntis eingeschränkt? ");
                }

                // Alte Noten holen

                webuntisLeistungen.AddRange(alleAtlantisLeistungen.HoleAlteNoten(webuntisLeistungen, interessierendeKlassen, AktSj));

                // Sortieren

                webuntisLeistungen.OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name);
                atlantisLeistungen.OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name);

                // Korrekturen durchführen

                webuntisLeistungen.WidersprechendeGesamtnotenKorrigieren(interessierendeKlassen);

                Zuordnungen fehlendezuordnungen = webuntisLeistungen.FächerZuordnen(atlantisLeistungen);
                webuntisLeistungen.ReligionsabwählerBehandeln(atlantisLeistungen);
                webuntisLeistungen.BindestrichfächerZuordnen(atlantisLeistungen);
                atlantisLeistungen.FehlendeZeugnisbemerkungBeiStrich(webuntisLeistungen, interessierendeKlassen);
                atlantisLeistungen.GetKlassenMitFehlendenZeugnisnoten(interessierendeKlassen, alleWebuntisLeistungen);
                //atlantisLeistungen.Gym12NotenInDasGostNotenblattKopieren(interessierendeKlassen, AktSj);

                fehlendezuordnungen.ManuellZuordnen(webuntisLeistungen, atlantisLeistungen);

                // Add-Delete-Update

                atlantisLeistungen.Add(webuntisLeistungen, interessierendeKlassen);
                atlantisLeistungen.Delete(webuntisLeistungen, interessierendeKlassen, AktSj);
                atlantisLeistungen.Update(webuntisLeistungen, interessierendeKlassen);

                if (targetAbsenceTimesTotal != null)
                {
                    atlantisAbwesenheiten.Add(webuntisAbwesenheiten);
                    atlantisAbwesenheiten.Delete(webuntisAbwesenheiten);
                    atlantisAbwesenheiten.Update(webuntisAbwesenheiten);
                }
                else
                {
                    int outputIndex = Global.Output.Count();
                    Global.PrintMessage(outputIndex, ("Es werden keine Abwesenheiten importiert, da die Importdatei nicht von heute ist."));
                }

                alleAtlantisLeistungen.ErzeugeSqlDatei(new List<string>() { targetAbsenceTimesTotal, targetMarksPerLesson, targetSql });

                //Global.Defizitleistungen.ErzeugeSerienbriefquelleNichtversetzer();

                Console.WriteLine("");
                Console.WriteLine("  -----------------------------------------------------------------");
                Console.WriteLine("  Verarbeitung abgeschlossen. Programm beenden mit Enter.");

                //    if ((char)13 == (Console.ReadKey()).KeyChar)
                //    {
                //        Environment.Exit(0);
                //    }
                //} while (true);
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

        private static string CheckFile(string targetPath, string user, string kriterium)
        {
            var sourceFile = (from f in Directory.GetFiles(@"c:\users\" + user + @"\Downloads", "*.csv", SearchOption.AllDirectories) where f.Contains(kriterium) orderby File.GetCreationTime(f) select f).LastOrDefault();

            var targetFile = Path.Combine(targetPath, Zeitstempel + "_" + kriterium + "_" + user + ".CSV");

            if ((sourceFile == null || System.IO.File.GetLastWriteTime(sourceFile).Date != DateTime.Now.Date))
            {
                Console.WriteLine("");
                Console.WriteLine("  Die Datei " + kriterium + "<...>.CSV" + (sourceFile == null ? " existiert nicht im Download-Ordner" : " im Download-Ordner ist nicht von heute. Es werden keine Daten aus der Datei importiert.") + ".");
                Console.WriteLine("  Exportieren Sie die Datei frisch aus Webuntis, indem Sie als Administrator:");

                if (kriterium.Contains("MarksPerLesson"))
                {
                    Console.WriteLine("   1. Klassenbuch > Berichte klicken");
                    Console.WriteLine("   2. Alle Klassen auswählen und ggfs. den Zeitraum einschränken");
                    Console.WriteLine("   3. Unter \"Noten\" die Prüfungsart (-Alle-) auswählen");
                    Console.WriteLine("   4. Unter \"Noten\" den Haken bei Notennamen ausgeben _NICHT_ setzen");
                    Console.WriteLine("   5. Hinter \"Noten pro Schüler\" auf CSV klicken");
                    Console.WriteLine("   6. Die Datei \"MarksPerLesson<...>.CSV\" im Download-Ordner zu speichern");
                    Console.WriteLine(" ");
                    Console.WriteLine(" ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                if (kriterium.Contains("AbsenceTimesTotal"))
                {
                    Console.WriteLine("   1. Administration > Export klicken");
                    Console.WriteLine("   2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
                    Console.WriteLine("   3. !!! Zeitraum begrenzen (also die Woche der) Zeugniskonferenz herauslassen !!!");
                    Console.WriteLine("   4. Die Datei \"AbsenceTimesTotal<...>.CSV\" im Download-Ordner zu speichern");
                }

                targetFile = null;
            }
            else
            {
                if (File.Exists(targetFile))
                {
                    File.Delete(targetFile);
                }
                File.Copy(sourceFile, targetFile);

                Process notepadPlus = new Process();
                notepadPlus.StartInfo.FileName = "notepad++.exe";
                //notepadPlus.StartInfo.Arguments = @"-multiInst -nosession " + targetFile;
                notepadPlus.StartInfo.Arguments = targetFile;
                notepadPlus.Start();
                Thread.Sleep(1500);
            }

            return targetFile;
        }

        static void CheckPassword(string EnterText)
        {
            try
            {
                Console.Write(EnterText);
                Passwort = "";
                do
                {
                    ConsoleKeyInfo key = Console.ReadKey(true);
                    // Backspace Should Not Work  
                    if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                    {
                        Passwort += key.KeyChar;
                        Console.Write("*");
                    }
                    else
                    {
                        if (key.Key == ConsoleKey.Backspace && Passwort.Length > 0)
                        {
                            Passwort = Passwort.Substring(0, (Passwort.Length - 1));
                            Console.Write("\b \b");
                        }
                        else if (key.Key == ConsoleKey.Enter)
                        {
                            if (string.IsNullOrWhiteSpace(Passwort))
                            {
                                Console.WriteLine("");
                                Console.WriteLine("Empty value not allowed.");
                                CheckPassword(EnterText);
                                break;
                            }
                            else
                            {
                                Console.WriteLine("");
                                break;
                            }
                        }
                    }
                } while (true);
            }
            catch (Exception ex)
            {
                throw ex;
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

                for (int spalte = 1; spalte < 8; spalte += 6)
                {
                    DateTime zeit = new DateTime();
                    string raum = "";
                    string bildungsgang = "";

                    var datum = Convert.ToString(xlRange.Cells[2, spalte].Value2);

                    for (int zeile = 3; zeile < rowCount + 1; zeile++)
                    {
                        var termin = new Termin();

                        bildungsgang = Convert.ToString(xlRange.Cells[zeile, spalte + 1].Value2) ?? bildungsgang;
                        termin.Bildungsgang = bildungsgang;
                        termin.Klasse = Convert.ToString(xlRange.Cells[zeile, spalte + 2].Value2);
                        zeit = GetZeit(zeit, datum, Convert.ToString(xlRange.Cells[zeile, spalte + 3].Value2));
                        termin.Uhrzeit = zeit;
                        raum = Convert.ToString(xlRange.Cells[zeile, spalte + 4].Value2) ?? raum;
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

        private static string SetTargetPath()
        {
            var pfad = @"\\fs01\SoftwarehausHeider\webuntisNoten2Atlantis\Dateien";

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
                        pfad = @"\\fs01\SoftwarehausHeider\webuntisNoten2Atlantis\Dateien";
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

            return pfad;
        }
    }
}
