// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Linq;

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

        static void Main(string[] args)
        {
            var width = Console.WindowWidth;
            var height = Console.WindowHeight;
            Console.SetWindowSize(width, height * 2);
            Global.SqlZeilen = new List<string>();
            Global.PadRight = 116;

            Global.WriteLine("*" + "---".PadRight(Global.PadRight, '-') + "--*");
            Global.WriteLine("| Webuntisnoten2Atlantis    |    Published under the terms of GPLv3    |    Stefan Bäumer   " + DateTime.Now.Year + "  |  Version 20230330  |");
            Global.WriteLine("|" + "---".PadRight(Global.PadRight, '-') + "--|");
            Global.WriteLine("| Webuntisnoten2Atlantis erstellt eine SQL-Datei mit Befehlen zum Import der Noten/Punkte aus Webuntis nach Atlantis   |");
            Global.WriteLine("| ACHTUNG:  Wenn es die Lehrkraft versäumt hat die Teilleistung zu dokumentieren, wird keine Gesamtnote von Webuntis   |");
            Global.WriteLine("| nach Atlantis übergeben!                                                                                             |");
            Global.WriteLine("*" + "---".PadRight(Global.PadRight, '-') + "--*");

            try
            {
                if (Properties.Settings.Default.DBUser == "" || Properties.Settings.Default.Klassenart == null || Properties.Settings.Default.Klassenart == "")
                {
                    Settings();
                }

                Global.HzJz = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

                string targetPath = SetTargetPath();
                string sourceAbsenceTimesTotal = CheckFile(User, "AbsenceTimesTotal");
                string sourceMarksPerLesson = CheckFile(User, "MarksPerLesson");

                Lehrers alleAtlantisLehrer = new Lehrers(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj);

                Leistungen möglicheKlassen = new Leistungen(sourceMarksPerLesson, alleAtlantisLehrer, new List<string>());
                var möglicheKlassenString = möglicheKlassen.GetMöglicheKlassen();

                do
                {
                    Global.SqlZeilen = new List<string>();
                    var interessierendeKlassen = new List<string>();
                    Console.WriteLine(möglicheKlassenString);
                    interessierendeKlassen.AddRange(möglicheKlassen.GetIntessierendeKlassen(AktSj));

                    var targetAbsenceTimesTotal = Path.Combine(targetPath, Zeitstempel + "_AbsenceTimesTotal_" + Global.List2String(interessierendeKlassen,'_') + "_" + User + ".CSV");
                    var targetMarksPerLesson = Path.Combine(targetPath, Zeitstempel + "_MarksPerLesson_" + Global.List2String(interessierendeKlassen,'_') + "_" + User + ".CSV");

                    Leistungen webuntisLeistungen = new Leistungen(sourceMarksPerLesson, alleAtlantisLehrer, interessierendeKlassen);

                    // Die eingelesenen Dateien für Protokollzwecke filtern

                    RelevanteDatensätzeAusCsvFiltern(sourceAbsenceTimesTotal, targetAbsenceTimesTotal, interessierendeKlassen);
                    RelevanteDatensätzeAusCsvFiltern(sourceMarksPerLesson, targetMarksPerLesson, interessierendeKlassen);

                    Leistungen atlantisLeistungen = new Leistungen(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj, User, interessierendeKlassen, webuntisLeistungen);
                    Abwesenheiten atlantisAbwesenheiten = targetAbsenceTimesTotal == null ? null : new Abwesenheiten(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj[0] + "/" + AktSj[1], interessierendeKlassen);
                    Global.WebuntisAbwesenheiten = targetAbsenceTimesTotal == null ? null : new Abwesenheiten(sourceAbsenceTimesTotal, interessierendeKlassen, webuntisLeistungen);

                    if (webuntisLeistungen.NotenblattNichtLeeren(atlantisLeistungen))
                    {                                             
                        // Noten vergangener Abschnitte ziehen

                        webuntisLeistungen.AddRange(atlantisLeistungen.NotenVergangenerAbschnitteZiehen(webuntisLeistungen, interessierendeKlassen, AktSj));

                        // Korrekturen durchführen

                        webuntisLeistungen = webuntisLeistungen.WidersprechendeGesamtnotenKorrigieren(atlantisLeistungen);
                        webuntisLeistungen.ReligionsabwählerBehandeln(atlantisLeistungen);
                        webuntisLeistungen.BindestrichfächerZuordnen(atlantisLeistungen);
                        atlantisLeistungen.FehlendeZeugnisbemerkungBeiStrich(webuntisLeistungen, interessierendeKlassen);

                        webuntisLeistungen.AtlantisLeistungenZuordnenUndQueryBauen(atlantisLeistungen, AktSj[0] + "/" + AktSj[1], interessierendeKlassen);

                    }

                    // Add-Delete-Update

                    string hinweis = webuntisLeistungen.Update(atlantisLeistungen);

                    if (targetAbsenceTimesTotal != null)
                    {
                        atlantisAbwesenheiten.Add(Global.WebuntisAbwesenheiten);
                        atlantisAbwesenheiten.Delete(Global.WebuntisAbwesenheiten);
                        atlantisAbwesenheiten.Update(Global.WebuntisAbwesenheiten);
                    }
                    else
                    {
                        int outputIndex = Global.SqlZeilen.Count();
                        Global.PrintMessage(outputIndex, ("Es werden keine Abwesenheiten importiert, da die Importdatei nicht von heute ist."));
                    }

                    string targetSql = Path.Combine(targetPath, Zeitstempel + "_webuntisnoten2atlantis_" + Zeichenkette(interessierendeKlassen) + "_" + User + ".SQL");

                    OpenFiles(new List<string> { targetSql });

                    if (hinweis != "")
                    {
                        Global.WriteLine("");
                        Global.WriteLine(hinweis);
                    }

                    Console.WriteLine("-".PadRight(Global.PadRight, '-') + "----");

                    atlantisLeistungen.ErzeugeSqlDatei(new List<string>() { targetAbsenceTimesTotal, targetMarksPerLesson, targetSql });

                } while (true);
            }
            catch (Exception ex)
            {
                Global.WriteLine("Ooops! Es ist etwas schiefgelaufen! Die Verarbeitung wird gestoppt.");
                Global.WriteLine("");
                Global.WriteLine(ex.ToString());
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private static string Zeichenkette(List<string> interessierendeKlassen)
        {
            var x = "";
            foreach (var item in interessierendeKlassen)
            {
                x += item;
            }
            return x;
        }

        private static void OpenFiles(List<string> files)
        {
            try
            {
                Process notepadPlus = new Process();
                notepadPlus.StartInfo.FileName = "notepad++.exe";

                for (int i = 0; i < files.Count; i++)
                {
                    notepadPlus.StartInfo.Arguments = files[i];
                    notepadPlus.StartInfo.Arguments = @"-multiInst -nosession " + files[i];
                    notepadPlus.Start();
                }
            }
            catch (Exception)
            {
                System.Diagnostics.Process.Start("Notepad.exe", files[0]);
            }
        }

        private static void RelevanteDatensätzeAusCsvFiltern(string sourceFile, string targetfile, List<string> interessierendeKlassen)
        {
            int anzahlZeilen = 0;

            using (var sr = new StreamReader(sourceFile))
            using (var sw = new StreamWriter(targetfile))
            {
                string line;
                int zeile = 0;

                while ((line = sr.ReadLine()) != null)
                {
                    bool relevanteZeile = false;

                    foreach (var iK in interessierendeKlassen)
                    {
                        if (line.Contains(iK))
                        {
                            relevanteZeile = true;
                            anzahlZeilen++;
                        }
                    }

                    if (relevanteZeile || zeile == 0)
                    {
                        sw.WriteLine(line);
                    }
                    zeile++;
                }
            }

            //Global.AufConsoleSchreiben(("Datensätze der Klasse " + interessierendeKlassen[0] + " in " + Path.GetFileName(targetfile)).PadRight(Global.PadRight, '.') + anzahlZeilen.ToString().PadLeft(4));
        }

        private static string CheckFile(string user, string kriterium)
        {
            var sourceFile = (from f in Directory.GetFiles(@"c:\users\" + user + @"\Downloads", "*.csv", SearchOption.AllDirectories) where f.Contains(kriterium) orderby File.GetLastWriteTime(f) select f).LastOrDefault();

            if ((sourceFile == null || System.IO.File.GetLastWriteTime(sourceFile).Date != DateTime.Now.Date))
            {
                Global.WriteLine("");
                Global.WriteLine(" Die " + kriterium + "<...>.csv" + (sourceFile == null ? " existiert nicht im Download-Ordner" : " im Download-Ordner ist nicht von heute. \n Es werden keine Daten aus der Datei importiert") + ".");
                Global.WriteLine(" Exportieren Sie die Datei frisch aus Webuntis, indem Sie als Administrator:");

                if (kriterium.Contains("MarksPerLesson"))
                {
                    Global.WriteLine("   1. Klassenbuch > Berichte klicken");
                    Global.WriteLine("   2. Alle Klassen auswählen und ggfs. den Zeitraum einschränken");
                    Global.WriteLine("   3. Unter \"Noten\" die Prüfungsart (-Alle-) auswählen");
                    Global.WriteLine("   4. Unter \"Noten\" den Haken bei Notennamen ausgeben _NICHT_ setzen");
                    Global.WriteLine("   5. Hinter \"Noten pro Schüler\" auf CSV klicken");
                    Global.WriteLine("   6. Die Datei \"MarksPerLesson<...>.CSV\" im Download-Ordner zu speichern");
                    Global.WriteLine(" ");
                    Global.WriteLine(" ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                if (kriterium.Contains("AbsenceTimesTotal"))
                {
                    Global.WriteLine("   1. Administration > Export klicken");
                    Global.WriteLine("   2. Zeitraum begrenzen, also die Woche der Zeugniskonferenz und vergange Abschnitte herauslassen");
                    Global.WriteLine("   2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
                    Global.WriteLine("   4. Die Datei \"AbsenceTimesTotal<...>.CSV\" im Download-Ordner zu speichern");
                }
                Global.WriteLine(" ");
                sourceFile = null;
            }

            if (sourceFile != null)
            {
                Global.WriteLine("Ausgewertete Datei: " + (Path.GetFileName(sourceFile) + " ").PadRight(53, '.') + ". Erstell-/Bearbeitungszeitpunkt heute um " + System.IO.File.GetLastWriteTime(sourceFile).ToShortTimeString());
            }

            return sourceFile;
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
                                Global.WriteLine("");
                                Global.WriteLine("Empty value not allowed.");
                                CheckPassword(EnterText);
                                break;
                            }
                            else
                            {
                                Global.WriteLine("");
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
                    Global.WriteLine("     Der Datenbankuser scheint nicht zu funktionieren. Flasche Eingabe? Ist Atlantis auf diesem Rechner nicht installiert?");
                }

                Properties.Settings.Default.Save();

            } while (Properties.Settings.Default.DBUser == "");

            Global.WriteLine("");

            do
            {
                Global.WriteLine(" 2. Teil- und Vollzeitklassen lassen sich in Atlantis über die Organisationsform (=Anlage) unterscheiden.");
                Global.WriteLine("    Typischerweise beginnt die Organisationsform der Teilzeitklassen in Atlantis mit 'A'.");
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

            Global.WriteLine("");

            do
            {
                Global.WriteLine(" 3. Bei 3,5-jährigen Teilzeit-Bildungsgängen müssen zum Halbjahr im 4.Jahrgang die alten Noten geholt werden.");
                Global.WriteLine("    Für alle anderen Teilzeitklassen werden die Noten am Ende des 3.Jahrgangs geholt.");
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
                    Global.WriteLine(" Wo sollen die Dateien gespeichert werden? [ " + pfad + " ]");
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

                        Global.WriteLine("Der Pfad " + pfad + " kann nicht angelegt werden.");
                    }

                } while (!Directory.Exists(pfad));
            }
            
            Global.WriteLine(("|     Pfad zu den Dateien: " + pfad).PadRight(Global.PadRight,' ') + "   |");
            Global.WriteLine("*" + "---".PadRight(Global.PadRight, '-') + "--*");
            Global.WriteLine(@" ");

            return pfad;
        }
    }
}