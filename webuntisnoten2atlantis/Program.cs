// Published under the terms of GPLv3 Stefan Bäumer 2020.

using Microsoft.Exchange.WebServices.Data;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
                Console.WriteLine(" Webuntisnoten2Atlantis | Published under the terms of GPLv3 | Stefan Bäumer " + DateTime.Now.Year + " | Version 20210507");
                Console.WriteLine("=====================================================================================================");
                Console.WriteLine(" *Webuntisnoten2Atlantis* erstellt eine SQL-Datei mit entsprechenden Befehlen zum Import in Atlantis.");
                Console.WriteLine(" ACHTUNG: Wenn der Lehrer es versäumt hat, mindestens 1 Teilleistung zu dokumentieren, wird keine Ge-");
                Console.WriteLine(" samtnote von Webuntis übergeben!");
                Console.WriteLine("=====================================================================================================");

                if (Properties.Settings.Default.DBUser == "")
                {
                    Settings();
                }

                string targetPath = SetTargetPath();
                Process notepadPlus = new Process();
                notepadPlus.StartInfo.FileName = "notepad++.exe";
                notepadPlus.Start();
                Thread.Sleep(1500);
                string targetAbsenceTimesTotal = CheckFile(targetPath, User, "AbsenceTimesTotal");
                string targetMarksPerLesson = CheckFile(targetPath, User, "MarksPerLesson");
                string targetSql = Path.Combine(targetPath, Zeitstempel + "_webuntisnoten2atlantis_" + User + ".SQL");

                Leistungen alleAtlantisLeistungen = new Leistungen(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj, User);
                Leistungen alleWebuntisLeistungen = new Leistungen(targetMarksPerLesson);

                Abwesenheiten alleAtlantisAbwesenheiten = new Abwesenheiten(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj[0] + "/" + AktSj[1]);
                Abwesenheiten alleWebuntisAbwesenheiten = new Abwesenheiten(targetAbsenceTimesTotal);

                Leistungen webuntisLeistungen = new Leistungen();
                Abwesenheiten webuntisAbwesenheiten = new Abwesenheiten();
                Leistungen atlantisLeistungen = new Leistungen();
                Abwesenheiten atlantisAbwesenheiten = new Abwesenheiten();

                FocusMe();

                var interessierendeKlassen = new List<string>();
                interessierendeKlassen = alleAtlantisLeistungen.GetIntessierendeKlassen(alleWebuntisLeistungen, AktSj);

                webuntisLeistungen.AddRange((from a in alleWebuntisLeistungen where interessierendeKlassen.Contains(a.Klasse) select a).OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name));
                atlantisLeistungen.AddRange((from a in alleAtlantisLeistungen where interessierendeKlassen.Contains(a.Klasse) select a).OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name));
                
                List<string> abschlussklassen = (from t in alleAtlantisLeistungen where t.Abschlussklasse where interessierendeKlassen.Contains(t.Klasse) select t.Klasse).Distinct().ToList();

                atlantisAbwesenheiten.AddRange((from a in alleAtlantisAbwesenheiten where interessierendeKlassen.Contains(a.Klasse) where !abschlussklassen.Contains(a.Klasse) select a));
                webuntisAbwesenheiten.AddRange((from a in alleWebuntisAbwesenheiten where interessierendeKlassen.Contains(a.Klasse) where !abschlussklassen.Contains(a.Klasse) select a));
                
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

                fehlendezuordnungen.ManuellZuordnen(webuntisLeistungen, atlantisLeistungen);

                // Add-Delete-Update

                atlantisLeistungen.Add(webuntisLeistungen, interessierendeKlassen);
                atlantisLeistungen.Delete(webuntisLeistungen, interessierendeKlassen, AktSj);
                atlantisLeistungen.Update(webuntisLeistungen, interessierendeKlassen);

                atlantisAbwesenheiten.Add(webuntisAbwesenheiten);
                atlantisAbwesenheiten.Delete(webuntisAbwesenheiten);
                atlantisAbwesenheiten.Update(webuntisAbwesenheiten);

                alleAtlantisLeistungen.ErzeugeSqlDatei(new List<string>() { targetAbsenceTimesTotal, targetMarksPerLesson, targetSql });

                Console.WriteLine("");
                Console.WriteLine("  -----------------------------------------------------------------");
                Console.WriteLine("  Verarbeitung abgeschlossen. Programm beenden mit Enter.");
                Console.ReadKey();
                Environment.Exit(0);

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

            if ((sourceFile == null || System.IO.File.GetLastWriteTime(sourceFile).Date != DateTime.Now.Date))
            {
                Console.WriteLine("");
                Console.WriteLine("  Die Datei " + kriterium + "<...>.CSV" + (sourceFile == null ? " existiert nicht im Download-Ordner" : " im Download-Ordner ist nicht von heute") + ".");
                Console.WriteLine("  Exportieren Sie die Datei frisch aus Webuntis, indem Sie als Administrator:");

                if (kriterium.Contains("MarksPerLesson"))
                {
                    Console.WriteLine("   1. Klassenbuch > Berichte klicken");
                    Console.WriteLine("   2. Alle Klassen auswählen und ggfs. den Zeitraum einschränken");
                    Console.WriteLine("   3. Unter \"Noten\" die Prüfungsart (-Alle-) auswählen");
                    Console.WriteLine("   4. Unter \"Noten\" den Haken bei Notennamen ausgeben _NICHT_ setzen");
                    Console.WriteLine("   5. Hinter \"Noten pro Schüler\" auf CSV klicken");
                    Console.WriteLine("   6. Die Datei \"MarksPerLesson<...>.CSV\" im Download-Ordner zu speichern");
                }

                if (kriterium.Contains("AbsenceTimesTotal"))
                {
                    Console.WriteLine("   1. Administration > Export klicken");
                    Console.WriteLine("   2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
                    Console.WriteLine("   3. !!! Zeitraum begrenzen (also die Woche der) Zeugniskonferenz herauslassen !!!");
                    Console.WriteLine("   4. Die Datei \"AbsenceTimesTotal<...>.CSV\" im Download-Ordner zu speichern");
                }

                Console.WriteLine("");
                Console.WriteLine(" ENTER beendet das Programm.");
                Console.ReadKey();
                Environment.Exit(0);

            }
            var targetFile = Path.Combine(targetPath, Zeitstempel + "_" + kriterium + "_" + user + ".CSV");

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
            return Path.Combine(targetFile);
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

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(HandleRef hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr zeroOnly, string lpWindowName);
        public const int SW_RESTORE = 9;
        static void FocusMe()
        {
            string originalTitle = Console.Title;
            string uniqueTitle = Guid.NewGuid().ToString();
            Console.Title = uniqueTitle;
            Thread.Sleep(50);
            IntPtr handle = FindWindowByCaption(IntPtr.Zero, uniqueTitle);

            Console.Title = originalTitle;

            ShowWindowAsync(new HandleRef(null, handle), SW_RESTORE);
            SetForegroundWindow(handle);
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
            //do
            //{
            //    Console.WriteLine(" 2. Teil- und Vollzeitklassen lassen sich in Atlantis über die Organisationsform (=Anlage) unterscheiden.");
            //    Console.WriteLine("    Typischerweise beginnt die Organisationsform der Teilzeitklassen in Atlantis mit 'A'.");
            //    Console.Write("    Wie lautet der Anfangsbuchstabe der Organisationsform Ihrer Teilzeitklassen? " + (Properties.Settings.Default.AnfangsbuchstabeTkKlassen == "" ? "" : "[ " + Properties.Settings.Default.AnfangsbuchstabeTkKlassen + " ]  "));
            //    var dbuser = Console.ReadLine();

            //    if (dbuser == "")
            //    {
            //        dbuser = Properties.Settings.Default.AnfangsbuchstabeTkKlassen;
            //    }
            //    else
            //    {
            //        Properties.Settings.Default.AnfangsbuchstabeTkKlassen = dbuser.Substring(0, 1).ToUpper();
            //    }

            //    Properties.Settings.Default.Save();

            //} while (Properties.Settings.Default.AnfangsbuchstabeTkKlassen == "");

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
