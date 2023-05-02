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
        public static bool Debug = false;
        public static string User = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1];
        public static string Zeitstempel = DateTime.Now.ToString("yyMMdd-HHmmss");        
        public static List<string> AktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };

        static void Main(string[] args)
        {   
            #if DEBUG
                Debug = true;
            #endif

            var width = Console.WindowWidth;
            var height = Console.WindowHeight;
            Console.SetWindowSize(width, height * 2);
            Global.SqlZeilen = new List<string>();
            Global.PadRight = 116;

            Global.WriteLine("*" + "---".PadRight(Global.PadRight, '-') + "--*");
            Global.WriteLine("| Webuntisnoten2Atlantis    |    Published under the terms of GPLv3    |    Stefan Bäumer   " + DateTime.Now.Year + "  |  Version 20230430  |");
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

                var hzJz = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

                string targetPath = SetTargetPath();
                string sourceAbsenceTimesTotal = CheckFile(User, "AbsenceTimesTotal");
                string sourceMarksPerLesson = CheckFile(User, "MarksPerLesson");

                var alleAtlantisLehrer = new Lehrers(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj);

                // Alle Webuntis-Leistungen, bereinigt um Leistungen mit leerem Fach, mit leerer Klasse, Dopplungen bei Fach, Lehrer & Gesamtnote 

                var alleWebuntisLeistungen = new Leistungen(sourceMarksPerLesson, alleAtlantisLehrer);

                if (BlaueBriefeErstellen())
                {
                    alleWebuntisLeistungen = alleWebuntisLeistungen.GetBlaueBriefeLeistungen();
                }

                var alleMöglicheKlassen = (from m in alleWebuntisLeistungen
                                           where m.Klasse != ""
                                           where m.Gesamtnote != "" // nur Klassen, für die schon Noten gegeben wurden
                                           select m.Klasse).Distinct().ToList();

                do
                {
                    Global.SqlZeilen = new List<string>();
                    
                    Console.WriteLine(Global.List2String(alleMöglicheKlassen,','));

                    var interessierendeKlassen = GetIntessierendeKlassen(alleMöglicheKlassen, AktSj);

                    var targetAbsenceTimesTotal = Path.Combine(targetPath, Zeitstempel + "_AbsenceTimesTotal_" + Global.List2String(interessierendeKlassen, '_') + "_" + User + ".CSV");
                    var targetMarksPerLesson = Path.Combine(targetPath, Zeitstempel + "_MarksPerLesson_" + Global.List2String(interessierendeKlassen, '_') + "_" + User + ".CSV");

                    RelevanteDatensätzeAusCsvFilternUndProtokollErstellen(sourceAbsenceTimesTotal, targetAbsenceTimesTotal, interessierendeKlassen);
                    RelevanteDatensätzeAusCsvFilternUndProtokollErstellen(sourceMarksPerLesson, targetMarksPerLesson, interessierendeKlassen);

                    Leistungen atlantisLeistungen = new Leistungen();

                    foreach (var interessierendeKlasse in interessierendeKlassen)
                    {
                        var interessierendeSchülerDieserKlasse = (from m in alleWebuntisLeistungen where m.Klasse == interessierendeKlasse select m.SchlüsselExtern).Distinct().ToList();

                        if (Debug)
                        {
                            interessierendeSchülerDieserKlasse = GetIntessierendeSchüler(AktSj, interessierendeSchülerDieserKlasse);
                        }

                        var interessierendeWebuntisLeistungen = alleWebuntisLeistungen.GetIntessierendeWebuntisLeistungen(interessierendeSchülerDieserKlasse, interessierendeKlasse);

                        interessierendeWebuntisLeistungen.LinkZumTeamsChatErzeugen(alleAtlantisLehrer);

                        atlantisLeistungen = new Leistungen(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj, User, interessierendeKlasse, interessierendeWebuntisLeistungen);

                        Leistungen geholteLeistungen = atlantisLeistungen.FilterNeuesteGeholteLeistungen(interessierendeSchülerDieserKlasse, interessierendeWebuntisLeistungen, interessierendeKlasse, AktSj, hzJz);
                        
                        atlantisLeistungen.TabelleErzeugen(interessierendeWebuntisLeistungen, geholteLeistungen, interessierendeKlasse, AktSj);

                        var webuntisAbwesenheiten = sourceAbsenceTimesTotal == null ? null : new Abwesenheiten(sourceAbsenceTimesTotal, interessierendeKlasse, interessierendeWebuntisLeistungen);

                        if (interessierendeWebuntisLeistungen.NotenblattNichtLeeren(atlantisLeistungen,webuntisAbwesenheiten,targetAbsenceTimesTotal))
                        {                            
                            interessierendeWebuntisLeistungen.AddRange(atlantisLeistungen.NotenVergangenerAbschnitteZiehen(interessierendeWebuntisLeistungen, geholteLeistungen, interessierendeKlasse, AktSj));

                            // Korrekturen durchführen

                            interessierendeWebuntisLeistungen = interessierendeWebuntisLeistungen.WidersprechendeGesamtnotenKorrigieren(atlantisLeistungen);
                            //interessierendeWebuntisLeistungen.ReligionsabwählerBehandeln(atlantisLeistungen);
                            interessierendeWebuntisLeistungen.BindestrichfächerZuordnen(atlantisLeistungen);
                            atlantisLeistungen.FehlendeZeugnisbemerkungBeiStrich(interessierendeWebuntisLeistungen, interessierendeKlasse);

                            interessierendeWebuntisLeistungen.AtlantisLeistungenZuordnenUndQueryBauen(atlantisLeistungen, AktSj[0] + "/" + AktSj[1], interessierendeKlasse);
                        }

                        // Add-Delete-Update

                        string hinweis = interessierendeWebuntisLeistungen.Update(atlantisLeistungen);

                        if (sourceAbsenceTimesTotal != null)
                        {
                            Abwesenheiten atlantisAbwesenheiten = new Abwesenheiten();
                            atlantisAbwesenheiten = targetAbsenceTimesTotal == null ? null : new Abwesenheiten(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj[0] + "/" + AktSj[1], interessierendeKlasse);
                            atlantisAbwesenheiten.Add(webuntisAbwesenheiten);
                            atlantisAbwesenheiten.Delete(webuntisAbwesenheiten);
                            atlantisAbwesenheiten.Update(webuntisAbwesenheiten);
                        }
                        else
                        {
                            int outputIndex = Global.SqlZeilen.Count();
                            Global.PrintMessage(outputIndex, ("Es werden keine Abwesenheiten importiert, da die Importdatei nicht von heute ist."));
                        }

                        if (hinweis != "")
                        {
                            Global.WriteLine("");
                            Global.WriteLine(hinweis);
                        }

                        Console.WriteLine("-".PadRight(Global.PadRight, '-') + "----");
                    }

                    string targetSql = Path.Combine(targetPath, Zeitstempel + "_webuntisnoten2atlantis_" + Zeichenkette(interessierendeKlassen) + "_" + User + ".SQL");
                    atlantisLeistungen.ErzeugeSqlDatei(new List<string>() { targetAbsenceTimesTotal, targetMarksPerLesson, targetSql });
                    OpenFiles(new List<string> { targetSql });

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

        private static Leistungen GetBlaueBriefeLeistungen()
        {
            throw new NotImplementedException();
        }

        private static bool BlaueBriefeErstellen()
        {
            ConsoleKeyInfo x;
            do
            {
                Console.WriteLine(" ");
                Global.WriteLine("Wollen Sie Blaue Briefe nach Atlantis übertragen?\n Wenn Sie j tippen, werden nur Leistungen der Prüfungsart 'Mahnung' berücksichtigt (j/N)");
                x = Console.ReadKey();
            } while (x.Key.ToString().ToLower() != "j" && x.Key.ToString().ToLower() != "n" && x.Key.ToString() != "Enter");

            Global.WriteLine("  Ihre Auswahl: " + (x.Key.ToString() == "Enter" || x.Key.ToString().ToUpper() == "N" ? "N" : "J"));
            Global.WriteLine(" ");

            if (x.Key.ToString().ToLower() == "j")
            {
                return true;
            }
            return false;
        }

        private static List<string> GetIntessierendeKlassen(List<string> möglicheKlassen, List<string> aktSj)
        {
            var interessierendeKlassen = new List<string>();

            try
            {
                do
                {
                    Console.Write("  Bitte die interessierenden Klassen kommasepariert angeben [" + GetVorschlag(möglicheKlassen) + "]: ");

                    var x = Console.ReadLine();

                    if (x == "")
                    {
                        interessierendeKlassen.AddRange(Properties.Settings.Default.InteressierendeKlassen.Split(','));
                        x = Properties.Settings.Default.InteressierendeKlassen;
                    }
                    else
                    {
                        var interessierendeKlassenString = "";

                        foreach (var klasse in möglicheKlassen)
                        {
                            foreach (var item in x.ToUpper().Replace(" ", "").Split(','))
                            {
                                if (klasse != "" && klasse.StartsWith(item.Replace(" ", "").ToUpper()))
                                {
                                    interessierendeKlassen.Add(klasse);
                                    interessierendeKlassenString += klasse + ",";
                                }
                            }
                        }

                        Properties.Settings.Default.InteressierendeKlassen = interessierendeKlassenString.TrimEnd(',');
                        Properties.Settings.Default.Save();
                    }
                } while (Properties.Settings.Default.InteressierendeKlassen == "");
            }
            catch (Exception ex)
            {
                Global.WriteLine("Bei der Auswahl der interessierenden Klasse ist es zum Fehler gekommen. \n " + ex);
            }
            Global.WriteLine("   Ihre Auswahl: " + Global.List2String(interessierendeKlassen, ','));
            Global.WriteLine(" ");

            return interessierendeKlassen;
        }

        private static List<int> GetIntessierendeSchüler(List<string> aktSj, List<int> möglicheSuS)
        {
            var interessierendeSuS = new List<int>();

            try
            {
                do
                {
                    Console.Write("  Bitte die ID der interessierenden SuS kommasepariert angeben oder 'alle' eingeben [" + GetVorschlag(möglicheSuS) + "]: ");

                    var x = Console.ReadLine();

                    if (x == "")
                    {
                        if (Properties.Settings.Default.InteressierendeSuS.ToUpper().StartsWith("A"))
                        {
                            foreach (var item in möglicheSuS)
                            {   
                                interessierendeSuS.Add(Convert.ToInt32(item));
                            }                            
                        }
                        else
                        {
                            foreach (var item in Properties.Settings.Default.InteressierendeSuS.Split(','))
                            {
                                if (item != null && item != "")
                                {
                                    interessierendeSuS.Add(Convert.ToInt32(item));
                                }
                            }
                        }
                                                
                        x = Properties.Settings.Default.InteressierendeSuS;
                    }
                    else
                    {
                        var interessierendeSuSstring = "";

                        foreach (var sch in möglicheSuS)
                        {
                            if (x.ToUpper().StartsWith("A"))
                            {      
                                interessierendeSuS.Add(sch);                                
                                interessierendeSuSstring = "alle";
                            }
                            else
                            {
                                foreach (var item in x.ToUpper().Replace(" ", "").Split(','))
                                {
                                    if (sch.ToString() == item)
                                    {
                                        interessierendeSuS.Add(sch);
                                        interessierendeSuSstring += sch + ",";
                                    }
                                }
                            }
                        }

                        Properties.Settings.Default.InteressierendeSuS = interessierendeSuSstring.TrimEnd(',');
                        Properties.Settings.Default.Save();
                    }
                } while (Properties.Settings.Default.InteressierendeSuS == "");
            }
            catch (Exception ex)
            {
                Global.WriteLine("Bei der Auswahl der interessierenden Klasse ist es zum Fehler gekommen. \n " + ex);
            }
            if (Properties.Settings.Default.InteressierendeSuS.ToLower().StartsWith("a"))
            {
                Global.WriteLine("   Ihre Auswahl: alle");
            }
            else
            {
                Global.WriteLine("   Ihre Auswahl: " + Global.List2String(interessierendeSuS, ','));
            }
            
            Global.WriteLine(" ");

            return interessierendeSuS;
        }

        private static string GetVorschlag(List<string> möglicheKlassen)
        {
            string vorschlag = "";

            // Wenn in den Properties ein Eintrag existiert, ...

            if (Properties.Settings.Default.InteressierendeKlassen != null && Properties.Settings.Default.InteressierendeKlassen != "")
            {
                // ... wird für alle kommaseparierten Einträge geprüft ...

                foreach (var item in Properties.Settings.Default.InteressierendeKlassen.Split(','))
                {
                    // ... ob es in den möglichen Klassen Kandidaten gibt.

                    if ((from t in möglicheKlassen where t.StartsWith(item) select t).Any())
                    {
                        // Falls ja, dann wird der Vorschlag aus den Properties übernommen.

                        vorschlag += item + ",";
                    }
                }
            }
            Properties.Settings.Default.InteressierendeKlassen = vorschlag;
            Properties.Settings.Default.Save();
            return vorschlag.TrimEnd(',');
        }

        private static string GetVorschlag(List<int> möglicheSuS)
        {
            string vorschlag = "";

            // Wenn in den Properties ein Eintrag existiert, ...

            if (Properties.Settings.Default.InteressierendeSuS != null && Properties.Settings.Default.InteressierendeSuS != "")
            {
                // ... wird für alle kommaseparierten Einträge geprüft ...

                foreach (var item in Properties.Settings.Default.InteressierendeSuS.Split(','))
                {
                    // ... ob es in den möglichen Klassen Kandidaten gibt.

                    if ((from t in möglicheSuS select t).Any())
                    {
                        // Falls ja, dann wird der Vorschlag aus den Properties übernommen.

                        vorschlag += item + ",";
                    }
                }
            }
            Properties.Settings.Default.InteressierendeSuS = vorschlag;
            Properties.Settings.Default.Save();
            return vorschlag.TrimEnd(',');
        }

        private static string MöglicheKlassenToString(List<string> möglicheKlassen)
        {
            var möglicheKlassenString = "\nMögliche Klassen aus der Webuntis-Datei: ";
            int i = 0;

            foreach (var ik in möglicheKlassen)
            {
                if (ik.Length > 3)
                {
                    if (i % 17 == 0)
                    {
                        möglicheKlassenString += "\n ";
                    }
                    möglicheKlassenString += ik + " ";
                }
                i++;
            }
            return möglicheKlassenString.TrimEnd(' ');
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

        private static void RelevanteDatensätzeAusCsvFilternUndProtokollErstellen(string sourceFile, string targetfile, List<string> interessierendeKlassen)
        {
            try
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
            }
            catch (Exception)
            {

            }
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