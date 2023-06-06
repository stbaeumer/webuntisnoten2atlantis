// Published under the terms of GPLv3 Stefan Bäumer 2023.

using iTextSharp.text;
using iTextSharp.text.pdf;
using PdfSharp.Pdf.Security;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Linq;
//using System.DirectoryServices;

namespace webuntisnoten2atlantis
{
    class Program
    {        
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis17u;uid=";
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
            Global.PadRight = 116;
            Global.SqlZeilen = new List<string>();

            Global.WriteLine("*" + "---".PadRight(Global.PadRight, '-') + "--*");
            Global.WriteLine("| Webuntisnoten2Atlantis    |    Published under the terms of GPLv3    |    Stefan Bäumer   " + DateTime.Now.Year + "  |  Version 20230603  |");
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
                string sourceStudents = CheckFile(User, "Student_", DateTime.Now.Date.AddDays(-20));
                string sourceAbsenceTimesTotal = CheckFile(User, "AbsenceTimesTotal", DateTime.Now.Date);
                string sourceMarksPerLesson = CheckFile(User, "MarksPerLesson", DateTime.Now.Date);
                string sourceExportLessons = CheckFile(User, "ExportLessons", DateTime.Now.Date.AddDays(-20));
                string sourceStudentgroupStudents = CheckFile(User, "StudentgroupStudents", DateTime.Now.Date.AddDays(-20));

                var alleAtlantisLehrer = new Lehrers(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj);

                // Alle Webuntis-Leistungen, ohne Leistungen mit leerem Fach, leerer Klasse, Dopplungen bei Fach, Lehrer & Gesamtnote, mit Leistungen ohne Gesamtnote 

                var alleSchüler = new Schülers(sourceStudents);
                var alleUnterrichte = new Unterrichte(sourceExportLessons);
                var alleGruppen = new Gruppen(sourceStudentgroupStudents);
                var alleWebuntisLeistungen = new Leistungen(sourceMarksPerLesson, alleAtlantisLehrer);

                if (gehtEsUmBlaueBriefe())
                {
                    alleWebuntisLeistungen = alleWebuntisLeistungen.GetBlaueBriefeLeistungen();
                }

                var alleMöglicheKlassen = alleWebuntisLeistungen.GetMöglicheKlassen();

                do
                {
                    Global.SqlZeilen = new List<string>();
                    Global.Rückmeldung = new List<string>();
                    Global.Tabelle = new List<string>();
                    Global.Rückmeldungen = new Rückmeldungen();
                    
                    var interessierendeKlasse = GetIntessierendeKlasse(alleMöglicheKlassen, AktSj);

                    var targetAbsenceTimesTotal = Path.Combine(targetPath, Zeitstempel + "_AbsenceTimesTotal_" + interessierendeKlasse + "_" + User + ".CSV");
                    var targetMarksPerLesson = Path.Combine(targetPath, Zeitstempel + "_MarksPerLesson_" + interessierendeKlasse + "_" + User + ".CSV");

                    RelevanteDatensätzeAusCsvFilternUndProtokollErstellen(sourceAbsenceTimesTotal, targetAbsenceTimesTotal, interessierendeKlasse);
                    RelevanteDatensätzeAusCsvFilternUndProtokollErstellen(sourceMarksPerLesson, targetMarksPerLesson, interessierendeKlasse);

                    Leistungen atlantisLeistungen = new Leistungen();

                    Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuellUnsortiert = new Unterrichte();

                    var interessierendeSchülerDieserKlasse = alleSchüler.GetMöglicheSchülerDerKlasse(interessierendeKlasse);

                    interessierendeSchülerDieserKlasse.GetWebuntisUnterrichte(alleUnterrichte, alleGruppen, interessierendeKlasse, hzJz, AktSj);

                    interessierendeSchülerDieserKlasse.GetWebuntisLeistungen(alleWebuntisLeistungen);

                    interessierendeSchülerDieserKlasse.GetAtlantisLeistungen(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj, User, interessierendeKlasse, hzJz);

                    interessierendeSchülerDieserKlasse.ZweiLehrerEinFach(interessierendeKlasse);

                    interessierendeSchülerDieserKlasse.TabelleZeichnen(interessierendeKlasse, User);

                    interessierendeSchülerDieserKlasse.GeholteLeistungenBehandeln(interessierendeKlasse, AktSj, hzJz);

                    interessierendeSchülerDieserKlasse.TabelleZeichnen(interessierendeKlasse, User);

                    foreach (var zeile in Global.Tabelle)
                    {
                        Global.PrintMessage(Global.SqlZeilen.Count(), zeile);
                    }

                    interessierendeSchülerDieserKlasse.ChatErzeugen(alleAtlantisLehrer, interessierendeKlasse, hzJz, User);

                    // Add-Delete-Update

                    interessierendeSchülerDieserKlasse.Update(interessierendeKlasse);

                    // Abwesenheiten 

                    if (sourceAbsenceTimesTotal != null)
                    {                    
                        Global.WriteLine(" ");
                        Global.WriteLine("Abwesenheiten in der Klasse " + interessierendeKlasse + ":");
                        Global.WriteLine("==================================".PadRight(interessierendeKlasse.Length, '='));
                        Global.WriteLine(" ");

                        var webuntisAbwesenheiten = new Abwesenheiten(sourceAbsenceTimesTotal, interessierendeKlasse, atlantisLeistungen);

                        var atlantisAbwesenheiten = targetAbsenceTimesTotal == null ? null : new Abwesenheiten(ConnectionStringAtlantis + Properties.Settings.Default.DBUser, AktSj, interessierendeKlasse);
                        atlantisAbwesenheiten.Add(webuntisAbwesenheiten);
                        atlantisAbwesenheiten.Delete(webuntisAbwesenheiten);
                        atlantisAbwesenheiten.Update(webuntisAbwesenheiten);                        
                    }
                    else
                    {
                        int outputIndex = Global.SqlZeilen.Count();
                        Global.PrintMessage(outputIndex, ("Es werden keine Abwesenheiten importiert, da die Importdatei nicht von heute ist."));
                    }
                    
                    string targetSql = Path.Combine(targetPath, Zeitstempel + "_webuntisnoten2atlantis_" + interessierendeKlasse + "_" + User + ".SQL");
                    atlantisLeistungen.ErzeugeSqlDatei(new List<string>() { targetAbsenceTimesTotal, targetMarksPerLesson, targetSql });
                    OpenFiles(new List<string> { targetSql });

                    Console.WriteLine("\nSie können einen Screenshot der Notensammelerfassung in Atlantis erstellen und auf dem Desktop ablegen.\n" +
                        "Nachdem Sie die Datei gespeichert haben, können Sie hier wählen, " +
                        "ob eine verschlüsselte PDF-Datei erstellt werden soll, die dann z.B. per Teams verschickt werden kann.\n" +
                        "Wenn Sie keine PNG-Datei gespeichert haben und trotzdem Ja klicken, passiert nichts.");

                    ConsoleKeyInfo x;

                    Console.WriteLine(" ");
                    Console.Write("Ihre Auswahl(J,n):");
                    x = Console.ReadKey();

                    if (x.Key.ToString().ToLower() != "n")
                    {
                        PdfKennwort(interessierendeKlasse);
                    }
                    Console.WriteLine(" ");

                } while (true);
            }
            catch (Exception ex)
            {
                Global.WriteLine(ex.Message);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private static void PdfKennwort(string interessierendeK)
        {
            var directory = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            try
            {
                var fileName = (from f in directory.GetFiles()
                                where f.FullName.ToLower().EndsWith(".png")
                                where f.LastWriteTime > DateTime.Now.AddMinutes(-5)
                                orderby f.LastWriteTime descending
                                select f.FullName).First();

                Document document = new Document(new Rectangle(288f, 144f), 10, 10, 10, 10);
                document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());

                using (var stream = new FileStream(fileName + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    PdfWriter.GetInstance(document, stream);
                    document.Open();
                    using (var imageStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        var image = Image.GetInstance(imageStream);
                        image.SetAbsolutePosition(0, 0); // set the position to bottom left corner of pdf
                        image.ScaleAbsolute(iTextSharp.text.PageSize.A4.Height, iTextSharp.text.PageSize.A4.Width); // set the height and width of image to PDF page size
                        document.Add(image);
                    }
                    document.Close();
                }

                PdfSharp.Pdf.PdfDocument pdocument = PdfSharp.Pdf.IO.PdfReader.Open(fileName + ".pdf");
                PdfSecuritySettings securitySettings = pdocument.SecuritySettings;
                securitySettings.UserPassword = "!7765Neun";
                securitySettings.OwnerPassword = "!7765Neun";
                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = true;
                securitySettings.PermitFullQualityPrint = false;
                securitySettings.PermitModifyDocument = true;
                securitySettings.PermitPrint = false;
                pdocument.Save(directory + "\\" + interessierendeK + "-Notenliste-" + DateTime.Now.ToShortDateString() + "-Kennwort.pdf");
                File.Delete(fileName + ".pdf");
                
                Console.WriteLine("Schauen Sie auf dem Desktop nach einer Datei namens " + interessierendeK + "-Notenliste-" + DateTime.Now.ToShortDateString() + "-Kennwort.pdf");
            }
            catch (Exception)
            {
                Console.WriteLine("Es gibt keine PNG-Datei, die in den letzten 5 Minuten auf dem Desktop abgelegt wurde. Also wurde nichts verändert.");
                Console.WriteLine("");
            }
            Console.WriteLine("");
        }

        private static bool gehtEsUmBlaueBriefe()
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

        private static string GetIntessierendeKlasse(List<string> möglicheKlassen, List<string> aktSj)
        {
            do
            {
                Console.Write(" Bitte die interessierende Klasse angeben [" + GetVorschlag(möglicheKlassen) + "]: ");

                var x = Console.ReadLine().ToUpper();

                if (x == "")
                {
                    x = Properties.Settings.Default.InteressierendeKlassen;
                }
                if (möglicheKlassen.Contains(x))
                {
                    Properties.Settings.Default.InteressierendeKlassen = x;
                    Properties.Settings.Default.Save();
                    Console.WriteLine("  Ihre Auswahl: " + x);
                    Console.WriteLine(" ");
                    return x;
                }
            } while (true);
        }

        private static string GetVorschlag(List<string> möglicheKlassen)
        {
            string vorschlag = "";

            // Wenn in den Properties ein Eintrag existiert, ...

            if (Properties.Settings.Default.InteressierendeKlassen != null && Properties.Settings.Default.InteressierendeKlassen != "")
            {
                // ... wird für alle kommaseparierten Einträge geprüft ...


                if ((from t in möglicheKlassen where t == Properties.Settings.Default.InteressierendeKlassen select t).Any())
                {
                    // Falls ja, dann wird der Vorschlag aus den Properties übernommen.

                    vorschlag = Properties.Settings.Default.InteressierendeKlassen;
                }
            }
            Properties.Settings.Default.InteressierendeKlassen = vorschlag;
            Properties.Settings.Default.Save();
            return vorschlag.TrimEnd(',');
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
                    //notepadPlus.StartInfo.Arguments = @"-multiInst -nosession " + files[i];
                    notepadPlus.Start();
                }
            }
            catch (Exception)
            {
                System.Diagnostics.Process.Start("Notepad.exe", files[0]);
            }
        }

        private static void RelevanteDatensätzeAusCsvFilternUndProtokollErstellen(string sourceFile, string targetfile, string interessierendeKlasse)
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
                                                
                        if (line.Contains(interessierendeKlasse))
                        {
                            relevanteZeile = true;
                            anzahlZeilen++;
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

        private static string CheckFile(string user, string kriterium, DateTime zeitpunkt)
        {
            var sourceFile = (from f in Directory.GetFiles(@"c:\users\" + user + @"\Downloads", "*.csv", SearchOption.AllDirectories) where f.Contains(kriterium) orderby File.GetLastWriteTime(f) select f).LastOrDefault();

            if ((sourceFile == null || System.IO.File.GetLastWriteTime(sourceFile).Date < DateTime.Now.Date))
            {
                Console.WriteLine("");
                Console.WriteLine(" Die " + kriterium + "<...>.csv" + (sourceFile == null ? " existiert nicht im Download-Ordner" : " im Download-Ordner ist nicht von heute. \n Es werden keine Daten aus der Datei importiert") + ".");
                Console.WriteLine(" Exportieren Sie die Datei frisch aus Webuntis, indem Sie als Administrator:");

                if (kriterium.Contains("Student_"))
                {
                    Console.WriteLine("   1. Stammdaten > Schülerinnen");
                    Console.WriteLine("   2. \"Berichte\" auswählen");                    
                    Console.WriteLine("   3. Bei \"Schüler\" auf CSV klicken");
                    Console.WriteLine("   4. Die Datei \"Student_<...>.CSV\" im Download-Ordner zu speichern");
                    Console.WriteLine(" ");
                    Console.WriteLine(" ENTER beendet das Programm.");
                    Console.ReadKey();
                    Environment.Exit(0);
                }
                else
                {
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
                    else
                    {
                        Console.WriteLine("   1. Administration > Export klicken");
                        Console.WriteLine("   2. Zeitraum begrenzen, also die Woche der Zeugniskonferenz und vergange Abschnitte herauslassen");
                        Console.WriteLine("   2. Das CSV-Icon hinter Gesamtfehlzeiten klicken");
                    }

                    if (kriterium.Contains("AbsenceTimesTotal"))
                    {
                        Console.WriteLine("   4. Die Gesamtfehlzeiten (\"AbsenceTimesTotal<...>.CSV\") im Download-Ordner zu speichern");
                        Console.WriteLine("WICHTIG: Es kann Sinn machen nur Abwesenheiten bis zur letzten Woche in Webuntis auszuwählen.");
                    }

                    if (kriterium.Contains("StudentgroupStudents"))
                    {
                        Console.WriteLine("   4. Die Schülergruppen  (\"StudentgroupStudents<...>.CSV\") im Download-Ordner zu speichern");
                    }

                    if (kriterium.Contains("ExportLessons"))
                    {
                        Console.WriteLine("   4. Die Unterrichte (\"ExportLessons<...>.CSV\") im Download-Ordner zu speichern");
                    }
                }

                Console.WriteLine(" ");
                sourceFile = null;
            }

            if (sourceFile != null)
            {
                Console.WriteLine((Path.GetFileName(sourceFile) + " ").PadRight(73, '.') + ". Erstell-/Bearbeitungszeitpunkt heute um " + System.IO.File.GetLastWriteTime(sourceFile).ToShortTimeString());
            }

            return sourceFile;
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
            
            Console.WriteLine(("|     Pfad zu den Dateien: " + pfad).PadRight(Global.PadRight,' ') + "   |");
            Console.WriteLine("*" + "---".PadRight(Global.PadRight, '-') + "--*");
            Console.WriteLine(@" ");

            return pfad;
        }
    }
}