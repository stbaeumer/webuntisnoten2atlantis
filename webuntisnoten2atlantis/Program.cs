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
                string outputSql = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\webuntisnoten2atlantis_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".SQL";
                
                List<string> aktSj = new List<string>
                {
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
                    (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 - 2000 : DateTime.Now.Year - 2000).ToString()
                };

                Console.WriteLine("Webuntisnoten2atlantis (Version 20190914)");
                Console.WriteLine("=========================================");
                Console.WriteLine("");

                Console.WriteLine("Noten aus Webuntis exportieren:");
                Console.WriteLine(" 1. Klasenbuch > Berichte");
                Console.WriteLine(" 2. Alle Klassen auswählen");
                Console.WriteLine(" 3. Unter \"Noten\" Prüfungsart HZ auswählen");
                Console.WriteLine(" 4. Hinter \"Noten pro Schüler\" auf CSV klicken.");
                Console.WriteLine(" 5. Die Datei \"MarksPerLesson.csv\" auf dem Desktop speichern.");

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

                Leistungen webuntisLeistungen = new Leistungen(inputCsv);
                Leistungen atlantisLeistungen = new Leistungen(ConnectionStringAtlantis, aktSj[0] + "/" + aktSj[1], outputSql);

                Console.WriteLine("");

                atlantisLeistungen.NeuZuSetzendeNoten(webuntisLeistungen);
                atlantisLeistungen.ZuLöschendeNoten(webuntisLeistungen);
                atlantisLeistungen.ZuÄnderendeNoten(webuntisLeistungen);

                atlantisLeistungen.ErzeugeSlqDatei(outputSql);

                Console.WriteLine("Beenden mit ANYKEY");
                Console.ReadKey();
            }            
            catch (Exception ex)
            {
                Console.WriteLine("Heiliger Bimam! Es ist etwas schiefgelaufen! Kaum zu glauben, wenn man doch weiß wer's programmiert hat! Die Verarbeitung wird gestoppt.");
                Console.WriteLine("");
                Console.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
        }
            
    }
}
