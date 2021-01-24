// Published under the terms of GPLv3 Stefan Bäumer 2020.

using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntisnoten2atlantis
{
    public class Zuordnungen : List<Zuordnung>
    {
        private Leistungen atlantisLeistungen;

        public Zuordnungen()
        {
        }

        public Zuordnungen(Leistungen webuntisLeistungen, Abwesenheiten atlantisAbwesenheiten)
        {
            this.Add(new Zuordnung(" FU", ""));
            this.Add(new Zuordnung("FS19*|INW", "IF"));
            this.Add(new Zuordnung("HBT19A|FPE", "FPEL"));
            this.Add(new Zuordnung("WE18A|KKS*", "KKS"));
            this.Add(new Zuordnung("KR", "REL"));
            this.Add(new Zuordnung("ER", "REL"));
            this.Add(new Zuordnung("KR ", "REL"));
            this.Add(new Zuordnung("ER ", "REL"));
            this.Add(new Zuordnung("NF ", ""));
            this.Add(new Zuordnung("", ""));
            this.Add(new Zuordnung("", ""));
            this.Add(new Zuordnung("", ""));

            

            //            // Sonderfall Niederländisch-Kurs
            //            // Wenn es mehrere Niederländisch-Kurse gibt, wird die Webuntis-Note allen Atlantis-Noten zugeordnet

            //            if (this[j].Fach.Contains(" G") && this[j].Fach.Split(' ')[0] == "NF")
            //            {
            //                var aa = (from aaa in atlantisFächerDiesesSchuelers where aaa.Fach.StartsWith("N") where aaa.Fach.Contains(" G") select aaa).ToList();

            //                if (aa != null)
            //                {
            //                    this[j].Beschreibung += this[j].Fach + "->" + aa[0].Fach + ",";
            //                    this[j].Fach = aa[0].Fach;
            //                    zugeordnet = true;
            //                    i++;

            //                    if (aa.Count > 1)
            //                    {
            //                        Leistung wel = new Leistung();

            //                        wel.Abschlussklasse = this[j].Abschlussklasse;
            //                        wel.Anlage = this[j].Anlage;
            //                        wel.Bemerkung = this[j].Bemerkung;
            //                        wel.Beschreibung = this[j].Fach + "->" + aa[1].Fach + ",";
            //                        wel.Datum = this[j].Datum;
            //                        wel.EinheitNP = this[j].EinheitNP;
            //                        wel.Fach = aa[1].Fach;
            //                        wel.GeholteNote = this[j].GeholteNote;
            //                        wel.Gesamtnote = this[j].Gesamtnote;
            //                        wel.Gesamtpunkte = this[j].Gesamtpunkte;
            //                        wel.Gliederung = this[j].Gliederung;
            //                        wel.HatBemerkung = this[j].HatBemerkung;
            //                        wel.HzJz = this[j].HzJz;
            //                        wel.Jahrgang = this[j].Jahrgang;
            //                        wel.Klasse = this[j].Klasse;
            //                        wel.Konferenzdatum = this[j].Konferenzdatum;
            //                        wel.Lehrkraft = this[j].Lehrkraft;
            //                        wel.LeistungId = this[j].LeistungId;
            //                        wel.Name = this[j].Name;
            //                        wel.ReligionAbgewählt = this[j].ReligionAbgewählt;
            //                        wel.SchlüsselExtern = this[j].SchlüsselExtern;
            //                        wel.SchuelerAktivInDieserKlasse = this[j].SchuelerAktivInDieserKlasse;
            //                        wel.Schuljahr = this[j].Schuljahr;
            //                        wel.Tendenz = this[j].Tendenz;
            //                        wel.Zeugnisart = this[j].Zeugnisart;
            //                        wel.Zeugnistext = this[j].Zeugnistext;
            //                        this.Add(wel);
            //                        i++;
            //                    }
            //                }
            //                break;
            //            }

            //            // Evtl. hängt die Niveaustufe in Untis am Namen 

            //            if (a.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "").Replace("KA2", "") == this[j].Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "").Replace("KA2", ""))
            //            {
            //                this[j].Beschreibung += this[j].Fach + "->" + a.Fach + ",";
            //                zugeordnet = true;
            //                i++;
            //                break;
            //            }

            //            // Es wird versucht auf die ersten beiden Buchstaben zu matchen, aber nicht, wenn das Webuntis-Fach eine Leerstelle enthält.
            //            // 

            //            if (a.Fach.Substring(0, Math.Min(2, a.Fach.Length)) == this[j].Fach.Substring(0, Math.Min(2, this[j].Fach.Length)))
            //            {
            //                // Zuordnung nur, wenn nicht ein anderes Fach besser passt. Das kann sein, wenn dieses Fach die FP zu einem anderen Fach ist.

            //                if (!(from af in atlantisFächerDiesesSchuelers where af.Fach == this[j].Fach select af).Any())
            //                {
            //                    this[j].Beschreibung += this[j].Fach + "->" + a.Fach + ",";
            //                    zugeordnet = true;
            //                    i++;
            //                    break;
            //                }
            //            }

            //            // EK -> E

            //            if (this[j].Fach.StartsWith("EK") && a.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "E")
            //            {
            //                this[j].Beschreibung += this[j].Fach + "->" + a.Fach + ",";
            //                zugeordnet = true;
            //                i++;
            //                break;
            //            }

            //            // N -> NKA1

            //            if (this[j].Fach.StartsWith("N") && a.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "NK")
            //            {
            //                this[j].Beschreibung += this[j].Fach + "->" + a.Fach + ",";
            //                zugeordnet = true;
            //                i++;
            //                break;
            //            }

            //            // IF -> WI

            //            if (this[j].Fach == "IF" && a.Fach == "WI")
            //            {
            //                this[j].Beschreibung += this[j].Fach + "->" + a.Fach + ",";
            //                zugeordnet = true;
            //                i++;
            //                break;
            //            }
            //        }

            //        if (!zugeordnet && atlantisFächerDiesesSchuelers.Count > 0 && this[j].Gesamtnote != null)
            //        {
            //            if (!nichtZugeordneteFächer.Contains(this[j].Klasse + "|" + this[j].Fach))
            //            {
            //                nichtZugeordneteFächer.Add(this[j].Klasse + "|" + this[j].Fach);
            //            }
            //        }
            //    }

        }

        public void ManuellZuordnen(Leistungen webuntisleistungen, Leistungen atlantisleistungen)
        {
            var gespeicherteZuordnungen = new Zuordnungen();

            foreach (var item in Properties.Settings.Default.Zuordnungen.Split(','))
            {
                if (item != "")
                {
                    var quelle = item.Split(';')[0];
                    var ziel = item.Split(';')[1];

                    Zuordnung zuordnung = new Zuordnung();
                    zuordnung.Quellklasse = quelle.Split('|')[0];
                    zuordnung.Quellfach = quelle.Split('|')[1];
                    zuordnung.Zielfach = ziel.Split('|')[0];
                    gespeicherteZuordnungen.Add(zuordnung);
                }                
            }

            // 

            foreach (var t in this)
            {
                var x = (from g in gespeicherteZuordnungen where g.Quellklasse == t.Quellklasse where g.Quellfach == t.Quellfach where g.Zielfach != null select g).FirstOrDefault();
                if (x != null)
                {
                    // Nur wenn es diese Zuordnung noch immer gibt:

                    if ((from a in atlantisleistungen where a.Klasse == t.Quellklasse where a.Fach == x.Zielfach select a).Any())
                    {
                        t.Zielfach = x.Zielfach;
                    }                    
                }
            }

            // Wenn keine Zuordnung vorgenommen werden konnte

            var liste = new List<string>();

            if (this.Count > 0)
            {
                ConsoleKeyInfo x;

                Console.Clear();

                do
                {                    
                    Console.WriteLine("Folgende Fächer können keinem Atlantisfach zugeordnet werden oder wurden bereits manuell zugeordnet:");

                    liste = new List<string>();

                    for (int i = 0; i < this.Count; i++)
                    {                        
                        Console.Write((" " + (i + 1).ToString().PadLeft(2) + ". " + this[i].Quellklasse.PadRight(6) + "|" + this[i].Quellfach.PadRight(6) + (this[i].Zielfach != null ? "   ->  " + this[i].Zielfach : "")).PadRight(34));
                        if ((i + 1) % 3 == 0)
                        {
                            Console.WriteLine("");
                        }
                        liste.Add((" " + (i + 1).ToString().PadLeft(2) + ". " + this[i].Quellklasse.PadRight(6) + "|" + this[i].Quellfach.PadRight(6) + (this[i].Zielfach != null ? "   ->  " + this[i].Zielfach : "")).PadRight(32));
                    }
                    Console.WriteLine("");
                    Console.Write("Wollen Sie eine Zuordnung manuell vornehmen? Wählen Sie [1" + (this.Count > 1 ? ", ..., " + this.Count : "") + "] oder ENTER, falls keine Änderung gewünscht ist: ");

                    x = Console.ReadKey();

                    Console.WriteLine("");

                    if (char.IsDigit(x.KeyChar))
                    {
                        var eingabe = int.Parse(x.KeyChar.ToString());

                        if (eingabe > 0 && eingabe <= this.Count)
                        {                            
                            Console.Write("Wie heißt das Atlantisfach in der Klasse " + this[eingabe - 1].Quellklasse + ", dem Sie das Untis-Fach *" + this[eingabe - 1].Quellfach + "* zuordnen wollen? ");

                            var xx = Console.ReadLine();
                            xx = xx.ToUpper();
                            
                            Console.WriteLine("");
                            Console.Clear();

                            if ((from a in atlantisleistungen where a.Klasse == this[eingabe - 1].Quellklasse where a.Fach == xx select a).Any() || xx == "")
                            {                                
                                this[eingabe - 1].Zielfach = xx;                                                                                                
                                Console.WriteLine("Die Zuordnung des Faches " + xx + " wurde erfolgreich vorgenommen.");
                                
                                if (xx == "")
                                {
                                    this[eingabe - 1].Zielfach = null;
                                    Console.WriteLine("Die Zuordnung wird entfernt.");
                                }
                            }
                            else
                            {   
                                Console.WriteLine("[FEHLER] Ihre Zuordnung war nicht möglich. Das Fach *" + xx + "* gibt es in Atlantis nicht. Die Fächer sind:");
                                
                                var verschiedeneFächer = (from a in atlantisleistungen
                                                          where a.Klasse == this[eingabe - 1].Quellklasse
                                                          select a.Fach).Distinct().ToList();

                                for (int i = 0; i < verschiedeneFächer.Count; i++)
                                {
                                    Console.Write("  " + verschiedeneFächer[i].PadRight(7));
                                    
                                    if ((i + 1) % 7 == 0)
                                    {
                                        Console.WriteLine("");
                                    }
                                }
                                Console.WriteLine("");
                            }
                        }
                    }

                } while (x.Key != ConsoleKey.Enter);

                // ZUordnung zu den Properties und den Webuntisfächern

                

                Properties.Settings.Default.Zuordnungen = "";
                Properties.Settings.Default.Save();

                int vorgenommeneZuordnung = 0;
                int keineZuordnung = 0;
                foreach (var item in this)
                {
                    if (item.Zielfach != null)
                    {
                        if ((from a in atlantisleistungen where a.Klasse ==item.Quellklasse where a.Fach == item.Zielfach select a).Any())
                        {
                            Properties.Settings.Default.Zuordnungen += item.Quellklasse + "|" + item.Quellfach + ";" + item.Zielfach + ",";

                            var we = (from w in webuntisleistungen where w.Klasse == item.Quellklasse where w.Fach == item.Quellfach select w).FirstOrDefault();

                            we.Beschreibung += we.Fach + " -> "+ item.Zielfach;

                            we.Fach = item.Zielfach;
                            
                            vorgenommeneZuordnung++;
                        }                        
                    }
                    else
                    {                    
                        keineZuordnung++;
                    }
                }

                Properties.Settings.Default.Zuordnungen = Properties.Settings.Default.Zuordnungen.TrimEnd(',');
                Properties.Settings.Default.Save();

                new Leistungen().AusgabeSchreiben("Es wurden " + vorgenommeneZuordnung + "x Fächer aus Webuntis einem Atlantisfach zugeordnet. " + keineZuordnung + "x wurde keine Zuordnung vorgenommen:", liste);
            }
        }
    }
}