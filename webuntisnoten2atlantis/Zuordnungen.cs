// Published under the terms of GPLv3 Stefan Bäumer 2021.

using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntisnoten2atlantis
{
    public class Zuordnungen : List<Zuordnung>
    {
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
                string x;

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
                    Console.Write("Wollen Sie eine Zuordnung manuell vornehmen? Wählen Sie [1" + (this.Count > 1 ? ", ... " + this.Count : "") + "] oder ENTER, falls keine Änderung gewünscht ist: ");

                    x = Console.ReadLine();

                    Console.WriteLine("");
                    int n;

                    if (int.TryParse(x, out n))
                    {
                        var eingabe = int.Parse(x);

                        if (eingabe > 0 && eingabe <= this.Count)
                        {
                            // Alle Atlantisfächer werden aufgelistet

                            var atlantisFächer = "";

                            foreach (var atlantisfach in (from a in atlantisleistungen where a.Klasse == this[eingabe - 1].Quellklasse select a.Fach).Distinct().ToList())
                            {
                                atlantisFächer += atlantisfach + ",";
                            }

                            atlantisFächer = atlantisFächer.TrimEnd(',');

                            Console.WriteLine("Alle Atlantisfächer: " + atlantisFächer);

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

                } while (x != "");

                // ZUordnung zu den Properties und den Webuntisfächern

                Properties.Settings.Default.Zuordnungen = "";
                Properties.Settings.Default.Save();

                int vorgenommeneZuordnung = 0;
                int keineZuordnung = 0;
                foreach (var item in this)
                {
                    if (item.Quellfach == "EUS F")
                    {
                        string aaaa = "";
                    }

                    if (item.Zielfach != null)
                    {
                        if ((from a in atlantisleistungen where a.Klasse == item.Quellklasse where a.Fach == item.Zielfach select a).Any())
                        {
                            Properties.Settings.Default.Zuordnungen += item.Quellklasse + "|" + item.Quellfach + ";" + item.Zielfach + ",";

                            var we = (from w in webuntisleistungen where w.Klasse == item.Quellklasse where w.Fach == item.Quellfach select w).ToList();

                            foreach (var w in we)
                            {
                                w.Beschreibung += w.Fach + " -> " + item.Zielfach;

                                w.Fach = item.Zielfach;

                            }
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

                string aus = "Es wurden " + vorgenommeneZuordnung + "x Fächer aus Webuntis einem Atlantisfach zugeordnet. " + keineZuordnung + "x wurde keine Zuordnung vorgenommen:";

                if (keineZuordnung == 0)
                {
                    aus = "Es wurden alle Fächer aus Webuntis einem Atlantisfach wie folgt zugeordnet:";
                }

                    
                new Leistungen().AusgabeSchreiben(aus, liste);
            }
        }
    }
}