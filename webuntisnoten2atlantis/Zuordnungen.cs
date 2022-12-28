﻿// Published under the terms of GPLv3 Stefan Bäumer 2023.

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



        public void AlleZulässigenAtlantisZielFächerAuflisten(Leistungen atlantisleistungen, string aktSj)
        {
            var atlantisFächer = "";

            foreach (var quellklasse in (from z in this select z.Quellklasse).Distinct().ToList())
            {
                foreach (var atlantisfach in (from a in atlantisleistungen
                                              where a.Klasse == quellklasse
                                              where a.Schuljahr == aktSj
                                              select a.Fach).Distinct().ToList())
                {
                    atlantisFächer += atlantisfach + ",";
                }

                atlantisFächer = atlantisFächer.TrimEnd(',');

                Console.WriteLine(" Für Auswahl zulässige Atlantisfächer (Klasse:" + quellklasse + "; Schuljahr: " + aktSj + "): " + atlantisFächer);
            }
        }

        internal void GespeicherteZuordnungenInFehlendenZuordnungenAusPropertiesErgänzen()
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

            // Die fehlenden Zuordnungen werden durch die Speocherungen ergänzt:

            foreach (var fehlendeZuordnung in this)
            {
                var gespeicherteZuordnung = (from g in gespeicherteZuordnungen
                                             where g.Quellfach == fehlendeZuordnung.Quellfach
                                             where g.Quellklasse == fehlendeZuordnung.Quellklasse
                                             select g).FirstOrDefault();

                // Falls es eine gespeicherte Zuordnung gibt ...

                if (gespeicherteZuordnung != null)
                {
                    // ... wird die fehlende Zuordnung angepasst:

                    fehlendeZuordnung.Zielfach = gespeicherteZuordnung.Zielfach;
                    fehlendeZuordnung.Zielklasse = gespeicherteZuordnung.Zielklasse;
                }
            }
        }

        internal void SpeichernInDenProperties()
        {
            Properties.Settings.Default.Zuordnungen = "";
            Properties.Settings.Default.Save();

            foreach (var item in this)
            {
                if (item.Zielfach != null)
                {
                    Properties.Settings.Default.Zuordnungen += item.Quellklasse + "|" + item.Quellfach + ";" + item.Zielfach + ",";
                }
            }

            Properties.Settings.Default.Zuordnungen = Properties.Settings.Default.Zuordnungen.TrimEnd(',');
            Properties.Settings.Default.Save();
        }

        internal void AlleNichtAutomatischZugewiesenenFächerAnzeigen()
        {
            if (this.Count > 0)
            {
                Console.WriteLine("");
                Console.WriteLine(" Folgende Fächer können keinem Atlantisfach zugeordnet werden oder wurden bereits manuell zugeordnet:");

                for (int i = 0; i < this.Count; i++)
                {
                    Console.WriteLine((" " + (i + 1).ToString().PadLeft(2) + ". " + this[i].Quellklasse.PadRight(6) + "|" + this[i].Quellfach.PadRight(6) + (this[i].Zielfach != null ? "   ->  " + this[i].Zielfach : "")).PadRight(34));
                }
            }
        }

        internal string DialogZurÄnderungAufrufen(Leistungen atlantisleistungen, string x)
        {
            int n;

            if (int.TryParse(x, out n))
            {
                var eingabe = int.Parse(x);

                if (eingabe > 0 && eingabe <= this.Count)
                {
                    Console.Write(" Wie heißt das Atlantisfach in der Klasse " + this[eingabe - 1].Quellklasse + ", dem Sie das Untis-Fach *" + this[eingabe - 1].Quellfach + "* zuordnen wollen? ");

                    var xx = Console.ReadLine();
                    xx = xx.ToUpper();

                    Console.WriteLine("");

                    if ((from a in atlantisleistungen where a.Klasse == this[eingabe - 1].Quellklasse where a.Fach == xx select a).Any() || xx == "")
                    {
                        this[eingabe - 1].Zielfach = xx;

                        int i = 0;

                        Console.WriteLine(" Die Zuordnung des Faches " + xx + " wurde " + i + "x erfolgreich vorgenommen.");

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
            return x;
        }
    }
}