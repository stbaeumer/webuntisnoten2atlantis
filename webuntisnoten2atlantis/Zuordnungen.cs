// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntisnoten2atlantis
{
    public class Zuordnungen : List<Zuordnung>
    {
        public Zuordnungen()
        {
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
                    this.Add(zuordnung);
                }
            }
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

        public List<string> AlleZulässigenAtlantisZielFächerAuflisten(Leistung dieseLeistung, Leistungen atlantisleistungen, string aktSj)
        {
            List<string> alle = new List<string>();

            var i = 1;

            foreach (var quellklasse in (from z in atlantisleistungen select z.Klasse).Distinct().ToList())
            {
                Console.WriteLine(" Für Auswahl zulässige Atlantisfächer im Schuljahr: " + aktSj + "): ");

                foreach (var atlantisfach in (from a in atlantisleistungen
                                              where a.Klasse == quellklasse
                                              where a.Schuljahr == aktSj
                                              select a.Fach).Distinct().ToList())
                {

                    var zielLeistung = (from t in this 
                                    where t.Quellklasse == dieseLeistung.Klasse 
                                    where t.Quellfach == dieseLeistung.Fach 
                                    where t.Zielfach == atlantisfach 
                                    select t).FirstOrDefault(); 
                    
                    Console.WriteLine( i.ToString().PadLeft(3) + ". " + (quellklasse + "|" + atlantisfach).PadRight(13) + (zielLeistung != null ? " <<<<= " + dieseLeistung.Fach : "")); i++;
                    alle.Add(atlantisfach);
                }
            }
            return alle;
        }

        internal void SpeichernInDenProperties()
        {
            foreach (var item in this)
            {
                Properties.Settings.Default.Zuordnungen += item.Quellklasse + "|" + item.Quellfach + ";" + item.Zielfach + ",";
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
            if (int.TryParse(x, out int n))
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