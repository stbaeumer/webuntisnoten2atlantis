// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntisnoten2atlantis
{
    public class Zuordnungen : List<Zuordnung>
    {        
        public List<string> AlleZulässigenAtlantisZielFächerAuflisten(Leistung dieseLeistung, Leistungen atlantisleistungen, string aktSj)
        {
            List<string> alle = new List<string>();

            var i = 1;

            //Global.AufConsoleSchreiben(" Für Auswahl zulässige Atlantisfächer im Schuljahr: " + aktSj + ": ");

            foreach (var quellklasse in (from z in atlantisleistungen where z.Klasse == dieseLeistung.Klasse select z.Klasse).Distinct().ToList())
            {
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
                
        internal void AlleNichtAutomatischZugewiesenenFächerAnzeigen()
        {
            if (this.Count > 0)
            {
                Global.WriteLine("");
                Global.WriteLine(" Folgende Fächer können keinem Atlantisfach zugeordnet werden oder wurden bereits manuell zugeordnet:");

                for (int i = 0; i < this.Count; i++)
                {
                    Global.WriteLine((" " + (i + 1).ToString().PadLeft(2) + ". " + this[i].Quellklasse.PadRight(6) + "|" + this[i].Quellfach.PadRight(6) + (this[i].Zielfach != null ? "   ->  " + this[i].Zielfach : "")).PadRight(34));
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

                    Global.WriteLine("");

                    if ((from a in atlantisleistungen where a.Klasse == this[eingabe - 1].Quellklasse where a.Fach == xx select a).Any() || xx == "")
                    {
                        this[eingabe - 1].Zielfach = xx;

                        int i = 0;

                        Global.WriteLine(" Die Zuordnung des Faches " + xx + " wurde " + i + "x erfolgreich vorgenommen.");

                        if (xx == "")
                        {
                            this[eingabe - 1].Zielfach = null;
                            Global.WriteLine("Die Zuordnung wird entfernt.");
                        }
                    }
                    else
                    {
                        Global.WriteLine("[FEHLER] Ihre Zuordnung war nicht möglich. Das Fach *" + xx + "* gibt es in Atlantis nicht. Die Fächer sind:");

                        var verschiedeneFächer = (from a in atlantisleistungen
                                                  where a.Klasse == this[eingabe - 1].Quellklasse
                                                  select a.Fach).Distinct().ToList();

                        for (int i = 0; i < verschiedeneFächer.Count; i++)
                        {
                            Console.Write("  " + verschiedeneFächer[i].PadRight(7));

                            if ((i + 1) % 7 == 0)
                            {
                                Global.WriteLine("");
                            }
                        }
                        Global.WriteLine("");
                    }
                }
            }
            return x;
        }
    }
}