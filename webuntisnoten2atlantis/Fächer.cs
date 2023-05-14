// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntisnoten2atlantis
{
    public class Fächer : List<Fach>
    {
        public Fächer()
        {
        }
                

        internal void KonferenzdatenZeileErzeugen()
        {
            string f = "|     |      |";

            foreach (var konferenzdatum in (from ff in this select ff.Konferenzdatum).Distinct().ToList())
            {
                var breite = Math.Max(2, (from ff in this where ff.Konferenzdatum == konferenzdatum select ff).Count());
                
                if (konferenzdatum.Year == 1) 
                {
                    f += "aus Webuntis".Substring(0,Math.Min(12, breite * 4 - 1)).PadRight(breite * 4 - 1) + "|";
                }
                else 
                {
                    f += (konferenzdatum.Month.ToString("00") + konferenzdatum.Year.ToString("00")).PadRight(breite * 4 - 1) + "|";
                }                
            }

            Global.WriteLine(f);
        }

        internal void FächerzeileErzeugen()
        {
            string f = "|     |      |";

            foreach (var konferenzdatum in (from ff in this select ff.Konferenzdatum).Distinct().ToList())
            {
                // Wenn nur ein Fach zu dieser Konferenz zugeordnet werden kann, dann wird die Spalte um eine leere Einheit verbreitert

                var breite = (from ff in this where ff.Konferenzdatum == konferenzdatum select ff).Count() == 1 ? 7 : 3;

                var x = (from ff in this where ff.Konferenzdatum == konferenzdatum select ff).ToList();

                foreach (var fach in (from t in this where t.Konferenzdatum == konferenzdatum select t).ToList())
                {
                    f += fach.Name.PadRight(3).Substring(0,3).PadRight(breite) + "|";
                }
            }

            Global.WriteLine(f);

            f = "|     |      |";

            foreach (var konferenzdatum in (from ff in this select ff.Konferenzdatum).Distinct().ToList())
            {
                // Wenn nur ein Fach zu dieser Konferenz zugeordnet werden kann, dann wird die Spalte um eine leere Einheit verbreitert

                var breite = (from ff in this where ff.Konferenzdatum == konferenzdatum select ff).Count() == 1 ? 7 : 3;

                foreach (var fach in (from t in this where t.Konferenzdatum == konferenzdatum select t).ToList())
                {
                    f += (fach.Lehrkraft == null ? "" :fach.Lehrkraft).PadRight(breite) + "|";
                }
            }

            Global.WriteLine(f);
        }

        internal void SchülerzeilenErzeugen(Leistungen webuntisLeistungen, Leistungen geholteLeistungen, int schülerId)
        {
            var alleAktuellenLeistungenDesSchülers = (from w in webuntisLeistungen where w.SchlüsselExtern == schülerId select w).ToList();

            var alleGeholtenLeistungenDesSchülers = (from g in geholteLeistungen where g.SchlüsselExtern == schülerId select g).ToList();

            string f = "|"+ (from w in webuntisLeistungen where w.SchlüsselExtern==schülerId select w.Name.PadRight(5).Substring(0,5)).FirstOrDefault()  + "|" + schülerId + "|";

            foreach (var konferenzdatum in (from ff in this select ff.Konferenzdatum).Distinct().ToList())
            {
                // Wenn nur ein Fach zu dieser Konferenz zugeordnet werden kann, dann wird die Spalte um eine leere Einheit verbreitert

                var breite = (from ff in this where ff.Konferenzdatum == konferenzdatum select ff).Count() == 1 ? 7 : 3;

                foreach (var fach in (from t in this where t.Konferenzdatum == konferenzdatum select t).ToList())
                {
                    // Wenn es zu diesemFach eine Schülerleistung gibt

                    var leistungen = (from a in alleAktuellenLeistungenDesSchülers where a.Fach == fach.Name select a).ToList();

                    leistungen.AddRange((from a in alleGeholtenLeistungenDesSchülers where a.Fach == fach.Name where a.Konferenzdatum == konferenzdatum select a).ToList());
                                        
                    f += (leistungen.Count == 0? "": leistungen.FirstOrDefault().Gesamtnote + leistungen.FirstOrDefault().Tendenz).PadRight(3).Substring(0, 3).PadRight(breite) + "|";
                }
            }
            Global.WriteLine(f.PadRight(8));
        }

        //internal void AddAktuelleFächer(Schülers schülers, string interessierendeKlasse)
        //{
        //    this.AddRange(new Fächer(schülers, interessierendeKlasse));
        //}

        //internal void AddAlteFächer(Leistungen geholteLeistungen, string interessierendeKlasse)
        //{
        //    foreach (var schülerId in (from t in geholteLeistungen where t.SchlüsselExtern != 0 select t.SchlüsselExtern).Distinct().ToList())
        //    {
        //        foreach (var atlantisleistung in (from a in geholteLeistungen                                                   
        //                                          where a.SchlüsselExtern == schülerId 
        //                                          where a.Klasse.Substring(0,2) == this[0].Klasse.Substring(0,2) // nur Leistungen des selben BG zählen. 
        //                                          select a).ToList())
        //        {
        //            if (!(from f in this where f.Name == atlantisleistung.Fach select f).Any())
        //            {
        //                this.Add(new Fach(atlantisleistung.Klasse, atlantisleistung.Fach, "", atlantisleistung.Konferenzdatum));
        //            }
        //        }
        //    }
        //}
    }
}