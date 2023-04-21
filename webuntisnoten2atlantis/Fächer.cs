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

        public Fächer(Leistungen webuntisLeistungen, string klasse)
        {
            var verschiedeneAktuelleFächer = (from s in webuntisLeistungen.OrderBy(x=>x.Klasse).ThenBy(x=>x.Fach) where s.Klasse == klasse select s.Fach).Distinct().ToList();

            // Alle aktuellen Fächer werden gesammelt.

            foreach (var aF in verschiedeneAktuelleFächer)
            {
                var x = (from s in webuntisLeistungen where s.Fach == aF where s.Klasse == klasse select s).FirstOrDefault();
                Fach fach = new Fach(x.Klasse, x.Fach, x.Lehrkraft, x.Konferenzdatum);
                this.Add(fach);
            }
        }

        public Leistungen GetAlleVergangenenLeistungen(Leistungen atlantisLeistungen, Leistungen webuntisLeistungen, List<int> alleVerschiedenenInteressierendenSchüler, List<DateTime> alleVerschiedenenAlteKonferenzdaten)
        {   
            var leistungen = new Leistungen();

            var f = new Fächer();
            f.AddRange(this);

            var schuelerLeistungExistiertSchon = new List<string>();

            foreach (var schülerId in alleVerschiedenenInteressierendenSchüler)
            {
                if (schülerId == 153274)
                {
                    string a = "";
                }
                foreach (var aktuellesFach in f)
                {
                    foreach (var alias in aktuellesFach.NameAliase)
                    {
                        schuelerLeistungExistiertSchon.Add(schülerId + alias.Replace("  ", " "));
                    }                    
                }

                var aktuelleFächerDiesesSchülers = (from s in webuntisLeistungen where s.SchlüsselExtern == schülerId select s.Fach).Distinct().ToList();
                var alteFächerDiesesSchülers = (from s in atlantisLeistungen where s.SchlüsselExtern == schülerId where !aktuelleFächerDiesesSchülers.Contains(s.Fach) select s.Fach).Distinct().ToList();
                var aktuelleKlasseDiesesSchülers = (from s in webuntisLeistungen where s.SchlüsselExtern == schülerId select s.Klasse).FirstOrDefault();

                // Für alle vergangenen Konferenzdaten, ...

                foreach (var konferenzdatum in alleVerschiedenenAlteKonferenzdaten)
                {
                    // ... werden alle vergangenen Fächer des Schülers durchlaufen.

                    foreach (var fach in (from s in Global.GeholteLeistungen.OrderBy(x => x.Klasse).ThenBy(x => x.Fach) where s.SchlüsselExtern == schülerId where !aktuelleFächerDiesesSchülers.Contains(s.Fach) where konferenzdatum == s.Konferenzdatum select s.Fach).Distinct().ToList())
                    {
                        // Wenn bereits eine Note in diesem Fach vorliegt, wird keine weitere Note gezogen

                        if (!schuelerLeistungExistiertSchon.Contains(schülerId + fach.Replace("  ", " ")))
                        {
                            schuelerLeistungExistiertSchon.Add(schülerId + fach.Replace("  ", " "));

                            var l = (from s in Global.GeholteLeistungen
                                     where s.SchlüsselExtern == schülerId
                                     where !aktuelleFächerDiesesSchülers.Contains(s.Fach)
                                     where s.Konferenzdatum == konferenzdatum
                                     where s.Fach == fach
                                     select s).ToList();

                            // Bei Religionsabwählern wird keine Religion als Leistung angelegt.

                            if (!(l[0].ReligionAbgewählt && (l[0].Fach == "KR" || l[0].Fach =="ER" || l[0].Fach == "REL" || l[0].Fach.StartsWith("ER ") || l[0].Fach.StartsWith("KR "))))
                            {
                                leistungen.Add(l[0]);

                                // Wenn diese Leistung an einem  

                                if (!(from xx in this
                                      where xx.Klasse.Substring(0, 3) == l[0].Klasse.Substring(0, 3) // Klassenwechsler nehmen die Note aus Parrallelklasse mit 
                                      where xx.Name == l[0].Fach
                                      where xx.Lehrkraft == l[0].Lehrkraft
                                      select xx).Any())
                                {
                                    this.Add(new Fach(l[0].Klasse, l[0].Fach, l[0].Lehrkraft, l[0].Konferenzdatum.Year == 1 ? DateTime.Now : l[0].Konferenzdatum));
                                }
                            }
                        }
                    }
                }
            }
            return leistungen;
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

                    leistungen.AddRange((from a in alleGeholtenLeistungenDesSchülers where a.Fach == fach.Name select a).ToList());
                                        
                    f += (leistungen.Count == 0? "": leistungen.FirstOrDefault().Gesamtnote + leistungen.FirstOrDefault().Tendenz).PadRight(3).Substring(0, 3).PadRight(breite) + "|";
                }
            }
            Global.WriteLine(f.PadRight(8));
        }
    }
}