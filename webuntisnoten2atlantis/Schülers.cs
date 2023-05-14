﻿// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace webuntisnoten2atlantis
{
    public class Schülers : List<Schüler>
    {
        public Schülers()
        {
        }

        public Schülers(string sourceExportLessons)
        {
            using (StreamReader reader = new StreamReader(sourceExportLessons))
            {
                var überschrift = reader.ReadLine();
                int i = 1;

                while (true)
                {
                    i++;
                    Schüler schüler = new Schüler();
                    schüler.UnterrichteAktuell = new Unterrichte();
                    schüler.UnterrichteGeholt = new Unterrichte();

                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            var x = line.Split('\t');

                            schüler = new Schüler();
                            schüler.StudentZeile = i;
                            schüler.Nachname = x[1];
                            schüler.Vorname = x[1];
                            schüler.Klasse = x[5];

                            try
                            {
                                schüler.EintrittInKlasse = DateTime.ParseExact(x[6], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            }
                            catch (Exception)
                            {
                            }
                            try
                            {
                                schüler.AustrittAusKlasse = DateTime.ParseExact(x[7], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            }
                            catch (Exception)
                            {
                            }
                            try
                            {
                                schüler.SchlüsselExtern = Convert.ToInt32(x[10]);
                            }
                            catch (Exception)
                            {
                            }

                            // Nur Schüler, die nicht längerer als 30 Tage ausgetreten sind.

                            if (schüler.Klasse != "" && (schüler.AustrittAusKlasse.Year == 1 || DateTime.Now.AddDays(-30) < schüler.AustrittAusKlasse))
                            {
                                this.Add(schüler);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string message = ex.Message + "Fehler in Zeile: " + i;

                        Console.WriteLine(message);
                        Console.ReadKey();
                    }

                    if (line == null)
                    {
                        break;
                    }
                }

                Global.WriteLine(("Alle aktiven Schüler, die nicht länger als 30 Tage ausgetreten sind aus Webuntis ").PadRight(Global.PadRight - 2, '.') + this.Count.ToString().PadLeft(6));
            }
        }

        public Schülers GetMöglicheSchülerDerKlasse(string interessierendeKlasse)
        {
            Schülers alleWebuntisSchülersDieserKlasse = new Schülers();
            alleWebuntisSchülersDieserKlasse.AddRange((from m in this where m.Klasse == interessierendeKlasse select m).ToList());
            Global.WriteLine(("Verschiedene SuS der Klasse " + interessierendeKlasse + " aus Webuntis ").PadRight(Global.PadRight - 2, '.') + alleWebuntisSchülersDieserKlasse.Count.ToString().PadLeft(6));
            return alleWebuntisSchülersDieserKlasse;
        }

        internal void GetIntessierendeWebuntisLeistungen(Leistungen alleWebuntisLeistungen)
        {
            foreach (var schüler in this)
            {
                foreach (var unterricht in schüler.UnterrichteAktuell)
                {                    
                    var w = (from a in alleWebuntisLeistungen
                             where a.FachAliases.Contains(unterricht.Fach)
                             where a.Lehrkraft == unterricht.Lehrkraft
                             where a.SchlüsselExtern == schüler.SchlüsselExtern                             
                             where a.Klasse.Substring(0, a.Klasse.Length - 1) == schüler.Klasse.Substring(0,schüler.Klasse.Length -1)
                             select a
                             ).ToList();

                    if (w.Count > 0)
                    {
                        unterricht.WebuntisLeistung = new Leistung(
                            w[0].Name, 
                            w[0].Fach, 
                            w[0].FachAliases, 
                            w[0].Gesamtnote,
                            w[0].Gesamtpunkte, 
                            w[0].Tendenz,
                            w[0].Datum, 
                            w[0].Nachname, 
                            w[0].SchlüsselExtern);                        
                    }
                }
            }
        }

        internal void LinkZumTeamsChatErzeugen(Lehrers alleAtlantisLehrer, string interessierendeKlasse)
        {
            var url = "https://teams.microsoft.com/l/chat/0/0?users=";

            foreach (var schüler in this)
            {
                foreach (var unterricht in schüler.UnterrichteAktuell)
                {
                    var mail = (from l in alleAtlantisLehrer where l.Kuerzel == unterricht.Lehrkraft select l.Mail).FirstOrDefault();

                    if (mail != null && mail != "" && !url.Contains(mail))
                    {
                        url += mail + ",";
                    }
                }
            }
            Global.WriteLine("Link zum Teams-Chat mit den LuL der Klasse " + interessierendeKlasse + ":");
            Global.WriteLine(" " + url.TrimEnd(','));
            Global.WriteLine("  ");
        }

        internal void GetAtlantisLeistungen(string connectionStringAtlantis, List<string> aktSj, string user, string interessierendeKlasse, string hzJz)
        {
            // Alle AtlantisLeistungen dieser Klasse und der Parallelklassen in allen Jahren

            var atlantisLeistungen = new Leistungen(connectionStringAtlantis, aktSj, user, interessierendeKlasse, this, hzJz);

            foreach (var schüler in this)
            {
                schüler.UnterrichteGeholt = new Unterrichte();

                // Wenn die Atlantisleistung zu einem aktuellen Fach in dieser oder der Parallelklasse passt des Schülers passt, ...

                bool zugeordnet = false;

                foreach (var u in schüler.UnterrichteAktuell)
                {
                    var atlantisLeistungGeholt = new Leistung();

                    foreach (var atlantisLeistung in (from al in atlantisLeistungen.OrderByDescending(x => x.Konferenzdatum)
                                                      where al.SchlüsselExtern == schüler.SchlüsselExtern
                                                      select al).ToList())
                    {
                        atlantisLeistungGeholt = atlantisLeistung;

                        if (atlantisLeistung.FachAliases.Contains(u.Fach))
                        {
                            if (atlantisLeistung.Gesamtnote != "")
                            {
                                if (atlantisLeistung.Konferenzdatum >= DateTime.Now.AddDays(-5))
                                {
                                    // ... wird sie zugeordnet.

                                    u.AtlantisLeistung = atlantisLeistung;
                                    u.Reihenfolge = atlantisLeistung.Reihenfolge; // Die Reihenfolge aus Atlantis wird an den Unterricht übergeben
                                    zugeordnet = true;
                                    break;
                                }
                            }
                        }
                    }

                    // Wenn die Atlantisleistung von diesem Schüler in dieser Klasse erbracht wurde, aber 
                    // trotzdem nicht zugeordnet werden kann, dann wird ein geholter Unterricht angelegt und
                    // die Leistung zugeordnet:

                    if (!zugeordnet)
                    {
                        // Unterricht vergangener Abschnitte
                        
                        schüler.UnterrichteGeholt.Add(new Unterricht(atlantisLeistungGeholt));
                    }
                }
            }
        }

        internal void WidersprechendeGesamtnotenImSelbenFachKorrigieren(string interessierendeKlasse)
        {
            var aktuelleFächer = (from s in this from u in s.UnterrichteAktuell select u.Fach + "|" + u.Lehrkraft).ToList();
            var doppelteFächer = new List<string>();

            foreach (var af in aktuelleFächer)
            {
                var sfOhneZiffer = Regex.Replace(af.Split(',')[0], @"[\d-]", string.Empty);

                if (aktuelleFächer.Contains(sfOhneZiffer))
                {
                    if (!doppelteFächer.Contains(sfOhneZiffer))
                    {
                        doppelteFächer.Add(sfOhneZiffer);
                    }
                }
            }

            // Wenn es doppelte Fächer gibt, wird auf konkurrierende Noten geprüft

            if (doppelteFächer.Count > 0)
            {
                foreach (var doppeltesFach in doppelteFächer)
                {
                    foreach (var schüler in this)
                    {
                        var doppelteNoten = new List<string>();
                        var doppelterLehrer = new List<string>();

                        foreach (var unterricht in schüler.UnterrichteAktuell)
                        {
                            if (Regex.Replace(unterricht.Fach, @"[\d-]", string.Empty) == doppeltesFach)
                            {
                                if (unterricht.WebuntisLeistung.Gesamtnote != null && unterricht.WebuntisLeistung.Gesamtnote != "")
                                {
                                    if (!doppelteNoten.Contains(unterricht.WebuntisLeistung.Gesamtnote + unterricht.WebuntisLeistung.Tendenz))
                                    {
                                        doppelteNoten.Add(unterricht.WebuntisLeistung.Gesamtnote + unterricht.WebuntisLeistung.Tendenz);
                                        doppelterLehrer.Add(unterricht.WebuntisLeistung.Lehrkraft);
                                    }
                                }
                            }
                        }

                        if (doppelteNoten.Count > 1)
                        {
                            Global.WriteLine("Im Fach " + doppeltesFach + "(" + Global.List2String(doppelterLehrer,',') + ") gibt es widersprechende Noten. Welche Welche Note soll gezoegen werden?");
                            Console.ReadKey();
                        }
                    }                    
                }
            }
        }

        internal void TabelleErzeugen(string interessierendeKlasse)
        {

            Global.WriteLine("*-----------------------------------------".PadRight(Global.PadRight + 3, '-') + "*");
            Global.WriteLine(("|Name |SuS-Id| Noten + Tendenzen der Klasse " + interessierendeKlasse + " aus Webuntis:").PadRight(Global.PadRight + 3) + "|");
            Global.WriteLine("*------------+----------------------------".PadRight(Global.PadRight + 3, '-') + "*");

            var aktuelleFächer = (from s in this from u in s.UnterrichteAktuell.OrderBy(x=>x.Reihenfolge) select u.Fach + "|" + interessierendeKlasse + "|" + u.Lehrkraft + "|" + u.Gruppe).Distinct().ToList();
            var geholteFächer = (from s in this from u in s.UnterrichteGeholt select u.Fach + "|" + interessierendeKlasse + "|" + u.Lehrkraft + "|" + u.Gruppe).Distinct().ToList();

            bool nichtMehrSchüler = false;
            bool reliabwähler = false;
            bool kursabwähler = false;

            string f = "|            |";
            string l = "|            |";
            string g = "|            |";
            string w = "|            |";
            string k = "|            |";

            foreach (var aF in aktuelleFächer)
            {
                f += aF.Split('|')[0].PadRight(5).Substring(0, 5) + "|";
                l += aF.Split('|')[2].PadRight(5).Substring(0, 5) + "|";
                g += aF.Split('|')[3] == "" ? "Alle |" : "Kurs |";
                w += "W  A |";
                k += "      ";
            }

            Global.WriteLine(f);
            Global.WriteLine(l);
            Global.WriteLine(g);
            Global.WriteLine(k.Substring(0,k.Length - 1) + "|");
            
            Global.WriteLine(w);
            Global.WriteLine("*------------+-----+-----+-----+-----+-----+-----+-----+-----+-----+-----+-----+----".PadRight(Global.PadRight + 3, '-') + "*");

            foreach (var schüler in this)
            {
                string s = "|" + schüler.Nachname.PadRight(5).Substring(0, 5) + "|" + schüler.SchlüsselExtern + "|";

                foreach (var aF in aktuelleFächer)
                {
                    var x = (from u in schüler.UnterrichteAktuell 
                             where u.Fach == aF.Split('|')[0] 
                             where u.Lehrkraft == aF.Split('|')[2] 
                             where aF.Split('|')[3] == u.Gruppe 
                             select u).FirstOrDefault();

                    // Wenn es zu dem Fach in der Schülergruppe und bei der Lehrkraft einen Unterricht gibt ... 

                    if (x != null)
                    {                        
                        if (x.AtlantisLeistung != null)
                        {
                            if (x.WebuntisLeistung != null)
                            {
                                s += (x.WebuntisLeistung.Gesamtnote + x.WebuntisLeistung.Tendenz).PadRight(2) + " " + (x.AtlantisLeistung.Gesamtnote + x.AtlantisLeistung.Tendenz).PadRight(2) + "|";
                            }
                            else
                            {
                                s += "   " + (x.AtlantisLeistung.Gesamtnote + x.AtlantisLeistung.Tendenz).PadRight(2) + "|";
                            }
                        }
                        else
                        {
                            if (x.WebuntisLeistung != null)
                            {
                                // Wenn der Schüler in Webuntis steht, aber kein Leistungsdatensatz in Atlantis hat
                                s += (x.WebuntisLeistung.Gesamtnote + x.WebuntisLeistung.Tendenz).PadRight(2) + " ? |";
                                nichtMehrSchüler = true;
                            }
                            else
                            {
                                s += "?  ? |";
                            }
                        }
                    }
                    else
                    {
                        // Wenn der Schüler in Webuntis dieses Fach nicht belegt hat, wird geixxt

                        if ((new List<string> { "KR", "ER", "REL" }).Contains(aF.Split('|')[0]))
                        {
                            s += "R*   |";
                            reliabwähler = true;
                        }
                        else
                        {
                            s += "X  X |";
                            kursabwähler = true;
                        }                        
                    }
                }

                Global.WriteLine(s);
            }

            Global.WriteLine("*-----+------+-----+-----+-----+-----+-----+-----+-----+-----+-----+-----+-----+----".PadRight(Global.PadRight + 3, '-') + "*");
            
            if (nichtMehrSchüler)
            {
                Global.WriteLine("  * ? Ist der Schüler evtl. schon in F***lantis ausgeschult, aber noch in Webuntis existent?");
            }
            if (reliabwähler)
            {
                Global.WriteLine("  * R* Reliabwähler? Evtl. Note in Konferenz geben. Bei Abschluss/Abgang Strich.");
            }
            if (kursabwähler)
            {
                Global.WriteLine("  * XX Der Schüler hat den Kurs nicht belegt.");
            }
            
            Global.WriteLine("  * Gibt es Fremdsprachenprüfungen anstelle von Englisch?");
            Global.WriteLine("  * Teilen sich KuK ein Fach? Note untereinander abgestimmt?");
            Global.WriteLine("");
        }

        internal void GetUnterrichteAktuell(Unterrichte alleUnterrichte, Gruppen alleGruppen, string interessierendeKlasse)
        {
            var unterrichteDerKlasse = (from a in alleUnterrichte 
                                        where a.Klassen.Contains(interessierendeKlasse) 
                                        where a.Startdate <= DateTime.Now
                                        where a.Enddate >= DateTime.Now.AddMonths(-2) // Unterrichte, die 2 Monat vor Konferenz beendet wurden, zählen
                                        select a).ToList();

            foreach (var schüler in this)
            {
                schüler.GetUnterrichte(unterrichteDerKlasse, alleGruppen);
            }
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

                    try
                    {
                        if (item != "" && möglicheSuS.Contains(Convert.ToInt32(item)))
                        {
                            // Falls ja, dann wird der Vorschlag aus den Properties übernommen.

                            vorschlag += item + ",";
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
                if (vorschlag.TrimEnd(',') == "")
                {
                    vorschlag = "alle";
                }
            }
            Properties.Settings.Default.InteressierendeSuS = vorschlag;
            Properties.Settings.Default.Save();
            return vorschlag.TrimEnd(',');
        }
    }
}