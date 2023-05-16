// Published under the terms of GPLv3 Stefan Bäumer 2023.

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
                            schüler.Vorname = x[2];
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
                             where a.Gesamtnote != null
                             where a.Gesamtnote != ""
                             where a.Klasse.Substring(0, a.Klasse.Length - 1) == schüler.Klasse.Substring(0, schüler.Klasse.Length - 1)
                             select a
                             ).ToList();

                    // Wenn es mehr als einen infragekommenden Webuntis-Unterricht gibt ...

                    if (w.Count > 1)
                    {
                        // ... und sich die Noten unterscheiden ...

                        if ((from ww in w select ww.Gesamtnote + ww.Tendenz).Distinct().Count() > 1)
                        {
                            // ... wird die Liste absteiegend nach Datum sortiert und dadurch die jüngste Leistung ausgewählt.

                            w = (from ww in w.OrderByDescending(x => x.Datum) select ww).ToList();

                            Global.WriteLine("ACHTUNG: " + w[0].Lehrkraft + "," + w[0].Fach + ": Noten widersprechen sich: " + Global.List2String((from ww in w select ww.Gesamtnote + ww.Tendenz).Distinct().ToList(), ',') + ". Nur die jüngste Eintragung (" + w[0].Gesamtnote + w[0].Tendenz + ") wird berücksichtigt.");

                            if ((from ww in w select ww.Datum).Distinct().ToList().Count() == 1)
                            {
                                throw new Exception("Achtung: Beide Unterrichte " + w[0].Lehrkraft + "," + w[0].Fach + " haben dasselbe Datum, aber unterschiedliche Noten.");
                            }
                        }
                    }


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

        internal void GeholteLeistungenBehandeln(string interessierendeKlasse)
        {
            var aktuelleUnterrichte = new List<string>();

            foreach (var schüler in this)
            {
                foreach (var uA in schüler.UnterrichteAktuell)
                {
                    if (!aktuelleUnterrichte.Contains(uA.Fach))
                    {
                        aktuelleUnterrichte.Add(uA.Fach);
                    }
                }
            }

            var geholteUnterrichte = new List<string>();

            foreach (var schüler in this)
            {
                foreach (var uG in schüler.UnterrichteGeholt)
                {
                    if (!geholteUnterrichte.Contains(uG.Fach + "|Konferenzdatum: " + uG.AtlantisLeistung.Konferenzdatum.ToShortDateString()))
                    {
                        geholteUnterrichte.Add(uG.Fach + "|Konferenzdatum: " + uG.AtlantisLeistung.Konferenzdatum.ToShortDateString());
                    }
                }
            }
            Console.WriteLine(" ");
            Console.WriteLine(" ");
            Console.WriteLine("Aktuell werden die Fächer " + Global.List2String(aktuelleUnterrichte, ',') + " unterrichtet." + (geholteUnterrichte.Count == 0 ? "Es gibt keine alten Unterrichte, die geholt werden könnten." : " Folgende Fächer können geholt werden:"));

            List<string> dieseFächerHolen = new List<string>();
            bool gehtNichtWeiter = true;
            int anzahlOriginal = geholteUnterrichte.Count();

            do
            {
                if (geholteUnterrichte.Count > 0)
                {
                    int i = 1;

                    foreach (var fach in geholteUnterrichte)
                    {
                        if (!dieseFächerHolen.Contains(fach))
                        {
                            int anzahl = (from s in this 
                                          from u in s.UnterrichteGeholt 
                                          where u.AtlantisLeistung.Konferenzdatum.ToShortDateString() == fach.Split(' ')[1] 
                                          where u.Fach == fach.Split('|')[0]
                                          select u).Count();


                            Console.WriteLine(" ".PadRight(dieseFächerHolen.Count + 1, ' ') + " " + i + ". " + fach + " Anzahl: " + anzahl + " SuS" );
                            i++;
                        }
                    }

                    var xx = new List<string>();
                    
                    foreach (var fach in geholteUnterrichte)
                    {
                        if (!dieseFächerHolen.Contains(fach))
                        {
                            xx.Add(fach);
                        }
                    }

                    geholteUnterrichte = xx;

                    Console.WriteLine(" ".PadRight(dieseFächerHolen.Count + 1,' ') + " Geben Sie die Ziffer 1, ..., " + (i - 1) + " ein oder drücken Sie ENTER, wenn keine " + (dieseFächerHolen.Count > 0? "weitere ": "") + "Note gezogen werden soll.");

                    string auswahl = Console.ReadKey().Key.ToString();

                    auswahl = auswahl.Substring(auswahl.Length - 1, 1);
                    Console.WriteLine("");

                    try
                    {
                        if (auswahl != "r") // Wenn nicht ENTER gedrückt wurde
                        {
                            var intAuswahl = Convert.ToInt32(auswahl.ToString());
                            dieseFächerHolen.Add(geholteUnterrichte[intAuswahl - 1]); 
                        }
                        if (dieseFächerHolen.Count == anzahlOriginal)
                        {
                            var intAuswahl = Convert.ToInt32(auswahl.ToString());                            
                            gehtNichtWeiter = false;
                        }
                        if (auswahl == "r") // Wenn nicht ENTER gedrückt wurde                              
                        {
                            gehtNichtWeiter = false;
                        }
                                                
                        Console.WriteLine("");
                    }
                    catch
                    {
                    }                    
                }
            } while (gehtNichtWeiter);

            Global.WriteLine(" ".PadRight(dieseFächerHolen.Count + 1, ' ') + "  Diese Fächer werden geholt: ");

            foreach (var fach in dieseFächerHolen)
            {
                Global.WriteLine(" ".PadRight(dieseFächerHolen.Count + 1, ' ') + "   " + fach);
            }

            foreach (var schüler in this)
            {
                schüler.HoleLeistungen(dieseFächerHolen);
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
                var aktuelleFächer = new List<string>();

                foreach (var u in schüler.UnterrichteAktuell)
                {
                    aktuelleFächer.Add(u.Fach);

                    foreach (var atlantisLeistung in (from al in atlantisLeistungen.OrderByDescending(x => x.Konferenzdatum)
                                                      where al.SchlüsselExtern == schüler.SchlüsselExtern
                                                      select al).ToList())
                    {
                        if (atlantisLeistung.FachAliases.Contains(Regex.Replace(u.Fach, @"[\d-]", string.Empty)))
                        {
                            if (atlantisLeistung.Gesamtnote != "")
                            {
                                if (atlantisLeistung.Konferenzdatum >= DateTime.Now.AddDays(-5))
                                {
                                    // ... wird sie zugeordnet.

                                    u.AtlantisLeistung = atlantisLeistung;
                                    u.Reihenfolge = atlantisLeistung.Reihenfolge; // Die Reihenfolge aus Atlantis wird an den Unterricht übergeben                                    
                                    break;
                                }
                            }
                        }
                    }

                    schüler.UnterrichteGeholt = new Unterrichte();
                }

                // Alle Atlantisleistungen, deren Fach nicht zu einem aktuellen Unterricht passt und deren Konferenzdatum in der Vergangenheit liegt, sind geholte Unterrichte

                foreach (var atlantisLeistung in (from al in atlantisLeistungen.OrderByDescending(x => x.Konferenzdatum)
                                                  where al.SchlüsselExtern == schüler.SchlüsselExtern
                                                  where al.Konferenzdatum < DateTime.Now.AddDays(-30)
                                                  where al.Gesamtnote != null
                                                  where al.Gesamtnote != ""
                                                  select al).ToList())
                {
                    bool geholt = true;
                    foreach (var atlantisFach in atlantisLeistung.FachAliases)
                    {
                        if (aktuelleFächer.Contains(atlantisFach))
                        {
                            geholt = false;
                            break;
                        }

                        // Wenn das Fach bereits geholt wurde, wird es nicht nochmal geholt

                        if ((from u in schüler.UnterrichteGeholt where u.Fach == atlantisFach select u).Any())
                        {
                            geholt = false;
                            break;
                        }
                    }                    
                    if (geholt)
                    {
                        schüler.UnterrichteGeholt.Add(new Unterricht(atlantisLeistung));
                    }
                }
            }
        }

        internal void Update(string interessierendeKlasse)
        {
            int outputIndex = Global.SqlZeilen.Count();

            int i = 0;

            Global.WriteLine((" "));
            Global.WriteLine((" "));
            Global.WriteLine(("Fehlende Einträge der Klasse " + interessierendeKlasse + " in Atlantis: "));
            Global.WriteLine(("=============================================== ").PadRight(interessierendeKlasse.Length,'='));
            Global.WriteLine((" "));

            foreach (var schüler in this)
            {
                var fehlendeNoten = new List<string>();

                foreach (var unterricht in schüler.UnterrichteAktuell)
                {
                    if (unterricht.AtlantisLeistung != null)
                    {
                        if (unterricht.WebuntisLeistung != null)
                        {
                            unterricht.QueryBauen();

                            if (unterricht.WebuntisLeistung.Beschreibung != null && unterricht.WebuntisLeistung.Query != null)
                            {
                                unterricht.WebuntisLeistung.Update();
                                i++;
                            }
                        }
                        else
                        {
                            fehlendeNoten.Add(unterricht.Fach + "(" + unterricht.Lehrkraft + ")");
                        }
                    }
                }
                
                foreach (var unterricht in schüler.UnterrichteGeholt)
                {
                    if (unterricht.AtlantisLeistung != null)
                    {
                        unterricht.WebuntisLeistung = new Leistung(unterricht.AtlantisLeistung);
                        unterricht.AtlantisLeistung.Gesamtpunkte = "";
                        unterricht.AtlantisLeistung.Gesamtnote = "";
                        unterricht.AtlantisLeistung.Tendenz = "";
                        unterricht.QueryBauen();                        
                        unterricht.WebuntisLeistung.Update();
                    }
                }

                if (fehlendeNoten.Count > 0)
                {
                    Global.WriteLine(" " + schüler.Nachname.PadRight(7).Substring(0, 7) + " " + schüler.Vorname.PadRight(2).Substring(0, 2) + ": Es fehlen noch Noten: " + Global.List2String(fehlendeNoten, ','));
                }
            }

            if (i == 0)
            {
                Global.WriteLine("A C H T U N G: Es wurde keine einzige Leistung angelegt oder verändert. Das kann daran liegen, dass das Notenblatt in Atlantis nicht richtig angelegt ist. Die Tabelle zum manuellen eintragen der Noten muss in der Notensammelerfassung sichtbar sein.");
            }
        }

        internal bool WidersprechendeGesamtnotenImSelbenFachKorrigieren(string interessierendeKlasse)
        {
            var aktuelleFächer = (from s in this from u in s.UnterrichteAktuell select u.Fach + "|" + u.Lehrkraft).Distinct().ToList();

            foreach (var af in aktuelleFächer)
            {
                if ((from aa in aktuelleFächer where Regex.Replace(aa.Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) select aa.Split('|')[0]).Count() > 1)
                {
                    foreach (var schüler in this)
                    {
                        var doppelteNoten = new List<string>();
                        var doppelterLehrer = new List<string>();

                        foreach (var unterricht in (from u in schüler.UnterrichteAktuell where Regex.Replace(u.Fach, @"[\d-]", string.Empty) == Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) select u).ToList())
                        {
                            if (unterricht.WebuntisLeistung != null && unterricht.WebuntisLeistung.Gesamtnote != null && unterricht.WebuntisLeistung.Gesamtnote != "")
                            {
                                if (!doppelteNoten.Contains(unterricht.WebuntisLeistung.Gesamtnote + unterricht.WebuntisLeistung.Tendenz))
                                {
                                    doppelteNoten.Add(unterricht.WebuntisLeistung.Gesamtnote + unterricht.WebuntisLeistung.Tendenz);
                                    doppelterLehrer.Add(unterricht.Lehrkraft);
                                }
                            }
                        }

                        if (doppelteNoten.Count > 1)
                        {
                            Global.WriteLine("Im Fach " + Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) + "(" + Global.List2String(doppelterLehrer, ',') + ") gibt es widersprechende Noten (" + Global.List2String(doppelteNoten, ',') + "). Welcher Lehrer soll gezogen werden?");
                            int i = 0;

                            foreach (var aF in (from a in aktuelleFächer 
                                                where doppelterLehrer.Contains(a.Split('|')[1]) 
                                                where Regex.Replace(a.Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty)
                                                select a).ToList())
                            {
                                i++;
                                Global.WriteLine(" " + i.ToString().PadLeft(2) + "." + aF);
                            }
                            bool gehtNichtWeiter = true;

                            do
                            {
                                Global.Write(" Wessen Noten (1, ..., " + i + ") sollen im Zeugnis erscheinen? ");
                                ConsoleKeyInfo input;
                                input = Console.ReadKey(true);

                                try
                                {
                                    var auswahl = Char.ToUpper(input.KeyChar);

                                    var intAuswahl = Convert.ToInt32(auswahl.ToString());

                                    var lehrkraft = ((from a in aktuelleFächer where doppelterLehrer.Contains(a.Split('|')[1]) where Regex.Replace(a.Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) select a).ToList()[intAuswahl - 1]).Split('|')[1];

                                    if (intAuswahl <= i)
                                    {
                                        // Alle Webuntisleistungen der Unterrichte anderer LuL in diesem Fach werden gelöscht.

                                        foreach (var s in this)
                                        {
                                            foreach (var unt in (from u in s.UnterrichteAktuell where Regex.Replace(u.Fach.Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) select u).ToList())
                                            {
                                                if (unt.Lehrkraft != lehrkraft)
                                                {
                                                    unt.WebuntisLeistung = null;                                                    
                                                    unt.AtlantisLeistung = null;
                                                    unt.Bemerkung = lehrkraft + " setzt die Note in " + unt.Fach + ".";
                                                }
                                            }
                                        }
                                    }
                                    gehtNichtWeiter = false;
                                }
                                catch (Exception)
                                {
                                }
                                Global.WriteLine("");
                            } while (gehtNichtWeiter);

                            // Die Tabelle wird aus den SQL-Zeilen wieder gelöscht, damit nur die letzte Tabelle angezeigt wird.

                            int index = Global.SqlZeilen.IndexOf("/*    Notenliste                                                                                                                                                                     */");

                            for (int ii = Global.SqlZeilen.Count - 1; ii >= index; ii--)
                            {
                                Global.SqlZeilen.RemoveAt(ii);
                            }

                            return true;
                        }
                    }
                }
            }

            return false;
        }

        internal void TabelleErzeugen(string interessierendeKlasse)
        {
            do
            {
                Global.WriteLine((" "));
                Global.WriteLine((" "));
                Global.WriteLine(("    Leistungen der Klasse " + interessierendeKlasse + " in Atlantis: "));
                Global.WriteLine(("    ======================================== ").PadRight(interessierendeKlasse.Length, '='));
                Global.WriteLine((" "));
                Global.WriteLine("*-----------------------------------------".PadRight(Global.PadRight + 3, '-') + "*");
                Global.WriteLine(("|Name |SuS-Id| Noten + Tendenzen der Klasse " + interessierendeKlasse + " aus Webuntis:").PadRight(Global.PadRight + 3) + "|");
                Global.WriteLine("*------------+----------------------------".PadRight(Global.PadRight + 3, '-') + "*");

                var aktuelleFächer = (from s in this from u in s.UnterrichteAktuell.OrderBy(x => x.Reihenfolge) select u.Fach + "|" + interessierendeKlasse + "|" + u.Lehrkraft + "|" + u.Gruppe + "|" + u.LessonNumber).Distinct().ToList();
                
                bool xxxx = false;
                bool reliabwähler = false;

                string f = "|            |";
                string l = "|            |";
                string g = "|            |";
                string w = "|            |";
                string k = "|            |";
                string n = "| Unterr.Nr.:|";

                foreach (var aF in aktuelleFächer)
                {
                    f += aF.Split('|')[0].PadRight(5).Substring(0, 5) + "|";
                    l += aF.Split('|')[2].PadRight(5).Substring(0, 5) + "|";
                    g += aF.Split('|')[3] == "" ? "Alle |" : "Kurs |";
                    w += "W  A |";
                    k += "      ";
                    n += aF.Split('|')[4].PadRight(4).Substring(0, 4) + " |";
                }

                Global.WriteLine(f);
                Global.WriteLine(l);
                Global.WriteLine(g);
                Global.WriteLine(n);
                //Global.WriteLine(k.Substring(0,k.Length - 1) + "|");

                Global.WriteLine(w);
                Global.WriteLine("*------------+-----+-----+-----+-----+-----+-----+-----+-----+-----+-----+-----+----".PadRight(Global.PadRight + 3, '-') + "*");

                foreach (var schüler in this)
                {
                    string s = "|" + schüler.Nachname.PadRight(5).Substring(0, 5) + "|" + schüler.SchlüsselExtern + "|";

                    foreach (var aF in aktuelleFächer)
                    {
                        var uA = (from u in schüler.UnterrichteAktuell
                                 where u.Fach == aF.Split('|')[0]
                                 where u.Lehrkraft == aF.Split('|')[2]
                                 where aF.Split('|')[3] == u.Gruppe
                                 select u).FirstOrDefault();

                        // Wenn es zu dem Fach in der Schülergruppe und bei der Lehrkraft einen Unterricht gibt ... 

                        if (uA != null)
                        {
                            if (uA.AtlantisLeistung != null)
                            {
                                if (uA.WebuntisLeistung != null)
                                {
                                    s += (uA.WebuntisLeistung.Gesamtnote + uA.WebuntisLeistung.Tendenz).PadRight(2) + " " + (uA.AtlantisLeistung.Gesamtnote + uA.AtlantisLeistung.Tendenz).PadRight(2) + "|";
                                }
                                else
                                {
                                    s += "   " + (uA.AtlantisLeistung.Gesamtnote + uA.AtlantisLeistung.Tendenz).PadRight(2) + "|";
                                }
                            }
                            else
                            {
                                if (uA.WebuntisLeistung != null)
                                {
                                    // Wenn der Schüler in Webuntis steht, aber kein Leistungsdatensatz in Atlantis hat
                                    s += (uA.WebuntisLeistung.Gesamtnote + uA.WebuntisLeistung.Tendenz).PadRight(2) + " XX|";
                                    xxxx = true;
                                }
                                else
                                {
                                    if (uA.Bemerkung.Contains("setzt die Note"))
                                    {
                                        s += uA.Bemerkung.Substring(0,3) + "XX|";
                                    }
                                    else
                                    {
                                        s += "XX XX|";
                                    }
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
                                s += "XX XX|";
                                xxxx = true;
                            }
                        }
                    }

                    Global.WriteLine(s);
                }

                Global.WriteLine("*-----+------+-----+-----+-----+-----+-----+-----+-----+-----+-----+-----+-----+----".PadRight(Global.PadRight + 3, '-') + "*");

                if (reliabwähler)
                {
                    Global.WriteLine("  * R* Reliabwähler? Evtl. Note in Konferenz geben. Bei Abschluss/Abgang Strich.");
                }
                if (xxxx)
                {
                    Global.WriteLine("  * X Der Schüler hat den Kurs nicht belegt oder es liegt keinen Datensatz vor.");
                }

                //Global.WriteLine("  * Gibt es Fremdsprachenprüfungen anstelle von Englisch?");
                //Global.WriteLine("  * Teilen sich KuK ein Fach? Note untereinander abgestimmt?");
                Global.WriteLine("");

            } while (WidersprechendeGesamtnotenImSelbenFachKorrigieren(interessierendeKlasse)); 
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