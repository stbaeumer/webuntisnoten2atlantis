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
                    schüler.UnterrichteAusWebuntis = new Unterrichte();
                    schüler.UnterrichteGeholt = new Unterrichte();
                    schüler.UnterrichteAktuellAusAtlantis = new Unterrichte();

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

                Console.WriteLine(("Alle aktiven Schüler, die nicht länger als 30 Tage ausgetreten sind aus Webuntis ").PadRight(Global.PadRight - 2, '.') + this.Count.ToString().PadLeft(6));
            }
        }

        public Schülers GetMöglicheSchülerDerKlasse(string interessierendeKlasse)
        {
            Schülers alleWebuntisSchülersDieserKlasse = new Schülers();
            alleWebuntisSchülersDieserKlasse.AddRange((from m in this where m.Klasse == interessierendeKlasse select m).ToList());
            Console.WriteLine(("Verschiedene SuS der Klasse " + interessierendeKlasse + " aus Webuntis ").PadRight(Global.PadRight - 2, '.') + alleWebuntisSchülersDieserKlasse.Count.ToString().PadLeft(6));
            return alleWebuntisSchülersDieserKlasse;
        }

        internal void GetWebuntisLeistungen(Leistungen alleWebuntisLeistungen)
        {
            foreach (var schüler in this)
            {
                foreach (var unterricht in schüler.UnterrichteAusWebuntis)
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
                            // ... wird die Liste absteigend nach Datum sortiert und dadurch die jüngste Leistung ausgewählt.

                            w = (from ww in w.OrderByDescending(x => x.Datum) select ww).ToList();

                            if (!Global.Rückmeldung.Contains(" * " + w[0].Lehrkraft.PadRight(3) + "," + w[0].Fach.PadRight(4) + ": Die Noten in Webuntis (" + w[0].Name + ") sind nicht eindeutig. Ist Note (" + w[0].Gesamtnote + w[0].Tendenz + ") korrekt? Evtl. in der Konferenz ändern lassen."))
                            {
                                Global.Rückmeldung.Add(" * " + w[0].Lehrkraft.PadRight(3) + "," + w[0].Fach.PadRight(4) + ": Die Noten in Webuntis (" + w[0].Name + ") sind nicht eindeutig. Ist Note (" + w[0].Gesamtnote + w[0].Tendenz + ") korrekt? Evtl. in der Konferenz ändern lassen.");
                            }
                    
                            if ((from ww in w select ww.Datum).Distinct().ToList().Count() == 1)
                            {
                                throw new Exception("Achtung: Beide Unterrichte " + w[0].Lehrkraft + "," + w[0].Fach + " haben dasselbe Datum, aber unterschiedliche Noten.");
                            }
                        }
                    }

                    if (w.Count > 0)
                    {
                        unterricht.WL = new Leistung(
                            w[0].Name,
                            w[0].Fach,
                            w[0].FachAliases,
                            w[0].Gesamtnote,
                            w[0].Gesamtpunkte,
                            w[0].Tendenz,
                            w[0].Datum,
                            w[0].Nachname,
                            w[0].Lehrkraft,
                            w[0].SchlüsselExtern,
                            w[0].MarksPerLessonZeile);

                        if (w[0].Gesamtnote == "-" && !Global.Rückmeldung.Contains(" * " + w[0].Lehrkraft.PadRight(3) + "," + w[0].Fach.PadRight(4) + ": Bei einem Strich (" + w[0].Name + ") muss rechtzeitig vor der Konferenz eine Bemerkung mit Ulla vereinbart werden."))
                        {
                            Global.Rückmeldung.Add(" * " + w[0].Lehrkraft.PadRight(3) + "," + w[0].Fach.PadRight(4) + ": Bei einem Strich (" + w[0].Name + ") muss rechtzeitig vor der Konferenz eine Bemerkung mit Ulla vereinbart werden.");
                        }
                    }
                }
            }
        }

        internal void GeholteLeistungenBehandeln(string interessierendeKlasse)
        {
            var aktuelleUnterrichte = new List<string>();

            foreach (var schüler in this)
            {
                foreach (var uA in schüler.UnterrichteAusWebuntis)
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
                    if (!geholteUnterrichte.Contains(uG.Fach + "|Konferenzdatum: " + uG.AL.Konferenzdatum.ToShortDateString()))
                    {
                        geholteUnterrichte.Add(uG.Fach + "|Konferenzdatum: " + uG.AL.Konferenzdatum.ToShortDateString());
                    }
                }
            }
            Console.WriteLine(" ");
            Console.WriteLine(" ");
            Console.WriteLine("Aktuell werden die Fächer " + Global.List2String(aktuelleUnterrichte, ",") + " unterrichtet." + (geholteUnterrichte.Count == 0 ? "Es gibt keine alten Unterrichte, die geholt werden könnten." : " Folgende Fächer können geholt werden:"));

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
                                          where u.AL.Konferenzdatum.ToShortDateString() == fach.Split(' ')[1] 
                                          where u.Fach == fach.Split('|')[0]
                                          select u).Count();

                            var klassen = (from s in this
                                          from u in s.UnterrichteGeholt
                                          where u.AL.Konferenzdatum.ToShortDateString() == fach.Split(' ')[1]
                                          where u.Fach == fach.Split('|')[0]
                                          select u.Klassen).Distinct().ToList();

                            var sus = (from s in this
                                           from u in s.UnterrichteGeholt
                                           where u.AL.Konferenzdatum.ToShortDateString() == fach.Split(' ')[1]
                                           where u.Fach == fach.Split('|')[0]
                                           select u.AL.SchlüsselExtern).Distinct().ToList();


                            Global.WriteLine(" ".PadRight(dieseFächerHolen.Count + 1, ' ') + " " + i + ". " + fach.Split('|')[0].PadRight(5) + "| " + fach.Split('|')[1].PadRight(27) + "| Anzahl: " + anzahl.ToString().PadLeft(2) + " SuS | Klassen: " + Global.List2String(klassen, ",") + (sus.Count < 3 ? "| SuS: " + Global.List2String(sus, ',') : ""));
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
                else
                {
                    gehtNichtWeiter = false;
                }
            } while (gehtNichtWeiter);

            if (dieseFächerHolen.Count > 0)
            {
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
        }

        internal void ChatErzeugen(Lehrers alleAtlantisLehrer, string interessierendeKlasse, string hzJz)
        {
            var url = "https://teams.microsoft.com/l/chat/0/0?users=";

            foreach (var schüler in this)
            {
                foreach (var unterricht in schüler.UnterrichteAusWebuntis)
                {
                    var mail = (from l in alleAtlantisLehrer where l.Kuerzel == unterricht.Lehrkraft select l.Mail).FirstOrDefault();

                    if (mail != null && mail != "" && !url.Contains(mail))
                    {
                        url += mail + ",";
                    }
                }
            }
            Global.WriteLine("  ");
            Global.WriteLine("Link zum Teams-Chat mit den LuL der Klasse " + interessierendeKlasse + ":");
            Global.WriteLine(" " + url.TrimEnd(','));

            var rückmeldung = "/*" +
                "\nHallo LuL der " + interessierendeKlasse + ", " +
                "\n\nin Vorbereitung auf die " + (hzJz == "JZ" ? "Jahreszeugniskonferenzen (https://wiki.berufskolleg-borken.de/doku.php?id=jahreszeugniskonferenzen)" : "Zeugniskonferenzen (https://wiki.berufskolleg-borken.de/doku.php?id=halbjahreszeugniskonferenzen)") + " sende ich eine Übersicht über alle Noten.";

            if (Global.Rückmeldung.Count > 0)
            {
                rückmeldung += "\n\nIch bitte um Beachtung:\n\n";
                                
                rückmeldung += Global.List2String(Global.Rückmeldung, "\n");                
            }

            rückmeldung += "\n\nKennwort: https://wiki.berufskolleg-borken.de/doku.php?id=schulleitungsmitteilungen:schulleitungsmitteilung-2023-04-28#verschluesselung\n";

            rückmeldung += "\n\nGruß,\nStefan\n\n */";

            Global.SqlZeilen.Insert(Global.SqlZeilen.Count(), rückmeldung );
            
            Fächer fehlendeFächer = new Fächer();

            foreach (var schüler in this)
            {
                foreach (var unterricht in schüler.UnterrichteAusWebuntis)
                {
                    if (unterricht.AL != null)
                    {
                        if (unterricht.WL == null)
                        {                            
                            // Die Note wird nur dann als fehlend moniert, wenn auch im selben Fach anderweitig keine Note vergeben wurde. 

                            if (!(from u in schüler.UnterrichteAusWebuntis
                                  where u.WL != null
                                  where Regex.Replace(u.WL.Fach, @"[\d-]", string.Empty) == Regex.Replace(unterricht.Fach, @"[\d-]", string.Empty)
                                  where u.WL.Gesamtnote != null
                                  where u.WL.Gesamtnote != ""
                                  select u).Any())
                            {
                                fehlendeFächer.Add(new Fach(unterricht.Fach, unterricht.Lehrkraft));
                            }
                        }
                    }
                }

                foreach (var unterricht in schüler.UnterrichteGeholt)
                {
                    if (unterricht.AL != null)
                    {
                        unterricht.WL = new Leistung(unterricht.AL);
                        unterricht.AL.Gesamtpunkte = "";
                        unterricht.AL.Gesamtnote = "";
                        unterricht.AL.Tendenz = "";
                        unterricht.QueryBauen();
                        unterricht.WL.Update();
                    }
                }
            }
            
            if (fehlendeFächer.Count > 0)
            {
                foreach (var fach in (from f in fehlendeFächer select f.Name).Distinct().ToList())
                {
                    var lehrer = (from f in fehlendeFächer where f.Name == fach select f.Lehrkraft).FirstOrDefault();
                    var anzahl = (from f in fehlendeFächer where f.Name == fach select f).Count();

                    Global.Rückmeldung.Add(lehrer.PadRight(3) + "," + fach.PadRight(3) + ": Kann es sein, dass " + anzahl.ToString().PadLeft(2) + " Note" + (anzahl > 1 ? "n" : "") + " fehlen?");
                }
            }
        }

        internal void GetAtlantisLeistungen(string connectionStringAtlantis, List<string> aktSj, string user, string interessierendeKlasse, string hzJz)
        {
            // Alle AtlantisLeistungen dieser Klasse und der Parallelklassen in allen Jahren

            var atlantisLeistungen = new Leistungen(connectionStringAtlantis, aktSj, user, interessierendeKlasse, this, hzJz);

            foreach (var schüler in this)
            {
                var aktuelleFächer = new List<string>();

                foreach (var u in schüler.UnterrichteAusWebuntis)
                {
                    aktuelleFächer.Add(u.Fach);

                    foreach (var atlantisLeistung in (from al in atlantisLeistungen.OrderByDescending(x => x.Konferenzdatum)
                                                      where al.SchlüsselExtern == schüler.SchlüsselExtern
                                                      select al).ToList())
                    {
                        if (atlantisLeistung.FachAliases.Contains(Regex.Replace(u.Fach, @"[\d-]", string.Empty)))
                        {
                            // Atlantisleistungen der Vergangenheit werden nicht mehr verändert. Leistungen ohne Konferenzdatum haben das Jahr 1.

                            if (atlantisLeistung.Konferenzdatum.Date >= DateTime.Now.Date || atlantisLeistung.Konferenzdatum.Year == 1)
                            {
                                if (hzJz == atlantisLeistung.HzJz)
                                {
                                    if (aktSj[0] + "/" + aktSj[1] == atlantisLeistung.Schuljahr)
                                    {
                                        // ... wird sie zugeordnet.

                                        u.AL = atlantisLeistung;
                                        u.Reihenfolge = atlantisLeistung.Reihenfolge; // Die Reihenfolge aus Atlantis wird an den Unterricht übergeben                                    
                                        break;
                                    }                                    
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

                // Alle Atlantisleistungen des aktuellen Abschnitts, die nicht zu den aktuellen Webuntis-Unterrichten gehören, werden zu neuen Unterrichten, denen später die geholten Noten zugeordnet werden.  

                schüler.UnterrichteAktuellAusAtlantis = new Unterrichte();

                foreach (var atlantisLeistung in (from al in atlantisLeistungen.OrderByDescending(x => x.Konferenzdatum)
                                                  where al.SchlüsselExtern == schüler.SchlüsselExtern
                                                  where al.Konferenzdatum >= DateTime.Now.Date
                                                  where !aktuelleFächer.Contains(al.Fach) 
                                                  where al.Gesamtnote == null || al.Gesamtnote == ""
                                                  select al).ToList())
                {
                    schüler.UnterrichteAktuellAusAtlantis.Add(new Unterricht(atlantisLeistung));
                }
            }
        }

        internal void Update(string interessierendeKlasse)
        {
            int i = 0;
                        
            foreach (var schüler in this)
            {
                foreach (var u in schüler.UnterrichteAusWebuntis)
                {
                    if (u.AL != null)
                    {
                        if (u.WL != null)
                        {
                            u.QueryBauen();

                            if (u.WL.Beschreibung != null && u.WL.Query != null)
                            {
                                u.WL.Update();
                                i++;
                            }
                        }
                    }
                }
                
                foreach (var unterricht in schüler.UnterrichteGeholt)
                {
                    if (unterricht.AL != null)
                    {
                        unterricht.WL = new Leistung(unterricht.AL);
                        unterricht.AL.Gesamtpunkte = "";
                        unterricht.AL.Gesamtnote = "";
                        unterricht.AL.Tendenz = "";
                        unterricht.QueryBauen();                        
                        unterricht.WL.Update();
                    }
                }
            }
            
            if (i == 0)
            {
                Global.WriteLine("A C H T U N G: Es wurde keine einzige Leistung angelegt oder verändert. Das kann daran liegen, dass das Notenblatt in Atlantis nicht richtig angelegt ist. Die Tabelle zum manuellen eintragen der Noten muss in der Notensammelerfassung sichtbar sein.");
            }
        }

        internal bool WidersprechendeGesamtnotenImSelbenFachKorrigieren(string interessierendeKlasse)
        {
            var aktuelleFächer = (from s in this from u in s.UnterrichteAusWebuntis select u.Fach + "|" + u.Lehrkraft).Distinct().ToList();

            foreach (var af in aktuelleFächer)
            {
                if ((from aa in aktuelleFächer where Regex.Replace(aa.Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) select aa.Split('|')[0]).Count() > 1)
                {
                    foreach (var schüler in this)
                    {
                        var doppelteNotenLoL = new List<string>();
                        var doppelteNoten = new List<string>();
                        var doppelterLehrer = new List<string>();

                        foreach (var unterricht in (from u in schüler.UnterrichteAusWebuntis where Regex.Replace(u.Fach, @"[\d-]", string.Empty) == Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) select u).ToList())
                        {
                            if (unterricht.WL != null && unterricht.WL.Gesamtnote != null && unterricht.WL.Gesamtnote != "")
                            {
                                if (!doppelteNotenLoL.Contains(unterricht.WL.Gesamtnote + unterricht.WL.Tendenz + "|" + unterricht.WL.Lehrkraft))
                                {
                                    doppelteNotenLoL.Add(unterricht.WL.Gesamtnote + unterricht.WL.Tendenz + "|" + unterricht.WL.Lehrkraft);
                                    doppelterLehrer.Add(unterricht.Lehrkraft);
                                    doppelteNoten.Add(unterricht.WL.Gesamtnote + unterricht.WL.Tendenz);
                                }
                            }
                        }

                        if (doppelteNotenLoL.Count > 1)
                        {
                            Global.WriteLine(" ");
                            Global.WriteLine(" ");
                            Global.WriteLine(" Die Lehrkräfte " + Global.List2String(doppelterLehrer, ",") + " scheinen im Fach " + Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) + " konkurrierend" + (doppelteNoten.Distinct().Count() > 1 ? " und widersprüchlich":"") + " einzutragen.");
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
                                Global.Write("   Wessen Noten (1, ..., " + i + ") sollen im Zeugnis erscheinen? ");
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
                                            foreach (var unt in (from u in s.UnterrichteAusWebuntis where Regex.Replace(u.Fach.Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty) select u).ToList())
                                            {
                                                if (unt.Lehrkraft != lehrkraft)
                                                {
                                                    unt.WL = null;                                                    
                                                    unt.AL = null;
                                                    unt.Bemerkung = lehrkraft + " setzt die Note in " + unt.Fach + ".";
                                                }
                                            }
                                        }
                                    }
                                    gehtNichtWeiter = false;
                                    
                                    if (Global.Rückmeldung.Contains(" * " + lehrkraft.PadRight(3) + "," + Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty).PadRight(4) + ": Die Noten wurden konkurrierend" + (doppelteNoten.Distinct().Count() > 1 ? " und widersprüchlich" : "") + " von mehr als einer Lehrkraft eingetragen. Es werden nur die Noten von " + lehrkraft + " berücksichtigt."))
                                    {
                                        Global.Rückmeldung.Add(" * " + lehrkraft.PadRight(3) + "," + Regex.Replace(af.Split('|')[0], @"[\d-]", string.Empty).PadRight(4) + ": Die Noten wurden konkurrierend" + (doppelteNoten.Distinct().Count() > 1 ? " und widersprüchlich" : "") + " von mehr als einer Lehrkraft eingetragen. Es werden nur die Noten von " + lehrkraft + " berücksichtigt.");
                                    }
                                }
                                catch (Exception)
                                {
                                }

                                Global.WriteLine("");

                            } while (gehtNichtWeiter);

                            // Die Tabelle wird aus den SQL-Zeilen wieder gelöscht, damit nur die letzte Tabelle angezeigt wird.

                            int index = Global.SqlZeilen.IndexOf("/* *-----*------*---------------------------------------------------------------------------------------------------------*                                                          */");

                            for (int ii = Global.SqlZeilen.Count - 1; ii >= index; ii--)
                            {
                                try
                                {
                                    Global.SqlZeilen.RemoveAt(ii);
                                }
                                catch (Exception)
                                {
                                }
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
                Console.WriteLine((" "));
                Console.WriteLine((" "));
                Console.WriteLine(("Leistungen der Klasse " + interessierendeKlasse + " in Atlantis: "));
                Console.WriteLine(("======================================== ").PadRight(interessierendeKlasse.Length, '='));
                Console.WriteLine((" "));
                Global.WriteLine("*-----*------*----------------------------".PadRight(Global.PadRight + 3, '-') + "*");
                Global.WriteLine(("|Name |SuS-Id| Noten + Tendenzen der Klasse " + interessierendeKlasse + " aus Webuntis:").PadRight(Global.PadRight + 3) + "|");
                
                // aktuelle Fächer
                var aF = (from s in this from u in s.UnterrichteAusWebuntis.OrderBy(xx => xx.Reihenfolge) select u.Fach + "|" + interessierendeKlasse + "|" + u.Lehrkraft + "|" + u.Gruppe + "|" + u.LessonNumber).Distinct().ToList();
                
                bool xxxx = false;
                bool reliabwähler = false;

                string b = "*------------+";
                string f = "|       Fach:|";
                string l = "|  Lehrkraft:|";
                string g = "| Teilnehmer:|";
                string w = "|WebUn Atlan:|";
                string k = "|            |";
                string n = "| Unterr.Nr.:|";
                string x = "*------------+";
                string y = "*-----*------*";

                foreach (var aktuelleFach in aF)
                {
                    b += "----+";
                    f += aktuelleFach.Split('|')[0].PadRight(4).Substring(0, 4) + "|";
                    l += aktuelleFach.Split('|')[2].PadRight(4).Substring(0, 4) + "|";
                    g += aktuelleFach.Split('|')[3] == "" ? "Alle|" : "Kurs|";
                    w += "W A |";
                    k += "      ";
                    n += aktuelleFach.Split('|')[4].PadRight(4).Substring(0, 4) + "|";
                    x += "-----";
                    y += "----*";
                }

                Global.WriteLine(x.PadRight(Global.PadRight + 3, '-') + "*");
                Global.WriteLine(f);
                Global.WriteLine(l);
                Global.WriteLine(g);
                Global.WriteLine(n);
                Global.WriteLine(w);
                Global.WriteLine(b.PadRight(Global.PadRight + 3, '-') + "*");

                foreach (var schüler in this)
                {
                    string s = "|" + schüler.Nachname.PadRight(5).Substring(0, 5) + "|" + schüler.SchlüsselExtern + "|";

                    for (int i = 0; i < aF.Count; i++)
                    {
                        var uA = (from u in schüler.UnterrichteAusWebuntis
                                 where u.Fach == aF[i].Split('|')[0]
                                 where u.Lehrkraft == aF[i].Split('|')[2]
                                 where aF[i].Split('|')[3] == u.Gruppe
                                 select u).FirstOrDefault();

                        // Wenn der Nachfolger das selbe Fach ist, ändert sich der delimiter

                        var delimiter = "|";

                        if (i<aF.Count-1)
                        {
                            if (Regex.Replace(aF[i].Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(aF[i+1].Split('|')[0], @"[\d-]", string.Empty))
                            {
                                delimiter = " ";
                            }
                        }
                        

                        // Wenn es zu dem Fach in der Schülergruppe und bei der Lehrkraft einen Unterricht gibt ... 

                        if (uA != null)
                        {
                            if (uA.AL != null)
                            {
                                if (uA.WL != null)
                                {
                                    s += (uA.WL.Gesamtnote + uA.WL.Tendenz).PadRight(2) + (uA.AL.Gesamtnote + uA.AL.Tendenz).PadRight(2) + delimiter;
                                }
                                else
                                {
                                    s += "  " + (uA.AL.Gesamtnote + uA.AL.Tendenz).PadRight(2) + delimiter;
                                }
                            }
                            else
                            {
                                if (uA.WL != null)
                                {
                                    // Wenn der Schüler in Webuntis steht, aber kein Leistungsdatensatz in Atlantis hat
                                    s += (uA.WL.Gesamtnote + uA.WL.Tendenz).PadRight(2) + "XX" + delimiter;
                                    xxxx = true;
                                }
                                else
                                {
                                    if (uA.Bemerkung.Contains("setzt die Note"))
                                    {
                                        s += uA.Bemerkung.Substring(0,2) + "XX" + delimiter;
                                    }
                                    else
                                    {
                                        s += "XXXX|";
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Wenn der Schüler in Webuntis dieses Fach nicht belegt hat, wird geixxt

                            if ((new List<string> { "KR", "ER", "REL" }).Contains(aF[i].Split('|')[0]))
                            {
                                s += "R*  " + delimiter;
                                reliabwähler = true;
                            }
                            else
                            {
                                s += "XXXX" + delimiter;
                                xxxx = true;
                            }
                        }
                    }

                    Global.WriteLine(s);
                }

                Global.WriteLine(y.PadRight(Global.PadRight + 3, '-') + "*");

                if (reliabwähler)
                {
                    var le = (from s in this
                              from u in s.UnterrichteAusWebuntis
                              where (new List<string> { "KR", "ER", "REL", "KR G1", "KR G2" }).Contains(u.Fach)
                              select u.Lehrkraft).FirstOrDefault();

                    if (!Global.Rückmeldung.Contains(" * " + le.PadRight(3) + ",REL : Gibt es Reliabwähler? Evtl. Note in Konferenz ergänzen, falls bewertbar. Bei Abgang/Abschluss Strich."))
                    {
                        Global.Rückmeldung.Add(" * " + le.PadRight(3) + ",REL : Gibt es Reliabwähler? Evtl. Note in Konferenz ergänzen, falls bewertbar. Bei Abgang/Abschluss Strich.");
                    }                    
                }
                if (xxxx)
                {
                    Global.WriteLine("  * X Der Schüler hat den Kurs nicht belegt oder das Untis-Fach kann keinem Datensatz in Atlantis zugeordnet werden.");
                }
                Global.WriteLine("");

            } while (WidersprechendeGesamtnotenImSelbenFachKorrigieren(interessierendeKlasse)); 
        }
        
        internal void GetWebuntisUnterrichte(Unterrichte alleUnterrichte, Gruppen alleGruppen, string interessierendeKlasse)
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
    }
}