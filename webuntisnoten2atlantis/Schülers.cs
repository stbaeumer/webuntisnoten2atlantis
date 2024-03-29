﻿// Published under the terms of GPLv3 Stefan Bäumer 2023.

using Microsoft.Office.Interop.Excel;
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

                            // Nur Schüler, die nicht längeer als 30 Tage ausgetreten sind.

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
            try
            {
                foreach (var schüler in this)
                {
                    foreach (var unterricht in schüler.UnterrichteAusWebuntis)
                    {
                        var w = (from a in alleWebuntisLeistungen
                                 where a.SchlüsselExtern == schüler.SchlüsselExtern
                                 where a.FachAliases.Contains(unterricht.Fach)
                                 where a.Lehrkraft == unterricht.Lehrkraft
                                 where a.Gesamtnote != null
                                 where a.Gesamtnote != ""
                                 where a.Klasse != null 
                                 where a.Klasse != ""
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

                                Global.Rückmeldungen.AddRückmeldung(new Rückmeldung(w[0].Lehrkraft, w[0].Fach, "Die Noten in Webuntis (" + w[0].Name + ") sind nicht eindeutig. Ist Note (" + w[0].Gesamtnote + w[0].Tendenz + ") korrekt? Evtl. in der Konferenz ändern lassen."));

                                if ((from ww in w select ww.Datum).Distinct().ToList().Count() == 1)
                                {
                                    throw new Exception("Achtung: Beide Unterrichte " + w[0].Lehrkraft + "," + w[0].Fach + " haben dasselbe Datum, aber unterschiedliche Noten.");
                                }
                            }
                        }

                        if (w.Count > 0)
                        {
                            unterricht.LeistungW = new Leistung(
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

                            if (w[0].Gesamtnote == "-")
                            {
                                Global.Rückmeldungen.AddRückmeldung(new Rückmeldung(w[0].Lehrkraft, w[0].Fach, "Bei einem Strich (" + w[0].Name + ") muss rechtzeitig vor der Konferenz eine Bemerkung mit Ulla vereinbart werden."));
                            }
                        }
                    }
                }
                var susMitLeistungen = (from s in this
                                        from u in s.UnterrichteAusWebuntis
                                        where u.LeistungW != null
                                        where u.LeistungW.Gesamtnote != null
                                        where u.LeistungW.Gesamtnote != ""
                                        select s).Distinct().ToList();

                var faecherMitLeistungen = (from s in this
                                            from u in s.UnterrichteAusWebuntis
                                            where u.LeistungW != null
                                            where u.LeistungW.Gesamtnote != null
                                            where u.LeistungW.Gesamtnote != ""
                                            select u.Fach).Distinct().ToList();

                Console.WriteLine(("Anzahl der verschiedenen SuS mit Noten in Webuntis").PadRight(Global.PadRight - 2, '.') + susMitLeistungen.Count.ToString().PadLeft(6));
                Console.WriteLine(("Verschiedene Fächer mit Gesamtleistungen: " + Global.List2String(faecherMitLeistungen, ",")).PadRight(Global.PadRight - 2, '.') + susMitLeistungen.Count.ToString().PadLeft(6));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void ZweiLehrerEinFach(string interessierendeKlasse)
        {
            foreach (var aktuellerU in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell)
            {
                foreach (var schüler in this)
                {
                    // Zwei oder mehr unterrichten Fach, 

                    var zweiUnterrichtenDasGleicheFach = (from u in schüler.UnterrichteAusWebuntis
                                                          where Regex.Replace(u.Fach, @"[\d-]", string.Empty) == Regex.Replace(aktuellerU.Fach, @"[\d-]", string.Empty)
                                                          select u.Lehrkraft).Distinct().ToList();

                    if (zweiUnterrichtenDasGleicheFach.Count > 1)
                    {
                        foreach (var leh in zweiUnterrichtenDasGleicheFach)
                        {
                            // davon darf mindestens einer keine einzige Gesamtnote eingetragen haben

                            var keineEinzigeNoteEingetragen = (from s in this
                                                               from u in s.UnterrichteAusWebuntis
                                                               where Regex.Replace(u.Fach, @"[\d-]", string.Empty) == Regex.Replace(aktuellerU.Fach, @"[\d-]", string.Empty)
                                                               where u.LeistungW == null || u.LeistungW.Gesamtnote == null || u.LeistungW.Gesamtnote == ""
                                                               select u.Lehrkraft).Distinct().ToList();

                            // und genau ein anderer muss Noten eingetragen haben

                            var noteEingetragen = (from s in this
                                                   from u in s.UnterrichteAusWebuntis
                                                   where Regex.Replace(u.Fach, @"[\d-]", string.Empty) == Regex.Replace(aktuellerU.Fach, @"[\d-]", string.Empty)
                                                   where u.LeistungW != null
                                                   where u.LeistungW.Gesamtnote != null && u.LeistungW.Gesamtnote != ""
                                                   select u.Lehrkraft).Distinct().ToList();

                            // Dann werden alle Webuntis- & Atlantisleistungen bei dem Nichteintrager entfernt

                            if (noteEingetragen.Count == 1 && keineEinzigeNoteEingetragen.Count > 0)
                            {
                                foreach (var s in this)
                                {
                                    foreach (var unt in (from u in s.UnterrichteAusWebuntis where Regex.Replace(u.Fach.Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(aktuellerU.Fach, @"[\d-]", string.Empty) select u).ToList())
                                    {
                                        foreach (var keineNote in (from k in keineEinzigeNoteEingetragen
                                                                   where k != noteEingetragen[0]
                                                                   select k).ToList())
                                        {
                                            if (unt.Lehrkraft == keineNote)
                                            {
                                                unt.LeistungW = null;
                                                unt.LeistungA = null;
                                                unt.Bemerkung = noteEingetragen[0] + " setzt die Note in " + Regex.Replace(unt.Fach.Split('|')[0], @"[\d-]", string.Empty) + ".";

                                                Global.Rückmeldungen.AddRückmeldung(new Rückmeldung(unt.Lehrkraft, Regex.Replace(unt.Fach.Split('|')[0], @"[\d-]", string.Empty), "In " + Regex.Replace(aktuellerU.Fach, @"[\d-]", string.Empty).PadRight(3) + " werden nur die Noten von " + noteEingetragen[0] + " in das Zeugnis übernommen."));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        internal void GeholteLeistungenBehandeln(string interessierendeKlasse, List<string> aktSj, string hzJz)
        {
            try
            {
                var aktuelleUnterrichte = new List<string>();

                foreach (var schüler in this)
                {
                    foreach (var uA in (from u in schüler.UnterrichteAusWebuntis
                                        where u.LeistungA != null
                                        select
                                        u).OrderBy(x => x.Reihenfolge).ToList())
                    {
                        if (!aktuelleUnterrichte.Contains(uA.LeistungA.Fach))
                        {
                            aktuelleUnterrichte.Add(uA.LeistungA.Fach);
                        }
                    }
                }

                var geholteUnterrichte = new Unterrichte();

                foreach (var schüler in this)
                {
                    foreach (var uG in (from g in schüler.UnterrichteGeholt
                                        where !(g.LeistungA.Schuljahr == aktSj[0] + "/" + aktSj[1] && g.LeistungA.HzJz == hzJz)
                                        select g).ToList())
                    {
                        if (!aktuelleUnterrichte.Contains(uG.Fach))
                        {
                            if (!(from g in geholteUnterrichte
                                  where g.Fach == uG.Fach
                                  where g.LeistungA.Konferenzdatum.ToShortDateString() == uG.LeistungA.Konferenzdatum.ToShortDateString()
                                  select g).Any())
                            {
                                geholteUnterrichte.Add(uG);
                            }
                        }
                    }
                }

                Console.WriteLine(" ");
                Console.WriteLine("Aktuell werden die Fächer " + Global.List2String90(aktuelleUnterrichte, ",") + " unterrichtet." + (geholteUnterrichte.Count == 0 ? "Es gibt keine alten Unterrichte, die geholt werden könnten." : " Folgende Fächer können geholt werden:"));

                List<string> dieseFächerHolen = new List<string>();
                bool gehtNichtWeiter = true;
                int anzahlOriginal = geholteUnterrichte.Count();
                int durchläufe = 0;

                do
                {
                    if (geholteUnterrichte.Count > 0)
                    {
                        int i = 1;

                        var xx = new Unterrichte();

                        foreach (var gU in geholteUnterrichte)
                        {
                            if (!dieseFächerHolen.Contains(gU.Fach + "|" + gU.LeistungA.Konferenzdatum.ToShortDateString()) &&
                                !dieseFächerHolen.Contains(gU.FachnameAtlantis + "|" + gU.LeistungA.Konferenzdatum.ToShortDateString()))
                            {
                                var fach = gU.Fach;
                                fach = fach.Replace("  ", " ");
                                fach = fach.Substring(0, (fach.IndexOf(' ') == -1 ? fach.Length : fach.IndexOf(' ') + 2));


                                int anzahl = (from s in this
                                              from u in s.UnterrichteGeholt
                                              where u.LeistungA.Konferenzdatum.ToShortDateString() == gU.LeistungA.Konferenzdatum.ToShortDateString().ToString()
                                              where u.Fach == gU.Fach
                                              where !(from ss in s.UnterrichteAusWebuntis
                                                      where ss.Fach.Substring(0, (ss.Fach.IndexOf(' ') == -1 ? ss.Fach.Length : ss.Fach.IndexOf(' ') + 2)) == fach
                                                      select ss).Any()
                                              select u).Count();

                                var klassen = (from s in this
                                               from u in s.UnterrichteGeholt
                                               where u.LeistungA.Konferenzdatum.ToShortDateString() == gU.LeistungA.Konferenzdatum.ToShortDateString()
                                               where u.Fach == gU.Fach
                                               where !(from ss in s.UnterrichteAusWebuntis
                                                       where ss.Fach.Substring(0, (ss.Fach.IndexOf(' ') == -1 ? ss.Fach.Length : ss.Fach.IndexOf(' ') + 2)) == fach
                                                       select ss).Any()
                                               select u.Klassen).Distinct().ToList();

                                var sus = (from s in this
                                           from u in s.UnterrichteGeholt
                                           where u.LeistungA.Konferenzdatum.ToShortDateString() == gU.LeistungA.Konferenzdatum.ToShortDateString()
                                           where u.Fach == gU.Fach
                                           where !(from ss in s.UnterrichteAusWebuntis
                                                   where ss.Fach.Substring(0, (ss.Fach.IndexOf(' ') == -1 ? ss.Fach.Length : ss.Fach.IndexOf(' ') + 2)) == fach
                                                   select ss).Any()
                                           select u.LeistungA.SchlüsselExtern).Distinct().ToList();

                                var susNamen = (from s in this
                                                from u in s.UnterrichteGeholt
                                                where u.LeistungA.Konferenzdatum.ToShortDateString() == gU.LeistungA.Konferenzdatum.ToShortDateString()
                                                where u.Fach == gU.Fach
                                                where !(from ss in s.UnterrichteAusWebuntis
                                                        where ss.Fach.Substring(0, (ss.Fach.IndexOf(' ') == -1 ? ss.Fach.Length : ss.Fach.IndexOf(' ') + 2)) == fach
                                                        select ss).Any()
                                                select s.Zähler + ".(" + u.LeistungA.Gesamtnote + ")").ToList();

                                var nokId = (from s in this
                                             from u in s.UnterrichteGeholt
                                             where u.LeistungA.Konferenzdatum.ToShortDateString() == gU.LeistungA.Konferenzdatum.ToShortDateString()
                                             where u.Fach == gU.Fach
                                             where !(from ss in s.UnterrichteAusWebuntis
                                                     where ss.Fach.Substring(0, (ss.Fach.IndexOf(' ') == -1 ? ss.Fach.Length : ss.Fach.IndexOf(' ') + 2)) == fach
                                                     select ss).Any()
                                             select u.LeistungA.NokId.ToString()).Distinct().Take(3).ToList();


                                if (anzahl > 0)
                                {
                                    if (!dieseFächerHolen.Contains(gU.Fach + "|" + gU.LeistungA.Konferenzdatum.ToShortDateString()) &&
                                        !dieseFächerHolen.Contains(gU.FachnameAtlantis)
                                        )
                                    {
                                        xx.Add(gU);
                                        Console.WriteLine("".PadRight(dieseFächerHolen.Count + 1, ' ') + " " + i + ". " + gU.Fach.PadRight(5) + " |" + gU.LeistungA.Konferenzdatum.ToShortDateString() + " |" + gU.LeistungA.Schuljahr + " |" + gU.LeistungA.HzJz + " |Anzahl: " + anzahl.ToString().PadLeft(2) + " SuS |Klassen: " + Global.List2String(klassen, ",") + " |Notenkopf-ID: " + Global.List2String(nokId, ",") + " ...\n           " + "".PadRight(durchläufe) + "|SuS: " + Global.List2String90(susNamen, ","));
                                        i++;
                                    }
                                }
                            }
                        }

                        geholteUnterrichte = xx;

                        Console.Write(" ".PadRight(dieseFächerHolen.Count + 1, ' ') + " Geben Sie die Ziffer 1, ..., " + (i - 1) + " ein oder drücken Sie ENTER, wenn keine " + (dieseFächerHolen.Count > 0 ? "weitere " : "") + "Note geholt werden soll.");

                        string auswahl = Console.ReadKey().Key.ToString();

                        auswahl = auswahl.Substring(auswahl.Length - 1, 1);
                        Console.WriteLine("");

                        try
                        {
                            if (auswahl != "r") // Wenn nicht ENTER gedrückt wurde
                            {
                                var intAuswahl = Convert.ToInt32(auswahl.ToString());

                                var x = (from s in this
                                         from uuu in s.UnterrichteAktuellAusAtlantis
                                         where uuu.Fach == geholteUnterrichte[intAuswahl - 1].Fach
                                         select uuu.LeistungA).ToList();

                                if (x.Count > 0)
                                {
                                    geholteUnterrichte[intAuswahl - 1].FachnameAtlantis = geholteUnterrichte[intAuswahl - 1].Fach;
                                    dieseFächerHolen.Add(geholteUnterrichte[intAuswahl - 1].Fach + "|" + geholteUnterrichte[intAuswahl - 1].LeistungA.Konferenzdatum.ToShortDateString());
                                    Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell.AddUnterrichte(geholteUnterrichte[intAuswahl - 1]);
                                    Global.Rückmeldungen.AddRückmeldung(new Rückmeldung("", geholteUnterrichte[intAuswahl - 1].Fach, "Die Noten aus dem Fach " + geholteUnterrichte[intAuswahl - 1].Fach + " werden für " + x.Count + " SuS aus dem Zeugnis vom " + geholteUnterrichte[intAuswahl - 1].LeistungA.Konferenzdatum.ToShortDateString() + " geholt."));
                                }
                                else
                                {
                                    bool nichtweiter = true;

                                    Console.WriteLine(" ");
                                    Console.WriteLine("     Wie heißt das Atlantisfach, dem Sie Ihre Auswahl '" + geholteUnterrichte[intAuswahl - 1].Fach + "' zuordnen möchten?");

                                    do
                                    {
                                        var möglicheFächer = (from s in this
                                                              from uuu in s.UnterrichteAktuellAusAtlantis
                                                              select uuu.LeistungA.Fach + "|" + geholteUnterrichte[intAuswahl - 1].LeistungA.Konferenzdatum.ToShortDateString()).Distinct().ToList();
                                        int m = 1;

                                        foreach (var item in (from s in this
                                                              from uuu in s.UnterrichteAktuellAusAtlantis
                                                              select uuu.LeistungA.Fach).Distinct().ToList())
                                        {
                                            Console.WriteLine("     " + m.ToString().PadLeft(2) + ". " + item);
                                            m++;
                                        }
                                        Console.Write("       Geben Sie die Ziffer 1, ..., " + (möglicheFächer.Count) + " ein.");

                                        string auswahl1 = Console.ReadKey().Key.ToString();

                                        auswahl1 = auswahl1.Substring(auswahl1.Length - 1, 1);
                                        Console.WriteLine("");

                                        try
                                        {
                                            var intAuswahl1 = Convert.ToInt32(auswahl1.ToString());

                                            if (intAuswahl1 >= 1 && m >= intAuswahl1)
                                            {
                                                dieseFächerHolen.Add(möglicheFächer[intAuswahl1 - 1]);
                                                //dieseFächerHolen.Add(geholteUnterrichte[intAuswahl - 1].Fach);
                                                nichtweiter = false;
                                                geholteUnterrichte[intAuswahl - 1].FachnameAtlantis = möglicheFächer[intAuswahl1 - 1].Split('|')[0];
                                                Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell.AddUnterrichte(geholteUnterrichte[intAuswahl - 1]);

                                                // FachnameAtlantis an alle SuS zuweisen

                                                int a = 0;
                                                foreach (var schüler in this)
                                                {
                                                    foreach (var unterrichtGeholt in schüler.UnterrichteGeholt)
                                                    {
                                                        if (unterrichtGeholt.Fach == geholteUnterrichte[intAuswahl - 1].Fach)
                                                        {
                                                            unterrichtGeholt.FachnameAtlantis = möglicheFächer[intAuswahl1 - 1].Split('|')[0];
                                                            a++;
                                                        }
                                                    }
                                                }
                                                Global.Rückmeldungen.AddRückmeldung(new Rückmeldung("", geholteUnterrichte[intAuswahl - 1].Fach, "Die Noten aus dem Fach " + geholteUnterrichte[intAuswahl - 1].Fach + " werden für " + a + " SuS aus dem Zeugnis vom " + geholteUnterrichte[intAuswahl - 1].LeistungA.Konferenzdatum.ToShortDateString() + " geholt."));
                                            }
                                        }
                                        catch (Exception)
                                        {
                                        }
                                    } while (nichtweiter);
                                }
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
                    durchläufe++;
                } while (gehtNichtWeiter);

                Global.WriteLine((dieseFächerHolen.Count > 0 ? "  Es werden diese Fächer geholt: " + Global.List2String(dieseFächerHolen, ",") : " Es werden keine Fächer geholt."));

                foreach (var schüler in this)
                {
                    string r = schüler.UnterrichteAusWebuntis.UmGeholteUnterrichteErweitern(
                           dieseFächerHolen,
                           schüler.UnterrichteGeholt,
                           schüler.UnterrichteAktuellAusAtlantis,
                           schüler.UnterrichteAusWebuntis,
                           schüler.Nachname
                           );
                    Global.Rückmeldungen.AddRückmeldung(new Rückmeldung("", "", r));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void ChatErzeugen(Lehrers alleAtlantisLehrer, string interessierendeKlasse, string hzJz, string user, List<string> aktSj)
        {
            var url = "https://teams.microsoft.com/l/chat/0/0?users=";

            if (user != "LG" && !interessierendeKlasse.StartsWith("G"))
            {
                url += "wolfgang.leuering@berufskolleg-borken.de,";
            }
            if (user != "LS" && interessierendeKlasse.StartsWith("G"))
            {
                url += "klaus.lienenklaus@berufskolleg-borken.de,";
            }
            if (user != "BM")
            {
                url += "stefan.baeumer@berufskolleg-borken.de,";
            }
            if (!user.ToUpper().StartsWith("MOR"))
            {
                url += "ursula.moritz@berufskolleg-borken.de,";
            }
            DateTime konferenz = DateTime.Now;

            int anzahlTeamsChat = 0;
            List<string> überzählige = new List<string>();

            foreach (var schüler in this)
            {
                foreach (var unterricht in schüler.UnterrichteAusWebuntis.OrderBy(x => x.Reihenfolge))
                {
                    var mail = (from l in alleAtlantisLehrer where l.Kuerzel == unterricht.Lehrkraft select l.Mail).FirstOrDefault();

                    if (unterricht.Konferenzdatum != null && unterricht.Konferenzdatum.Year != 1)
                    {
                        konferenz = unterricht.Konferenzdatum;
                    }
                    if (mail != null && mail != "" && !url.Contains(mail))
                    {
                        anzahlTeamsChat++;

                        if (anzahlTeamsChat < 7)
                        {
                            url += mail + ",";
                        }
                        else
                        {
                            if (!überzählige.Contains(mail))
                            {
                                überzählige.Add(mail);
                            }
                        }
                    }
                }
            }

            Global.WriteLine("  ");
            Global.WriteLine("Link zum Teams-Chat mit den LuL der Klasse " + interessierendeKlasse + ":");
            Global.WriteLine(" " + url.TrimEnd(','));

            if (überzählige.Count > 0)
            {
                Global.WriteLine("  A C H T U N G: Folgende LuL müssen dem Teams-Chat zusätzlich manuell hinzugefügt werden:");
                foreach (var item in überzählige)
                {
                    Global.WriteLine("   " + item);
                }
            }

            var rückmeldung = "/*" +
                "\n" + interessierendeKlasse + "_" + aktSj[0] + "-" + aktSj[1] + "_" + hzJz + "_Notenliste" +
                "\n " +
                "\nHallo LuL der " + interessierendeKlasse + ", " +
                "\ndie aktuelle Notenlisten befindet sich unter folegndem Link:" +
                "\n\nhttps://wiki.berufskolleg-borken.de/doku.php?id=notenlisten:start" +
                "\n\nBeste Grüße" +
                "\nZeugnisschreibung";   

            if (Global.Rückmeldungen.Count() > 0)
            {
                rückmeldung += "\n\nFolgende automatisch generierte Meldungen sollen uns helfen Handlungsbedarfe bzw. Fehler zu erkennen:\n\n";

                foreach (var r in Global.Rückmeldungen)
                {
                    rückmeldung += r.ToString();
                }
            }

            rückmeldung += "\n\nBei Unklarheiten bitte immer gerne melden.";

            rückmeldung += "\n\nKennwort: https://wiki.berufskolleg-borken.de/doku.php?id=schulleitung:mitteilungen:2023-04-28#verschluesselung\n";

            rückmeldung += "\n\nBeste Grüße,\nZeugnisschreibung" + "\n\n */";

            Global.SqlZeilen.Insert(Global.SqlZeilen.Count(), rückmeldung);

            Fächer fehlendeFächer = new Fächer();

            foreach (var schüler in this)
            {
                foreach (var unterricht in schüler.UnterrichteAusWebuntis)
                {
                    if (unterricht.LeistungA != null)
                    {
                        if (unterricht.LeistungW == null)
                        {
                            // Die Note wird nur dann als fehlend moniert, wenn auch im selben Fach anderweitig keine Note vergeben wurde. 

                            if (!(from u in schüler.UnterrichteAusWebuntis
                                  where u.LeistungW != null
                                  where Regex.Replace(u.LeistungW.Fach, @"[\d-]", string.Empty) == Regex.Replace(unterricht.Fach, @"[\d-]", string.Empty)
                                  where u.LeistungW.Gesamtnote != null
                                  where u.LeistungW.Gesamtnote != ""
                                  select u).Any())
                            {
                                fehlendeFächer.Add(new Fach(unterricht.Fach, unterricht.Lehrkraft));
                            }
                        }
                    }
                }
            }

            if (fehlendeFächer.Count > 0)
            {
                foreach (var fach in (from f in fehlendeFächer select f.Name).Distinct().ToList())
                {
                    var lehrer = (from f in fehlendeFächer where f.Name == fach select f.Lehrkraft).FirstOrDefault();
                    var anzahl = (from f in fehlendeFächer where f.Name == fach select f).Count();

                    Global.Rückmeldungen.AddRückmeldung(new Rückmeldung(lehrer, fach, "Kann es sein, dass " + anzahl.ToString().PadLeft(2) + " Note" + (anzahl > 1 ? "n" : "") + " fehlen?"));
                }
            }
        }

        internal void GetAtlantisLeistungen(string connectionStringAtlantis, List<string> aktSj, string user, string interessierendeKlasse, string hzJz)
        {
            // Alle AtlantisLeistungen dieser Klasse und der Parallelklassen in allen Jahren

            var atlantisLeistungen = new Leistungen(connectionStringAtlantis, aktSj, user, interessierendeKlasse, this, hzJz);

            atlantisLeistungen.IstNotenblattAngelegt(hzJz, aktSj, interessierendeKlasse);

            int a = 0;
            int i = 0;
            int g = 0;
            int u = 0;

            foreach (var schüler in this)
            {
                a += schüler.GetAtlantisLeistungen(atlantisLeistungen, hzJz, aktSj);

                i += schüler.InfragekommendeUnterrichteFürSpätereZuordnungAnlegen(atlantisLeistungen, hzJz, aktSj);

                g += schüler.GeholteUnterrichteHinzufügen(atlantisLeistungen);

                u += schüler.UnterrichteAktuellAusAtlantisHolen(atlantisLeistungen, hzJz, aktSj);

                if (hzJz == "JZ")
                {
                    // Für alle im erten Halbjahr abgeschlossene Unterrichte aus Webuntis wird geprüft, ob ...

                    var abgeschlosseneUdesErstenHalbjahresAusWebuntis = (from uAW in schüler.UnterrichteAusWebuntis
                                                                         where uAW.Startdate >= new DateTime(Convert.ToInt32(aktSj[0]), 08, 01)
                                                                         where uAW.Enddate <= new DateTime(Convert.ToInt32(aktSj[1]) + 2000, 02, 1)
                                                                         select uAW).ToList();

                    foreach (var item in abgeschlosseneUdesErstenHalbjahresAusWebuntis)
                    {
                        // ... es es auch einen geholten Atlantis-Unterricht im selben Fach gibt.

                        if ((from ge in schüler.UnterrichteGeholt where ge.Fach == item.Fach select ge).Any())
                        {
                            // Wenn ja, dann hat der geholte Unterricht aus Atlantis Vorrang vor dem abgeschlossenen Unterricht in Webuntis

                            schüler.UnterrichteAusWebuntis.Remove(item);
                        }
                    }
                }
            }

            Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell = new Unterrichte();
            Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell.AddRange(Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuellUnsortiert.OrderBy(x => x.Reihenfolge));

            Console.WriteLine(("Alle Unterrichte aus Atlantis der Klasse " + interessierendeKlasse + " aus Webuntis ").PadRight(Global.PadRight - 2, '.') + a.ToString().PadLeft(6));
            Console.WriteLine(("Infragekommende Unterrichte für spätere Zuordnung").PadRight(Global.PadRight - 2, '.') + i.ToString().PadLeft(6));
            Console.WriteLine(("Geholte Unterrichte der Klasse " + interessierendeKlasse).PadRight(Global.PadRight - 2, '.') + g.ToString().PadLeft(6));
            Console.WriteLine(("Aktuelle Unterrichte aus Atlantis").PadRight(Global.PadRight - 2, '.') + u.ToString().PadLeft(6));
            Console.WriteLine(("Alle verschiedenen Unterrichte der Klasse " + interessierendeKlasse + " aus Webuntis ").PadRight(Global.PadRight - 2, '.') + i.ToString().PadLeft(6));
        }

        internal void Update(string interessierendeKlasse)
        {
            foreach (var schüler in this)
            {
                schüler.QueriesBauenUpdate();
            }
        }

        internal bool KorrekturenVornehmen(string interessierendeKlasse)
        {
            // Wenn Fächer aus Untis keiner Leistung in Atlantis zugeordnet werden konnte, wird nachgefragt:

            var unterrichtOhneAtlantisLeistung = (from af in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell
                                                  where af.LeistungA == null
                                                  select af).ToList();

            foreach (var uOhneA in unterrichtOhneAtlantisLeistung)
            {
                bool gehtNichtWeiter = true;

                Console.WriteLine("\nDer Unterricht im Fach " + uOhneA.Fach + " konnte keinem Unterricht in Atlantis zugeordnet werden. Bitte manuell zuordnen:");

                int i = 0;

                // Damit die Liste nicht zweistellig wird, wird auf den Anfangsbuchstaben gefiltert.

                var infragekommendeLeistungenAgefiltert = uOhneA.InfragekommendeLeistungenA.Where(x => x.Fach.Substring(0, 1) == uOhneA.Fach.Substring(0, 1)).ToList();

                // Falls es kein Fach mit dem selben Anfangsbuchstaben gibt, werden alle Fächer vorgeschlagen

                if (infragekommendeLeistungenAgefiltert.Count == 0)
                {
                    infragekommendeLeistungenAgefiltert = uOhneA.InfragekommendeLeistungenA;
                }

                foreach (var infragekommendeA in infragekommendeLeistungenAgefiltert)
                {
                    // Bereits erfolgreich zugeordnete Unterrichte werden nicht zur Auswahl angeboten.

                    if (!(from af in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell
                          where af.LeistungA != null
                          where infragekommendeA.FachAliases.Contains(af.LeistungA.Fach)
                          select af).Any())
                    {
                        i++;
                        Console.WriteLine(" " + i.ToString().PadLeft(2) + "." + infragekommendeA.Fach.PadRight(6) + ": Konferenzdatum: " + (infragekommendeA.Konferenzdatum.Year == 1 ? " Noch kein Konferenzdatum festgelegt." : infragekommendeA.Konferenzdatum.ToShortDateString()));
                    }
                }

                do
                {
                    Console.Write("   Welches Atlantisfach wollen Sie zuordnen? (1, ..., " + i + "). Mit Enter wird das Fach verworfen:");
                    ConsoleKeyInfo input;
                    input = Console.ReadKey(true);

                    if (input.Key == ConsoleKey.Enter)
                    {
                        gehtNichtWeiter = false;

                        for (int ii = 0; ii < Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell.Count; ii++)
                        {
                            if (uOhneA.Fach == Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell[ii].Fach)
                            {
                                Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell.RemoveAt(ii);
                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            var auswahl = Char.ToUpper(input.KeyChar);

                            var intAuswahl = Convert.ToInt32(auswahl.ToString());

                            if (intAuswahl >= 1 && intAuswahl <= i)
                            {
                                uOhneA.LeistungA = new Leistung(infragekommendeLeistungenAgefiltert[intAuswahl - 1], "");
                                uOhneA.Reihenfolge = uOhneA.LeistungA.Reihenfolge;
                                gehtNichtWeiter = false;

                                foreach (var schüler in this)
                                {
                                    var u = (from uu in schüler.UnterrichteAusWebuntis
                                             where uu.Fach == uOhneA.Fach
                                             where uu.Lehrkraft == uOhneA.Lehrkraft
                                             where uu.LeistungA == null
                                             select uu).FirstOrDefault();

                                    var ala = (from al in schüler.UnterrichteAktuellAusAtlantis
                                               where al.Fach == infragekommendeLeistungenAgefiltert[intAuswahl - 1].Fach
                                               select al.LeistungA).FirstOrDefault();

                                    var x = (from aa in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell
                                             where aa.Fach == uOhneA.Fach
                                             where aa.FachnameAtlantis == null || aa.FachnameAtlantis == ""
                                             select aa).FirstOrDefault();

                                    // Frisch zugeordnete Unterrichte müssen aus den geholten Unterrichten entfernt werden,
                                    // damit sie später nicht zum holen angeboten werden.

                                    for (int g = 0; g < schüler.UnterrichteGeholt.Count; g++)
                                    {
                                        if (schüler.UnterrichteGeholt[g].Fach == infragekommendeLeistungenAgefiltert[intAuswahl - 1].Fach)
                                        {
                                            schüler.UnterrichteGeholt.RemoveAt(g);
                                        }
                                    }


                                    if (x != null)
                                    {
                                        x.FachnameAtlantis = infragekommendeLeistungenAgefiltert[intAuswahl - 1].Fach;
                                    }

                                    if (u != null)
                                    {
                                        if (ala != null)
                                        {
                                            u.LeistungA = new Leistung(ala, uOhneA.Fach + "->" + infragekommendeLeistungenAgefiltert[intAuswahl - 1].Fach + "|");
                                            u.LeistungA.Lehrkraft = u.Lehrkraft;
                                            Global.Rückmeldungen.AddRückmeldung(new Rückmeldung(uOhneA.Lehrkraft, uOhneA.Fach, "Es wurde das Atlantisfach " + infragekommendeLeistungenAgefiltert[intAuswahl - 1].Fach + " zugeordnet."));
                                            u.Reihenfolge = u.LeistungA.Reihenfolge;
                                        }
                                    }
                                }
                            }
                            Console.WriteLine("");
                            Console.WriteLine("    Ihre Auswahl: " + infragekommendeLeistungenAgefiltert[intAuswahl - 1].Fach);

                            return true;
                        }
                        catch (Exception)
                        {
                        }
                    }
                } while (gehtNichtWeiter);
            }

            // Wenn mehrere LoL ein Fach unterrichten:

            foreach (var af in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell)
            {
                if ((from aa in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell where Regex.Replace(aa.Fach, @"[\d-]", string.Empty) == Regex.Replace(af.Fach, @"[\d-]", string.Empty) select aa.Fach).Count() > 1)
                {
                    foreach (var schüler in this)
                    {
                        var doppelteNotenLoL = new List<string>();
                        var doppelteNoten = new List<string>();
                        var doppelterLehrer = new List<string>();
                        var doppelteNotenSuS = new List<string>();
                        var doppelteLuLKlasseFachSuS = new List<string>();

                        foreach (var unterricht in (from u in schüler.UnterrichteAusWebuntis where Regex.Replace(u.Fach, @"[\d-]", string.Empty) == Regex.Replace(af.Fach, @"[\d-]", string.Empty) select u).ToList())
                        {
                            if (unterricht.LeistungW != null && unterricht.LeistungW.Gesamtnote != null && unterricht.LeistungW.Gesamtnote != "")
                            {
                                if (!doppelteNotenLoL.Contains(unterricht.LeistungW.Lehrkraft + "|" + unterricht.LeistungW.Gesamtnote + unterricht.LeistungW.Tendenz))
                                {
                                    doppelteNotenLoL.Add(unterricht.LeistungW.Lehrkraft + "|" + unterricht.LeistungW.Gesamtnote + unterricht.LeistungW.Tendenz);
                                    doppelterLehrer.Add(unterricht.Lehrkraft);
                                    doppelteNoten.Add(unterricht.LeistungW.Gesamtnote + unterricht.LeistungW.Tendenz);
                                    doppelteNotenSuS.Add(unterricht.LeistungW.Name);
                                    doppelteLuLKlasseFachSuS.Add(unterricht.LeistungW.Lehrkraft + "|Bsp.:" + schüler.Nachname + ":" + unterricht.LeistungW.Gesamtnote + unterricht.LeistungW.Tendenz);
                                }
                            }
                        }

                        foreach (var unterricht in (from u in schüler.UnterrichteAusWebuntis where Regex.Replace(u.Fach, @"[\d-]", string.Empty) == Regex.Replace(af.Fach, @"[\d-]", string.Empty) select u).ToList())
                        {
                            if (unterricht.LeistungW != null && unterricht.LeistungW.Gesamtnote != null && unterricht.LeistungW.Gesamtnote != "")
                            {
                                if (!doppelteNotenLoL.Contains(unterricht.LeistungW.Lehrkraft + "|" + unterricht.LeistungW.Gesamtnote + unterricht.LeistungW.Tendenz))
                                {
                                    doppelteNotenLoL.Add(unterricht.LeistungW.Lehrkraft + "|" + unterricht.LeistungW.Gesamtnote + unterricht.LeistungW.Tendenz);
                                    doppelterLehrer.Add(unterricht.Lehrkraft);
                                    doppelteNoten.Add(unterricht.LeistungW.Gesamtnote + unterricht.LeistungW.Tendenz);
                                    doppelteNotenSuS.Add(unterricht.LeistungW.Name);
                                    doppelteLuLKlasseFachSuS.Add(unterricht.LeistungW.Lehrkraft + "|Bsp.:" + schüler.Nachname + ":" + unterricht.LeistungW.Gesamtnote + unterricht.LeistungW.Tendenz);
                                }
                            }
                        }

                        if (doppelteNotenLoL.Count > 1)
                        {
                            Console.WriteLine(" ");
                            Console.WriteLine(" ");
                            Console.WriteLine(" Die Lehrkräfte " + Global.List2String(doppelterLehrer, ",") + " haben konkurrierend" + (doppelteNoten.Distinct().Count() > 1 ? " und widersprüchlich" : "") + " eingetragen.");
                            int i = 0;

                            foreach (var aF in (from a in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell
                                                where doppelterLehrer.Contains(a.Lehrkraft)
                                                where Regex.Replace(a.Fach, @"[\d-]", string.Empty) == Regex.Replace(af.Fach, @"[\d-]", string.Empty)
                                                select a).ToList())
                            {
                                i++;

                                var x = (from s in schüler.UnterrichteAusWebuntis
                                         where s.LeistungW != null
                                         where s.LeistungW.Fach == aF.Fach
                                         where s.LeistungW.Gesamtnote != null
                                         where s.LeistungW.Gesamtnote != ""
                                         select s.LeistungW).FirstOrDefault();

                                Console.WriteLine("  " + i + ". " + aF.Fach.PadRight(7) + "|" + aF.Lehrkraft.PadRight(3) + "|Bsp: " + x.Name + "|Note: " + x.Gesamtnote + "|Zeile: " + x.MarksPerLessonZeile + "|Eingetragen: " + x.Datum.ToShortDateString());
                            }
                            bool gehtNichtWeiter = true;

                            do
                            {
                                Console.Write("   Welche Noten (1, ..., " + i + ") sollen im Fach " + Regex.Replace(af.Fach, @"[\d-]", string.Empty) + " im Zeugnis erscheinen? ");
                                ConsoleKeyInfo input;
                                input = Console.ReadKey(true);

                                try
                                {
                                    var auswahl = Char.ToUpper(input.KeyChar);

                                    var intAuswahl = Convert.ToInt32(auswahl.ToString());

                                    var lehrkraft = ((from a in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell where doppelterLehrer.Contains(a.Lehrkraft) where Regex.Replace(a.Fach, @"[\d-]", string.Empty) == Regex.Replace(af.Fach, @"[\d-]", string.Empty) select a).ToList()[intAuswahl - 1]).Lehrkraft;
                                    var fach = ((from a in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell where doppelterLehrer.Contains(a.Lehrkraft) where Regex.Replace(a.Fach, @"[\d-]", string.Empty) == Regex.Replace(af.Fach, @"[\d-]", string.Empty) select a).ToList()[intAuswahl - 1]).Fach;

                                    if (intAuswahl <= i)
                                    {
                                        // Alle Webuntisleistungen der Unterrichte anderer LuL in diesem Fach werden gelöscht.

                                        foreach (var s in this)
                                        {
                                            foreach (var unt in (from u in s.UnterrichteAusWebuntis where Regex.Replace(u.Fach.Split('|')[0], @"[\d-]", string.Empty) == Regex.Replace(af.Fach, @"[\d-]", string.Empty) select u).ToList())
                                            {
                                                if (unt.Lehrkraft != lehrkraft || unt.Fach != fach)
                                                {
                                                    unt.LeistungW = null;
                                                    unt.LeistungA = null;
                                                    unt.Bemerkung = lehrkraft + " überschreibt die Noten in " + unt.Fach + ".";
                                                }
                                            }
                                        }
                                    }
                                    gehtNichtWeiter = false;

                                    Global.Rückmeldungen.AddRückmeldung(new Rückmeldung(lehrkraft, Regex.Replace(af.Fach, @"[\d-]", string.Empty).PadRight(4), "Die Noten wurden konkurrierend" + (doppelteNoten.Distinct().Count() > 1 ? " und widersprüchlich" : "") + " von mehr als einer Lehrkraft eingetragen. Es werden nur die Noten von " + lehrkraft + " berücksichtigt."));
                                }
                                catch (Exception)
                                {
                                }
                            } while (gehtNichtWeiter);

                            return true;
                        }
                    }
                }
            }

            return false;
        }

        internal void TabelleZeichnen(string interessierendeKlasse, string user)
        {
            try
            {
                do
                {
                    Global.Tabelle = new List<string>();
                    Console.WriteLine((" "));
                    Console.WriteLine((" "));
                    Global.WriteLineTabelle(("Leistungen der Klasse " + interessierendeKlasse + " aus Webuntis & Atlantis: "));
                    Global.WriteLineTabelle(("=====================================================").PadRight(interessierendeKlasse.Length, '='));
                    Global.WriteLineTabelle("*-----*------*----------------------------".PadRight(Global.PadRight + 3, '-') + "*");
                    Global.WriteLineTabelle(("|Name |SuS-Id| Noten + Tendenzen der Klasse " + interessierendeKlasse + " aus Webuntis & F***lantis").PadRight(70) + (DateTime.Now.ToLongDateString() + ", " + DateTime.Now.ToShortTimeString() + ", " + user.PadRight(3).Substring(0, 3) + "|").PadLeft(45));

                    var alleVerschiedenenUnterrichteInDieserKlasseAktuell = new Unterrichte();
                    alleVerschiedenenUnterrichteInDieserKlasseAktuell.AddRange(Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuell.OrderBy(xxx => xxx.Reihenfolge));

                    bool xxxx = false;
                    bool xx = false;
                    bool reliabwähler = false;
                    bool gartenzaun = false;
                    bool prozent = false;

                    var breiteSpalteEins = Math.Max(Global.PadRight - alleVerschiedenenUnterrichteInDieserKlasseAktuell.Count * 5, 10);

                    string x = "*".PadRight(breiteSpalteEins - 8, '-') + "-----------+";
                    string webUntFach = "|".PadRight(breiteSpalteEins - 8, ' ') + "WebUnt-Fach|";
                    string atlantFach = "|".PadRight(breiteSpalteEins - 8, ' ') + "Atlant-Fach|";
                    string l = "|".PadRight(breiteSpalteEins - 8, ' ') + "  Lehrkraft|";
                    string g = "|".PadRight(breiteSpalteEins - 8, ' ') + " Teilnehmer|";
                    string w = "|".PadRight(breiteSpalteEins - 8, ' ') + "WebUn Atlan|";
                    string b = "*".PadRight(breiteSpalteEins - 9, '-') + "-----*------+";
                    string k = "|".PadRight(breiteSpalteEins - 8, ' ') + "            |";
                    string y = "*".PadRight(breiteSpalteEins - 8, '-') + "----*------*";
                    string sj = "|".PadRight(breiteSpalteEins - 8, ' ') + " HZ/JZ + SJ|";

                    foreach (var aktuellerU in alleVerschiedenenUnterrichteInDieserKlasseAktuell)
                    {
                        b += "----+";
                        webUntFach += aktuellerU.Fach.PadRight(4).Substring(0, 4) + "|";
                        atlantFach += (aktuellerU.FachnameAtlantis != null ? aktuellerU.FachnameAtlantis : "").PadRight(4).Substring(0, 4) + "|";
                        l += aktuellerU.Lehrkraft == null ? "    |" : aktuellerU.Lehrkraft.PadRight(4).Substring(0, 4) + "|";
                        g += aktuellerU.KursOderAlle == null ? "geh.|" : aktuellerU.KursOderAlle == "Alle" ? "Alle|" : "Kurs|";
                        w += "W A |";
                        k += "      ";
                        x += "----*";
                        y += "----*";
                        sj += aktuellerU.LeistungA == null || aktuellerU.LeistungA.Schuljahr == null || aktuellerU.LeistungA.HzJz == null ? "    |" : aktuellerU.LeistungA.HzJz + aktuellerU.LeistungA.Schuljahr.Substring(2, 2) + "|";
                    }

                    Global.WriteLineTabelle(x);
                    Global.WriteLineTabelle(webUntFach);
                    Global.WriteLineTabelle(atlantFach);
                    Global.WriteLineTabelle(l);
                    Global.WriteLineTabelle(sj);
                    Global.WriteLineTabelle(g);

                    // Wenn es mehrere Unterrichtsnummern zu einem Unterricht gibt, werden weitere Zeilen hinzugefügt.

                    var anzahlVerschiedenerUnterrichtsnummern = 0;

                    foreach (var item in (from gg in alleVerschiedenenUnterrichteInDieserKlasseAktuell select gg).ToList())
                    {
                        if (item.LessonNumbers != null && item.LessonNumbers.Count > anzahlVerschiedenerUnterrichtsnummern)
                        {
                            anzahlVerschiedenerUnterrichtsnummern = item.LessonNumbers.Count;
                        }
                    }

                    for (int i = 0; i < anzahlVerschiedenerUnterrichtsnummern; i++)
                    {
                        string n = "|".PadRight(breiteSpalteEins - 9, ' ') + "  Unterr.Nr.|";

                        foreach (var aktuellerU in alleVerschiedenenUnterrichteInDieserKlasseAktuell)
                        {
                            n += (aktuellerU.LessonNumbers == null ? "    |" : aktuellerU.LessonNumbers.Count > i ? aktuellerU.LessonNumbers[i].ToString() : "").PadLeft(4).Substring(0, 4) + "|";
                        }
                        Global.WriteLineTabelle(n);
                    }

                    Global.WriteLineTabelle(w);
                    Global.WriteLineTabelle(b);

                    var ii = 1;
                    foreach (var schüler in this)
                    {
                        string s = "|" + (ii.ToString().PadLeft(2) + "." + schüler.Nachname + ", " + schüler.Vorname).PadRight(breiteSpalteEins).Substring(0, breiteSpalteEins - 5) + "|" + schüler.SchlüsselExtern + "|";
                        schüler.Zähler = ii;
                        ii++;
                        for (int i = 0; i < alleVerschiedenenUnterrichteInDieserKlasseAktuell.Count; i++)
                        {
                            Unterricht uA = new Unterricht();
                            bool zugeordnet = false;

                            foreach (var u in schüler.UnterrichteAusWebuntis)
                            {
                                if (u.Fach == alleVerschiedenenUnterrichteInDieserKlasseAktuell[i].Fach)
                                {
                                    if (u.Lehrkraft == alleVerschiedenenUnterrichteInDieserKlasseAktuell[i].Lehrkraft || u.FachnameAtlantis == alleVerschiedenenUnterrichteInDieserKlasseAktuell[i].Fach)
                                    {
                                        if (alleVerschiedenenUnterrichteInDieserKlasseAktuell[i].Gruppe == u.Gruppe)
                                        {
                                            if (alleVerschiedenenUnterrichteInDieserKlasseAktuell[i].Konferenzdatum == u.Konferenzdatum)
                                            {
                                                uA = u;
                                                zugeordnet = true;
                                            }
                                        }
                                    }
                                }
                            }

                            // Wenn der Nachfolger das selbe Fach ist, ändert sich der delimiter

                            var delimiter = "|";

                            if (i < alleVerschiedenenUnterrichteInDieserKlasseAktuell.Count - 1)
                            {
                                if (Regex.Replace(alleVerschiedenenUnterrichteInDieserKlasseAktuell[i].Fach, @"[\d-]", string.Empty) == Regex.Replace(alleVerschiedenenUnterrichteInDieserKlasseAktuell[i + 1].Fach, @"[\d-]", string.Empty))
                                {
                                    delimiter = ":";
                                }
                            }

                            // Wenn es zu dem Fach in der Schülergruppe und bei der Lehrkraft einen Unterricht gibt ... 

                            if (zugeordnet)
                            {
                                if (uA.LeistungA != null)
                                {
                                    if (uA.LeistungW != null)
                                    {
                                        s += (uA.LeistungW.Gesamtnote + uA.LeistungW.Tendenz).PadRight(2) + (uA.LeistungA.Gesamtnote + uA.LeistungA.Tendenz).PadRight(2) + delimiter;

                                        if ((uA.LeistungW.Gesamtnote == null || uA.LeistungW.Gesamtnote == "") && !uA.LeistungA.FachAliases.Contains("REL"))
                                        {
                                            if (!uA.Bemerkung.Contains("eholt"))
                                            {
                                                Global.Rückmeldungen.AddRückmeldung(new Rückmeldung("alle", "", "Sind alle Noten eingetragen?"));
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (uA.LeistungA.FachAliases != null && !uA.LeistungA.FachAliases.Contains("REL"))
                                        {
                                            if (uA.Bemerkung != null)
                                            {
                                                s += "  " + (uA.LeistungA.Gesamtnote + uA.LeistungA.Tendenz).PadRight(2) + delimiter;
                                                if (!uA.Bemerkung.Contains("eholt"))
                                                {
                                                    Global.Rückmeldungen.AddRückmeldung(new Rückmeldung("alle", "", "Sind alle Noten eingetragen?"));
                                                }
                                            }
                                            else
                                            {
                                                // Wenn ein Fach geholt wird, das dieser Schüler aber nicht belegt hat.
                                                s += "XXXX|";
                                            }
                                        }
                                        else
                                        {
                                            s += "  " + (uA.LeistungA.Gesamtnote + uA.LeistungA.Tendenz).PadRight(2) + delimiter;
                                        }
                                    }
                                }
                                else
                                {
                                    if (uA.LeistungW != null)
                                    {
                                        // Wenn der Schüler in Webuntis steht, aber kein Leistungsdatensatz in Atlantis hat
                                        s += (uA.LeistungW.Gesamtnote + uA.LeistungW.Tendenz).PadRight(2) + "$$" + delimiter;
                                        xx = true;
                                    }
                                    else
                                    {
                                        if (uA.Bemerkung != null && uA.Bemerkung.Contains("die Note"))
                                        {
                                            if (uA.Bemerkung.Contains("setzt die Note"))
                                            {
                                                s += (uA.Bemerkung.Split(' ')[0]).PadRight(4, '#') + delimiter;
                                                gartenzaun = true;
                                            }
                                            if (uA.Bemerkung.Contains("überschreibt die Note"))
                                            {
                                                s += (uA.Bemerkung.Split(' ')[0]).PadRight(4, '%') + delimiter;
                                                prozent = true;
                                            }
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

                                if ((new List<string> { "KR", "ER", "REL" }).Contains(alleVerschiedenenUnterrichteInDieserKlasseAktuell[i].Fach))
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

                        Global.WriteLineTabelle(s);
                    }

                    Global.WriteLineTabelle(y);

                    if (reliabwähler)
                    {
                        var le = (from s in this
                                  from u in s.UnterrichteAusWebuntis
                                  where (new List<string> { "KR", "ER", "REL", "KR G1", "KR G2" }).Contains(u.Fach)
                                  select u.Lehrkraft).FirstOrDefault();

                        Global.Rückmeldungen.AddRückmeldung(new Rückmeldung(le, "REL", "Gibt es Reliabwähler? Evtl. Note in Konferenz ergänzen, falls bewertbar. Bei Abgang/Abschluss Strich."));

                        Global.WriteLineTabelle("  * R* Gibt es Reliabwähler? Evtl. Note in Konferenz ergänzen, falls bewertbar. Bei Abgang/Abschluss Strich.");
                    }
                    if (xxxx)
                    {
                        Global.WriteLineTabelle("  * XXXX Der Schüler hat den Kurs nicht belegt.");
                    }
                    if (xx)
                    {
                        Global.WriteLineTabelle("  * $$ Keine Zuordnung zu einem Atlantis-Fach. Abweichende Fächernamen?");
                    }
                    if (gartenzaun)
                    {
                        Global.WriteLineTabelle("  * # Mehrere LuL unterrichten das Fach. Nur einer hat Noten eingetragen. Nur seine Noten werden gedruckt.");
                    }
                    if (prozent)
                    {
                        Global.WriteLineTabelle("  * % Mehrere LuL unterrichten das Fach. Mehrere haben Noten eingetragen. Ihre Entscheidung: Nur seine Noten werden gedruckt.");
                    }

                } while (KorrekturenVornehmen(interessierendeKlasse));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void GetWebuntisUnterrichte(Unterrichte alleUnterrichte, Gruppen alleGruppen, string interessierendeKlasse, string hzJz, List<string> aktSj)
        {
            int i = 0;

            try
            {
                var unterrichteDerKlasse = (from a in alleUnterrichte
                                            where a.Klassen.Split('~').Contains(interessierendeKlasse)
                                            where a.Startdate <= DateTime.Now
                                            where a.Enddate >= DateTime.Now.AddMonths(-2) // Unterrichte, die 2 Monat vor Konferenz beendet wurden, zählen
                                            select a).ToList();


                if (hzJz == "JZ")
                {
                    // Im Jahreszeugnis der Unter- und Mittelstufen der Anlage A kann es sein, dass in einem
                    // Unterricht des 1.Hj Notes erteilt wurden, die aber noch nicht in F***lantis stehen.

                    var unterrichteDesErstenHj = (from a in alleUnterrichte
                                                  where a.Klassen.Split(',').Contains(interessierendeKlasse)
                                                  where a.Startdate >= new DateTime(Convert.ToInt32(aktSj[0]), 08, 01)
                                                  where a.Enddate <= new DateTime(Convert.ToInt32(aktSj[1]) + 2000, 02, 1)
                                                  where !(from u in unterrichteDerKlasse where u.Fach == a.Fach select u).Any()
                                                  select a).ToList();

                    if (unterrichteDesErstenHj.Count > 0)
                    {
                        Console.WriteLine("Der Unterricht " + (unterrichteDesErstenHj.Count == 1 ? "im Fach " : "in den Fächern ") + Global.List2String(unterrichteDesErstenHj.Select(x => x.Fach).ToList(), ",") + " ist im Halbjahr geendet.");
                        Global.Rückmeldungen.AddRückmeldung(new Rückmeldung("Klassenleitung", "", "Sind evtl. Kurse zum Halbjahr ausgelaufen, deren Noten in das Jahreszeugnis übernommen werden müssen?"));
                    }

                    unterrichteDerKlasse.AddRange(unterrichteDesErstenHj);
                }

                foreach (var schüler in this)
                {
                    i += schüler.GetUnterrichte(unterrichteDerKlasse, alleGruppen);
                }
                Console.WriteLine(("Alle Unterrichte der Klasse " + interessierendeKlasse + " aus Webuntis ").PadRight(Global.PadRight - 2, '.') + i.ToString().PadLeft(6));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}