// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace webuntisnoten2atlantis
{
    public class Schüler
    {
        public int StudentZeile { get; internal set; }
        public string Klasse { get; internal set; }
        public DateTime EintrittInKlasse { get; internal set; }
        public DateTime AustrittAusKlasse { get; internal set; }
        public string Vorname { get; internal set; }
        public string Nachname { get; internal set; }
        public Unterrichte UnterrichteAusWebuntis { get; set; }
        public Unterrichte UnterrichteGeholt { get; set; }
        public int SchlüsselExtern { get; internal set; }
        /// <summary>
        /// Das sind die Unterrichte, die aktuell in Webuntis nicht unterricht werden, aber in Atlantis existieren. Die werden gezogen, damit dort geholte Noten aus der Vergangenheit hineingeschrieben werden können.
        /// </summary>
        public Unterrichte UnterrichteAktuellAusAtlantis { get; internal set; }
        public Unterrichte AlleVerschiedenenUnterrichteInDieserKlasseAktuellUnsortiert { get; private set; }

        internal void ErstelleUnterricht(Leistungen atlantisLeistungen)
        {
            foreach (var u in UnterrichteAusWebuntis)
            {
                // Wenn einer Webuntis-Leistung keine Atlantis-Leistung zugeordnet werden konnte, ...

                if (u.LeistungA == null)
                {
                    // ... muss trotzdem ein Unterricht der Klasse angelegt werden.
                    // Alle infragekommenden Atlantisleistungen werden für eine späetere Auswahl angehangen.

                    foreach (var iA in (from al in atlantisLeistungen.OrderByDescending(x => x.Konferenzdatum)
                                        where al.SchlüsselExtern == SchlüsselExtern
                                        //where al.HzJz == hzJz
                                        //where al.Schuljahr == aktSj[0] + "/" + aktSj[1]
                                        where al.Konferenzdatum.Date > DateTime.Now.Date || al.Konferenzdatum.Year == 1
                                        select al).ToList())
                    {
                        // Nur Fächer, die nicht bereits erfolgreich zugeordnet wurden, werden als Infragekommende Fächer
                        // für eine spätere manuelle Auswahl erfasst.

                        if (!(from a in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuellUnsortiert
                              where a.FachnameAtlantis == iA.Fach
                              select a).Any())
                        {
                            u.InfragekommendeLeistungenA.Add(iA);
                        }
                    }

                    Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuellUnsortiert.AddUnterrichte(u);
                }
            }
        }

        /// <summary>
        /// Alle Atlantisleistungen, deren Fach nicht zu einem aktuellen Unterricht passt und deren Konferenzdatum in der Vergangenheit liegt, 
        /// sind geholte Unterrichte.
        /// </summary>
        /// <param name="atlantisLeistungen"></param>
        internal void GeholteUnterrichteHinzufügen(Leistungen atlantisLeistungen)
        {
            UnterrichteGeholt = new Unterrichte();

            foreach (var atlantisLeistung in (from al in atlantisLeistungen.OrderByDescending(x => x.Konferenzdatum)
                                              where al.SchlüsselExtern == SchlüsselExtern
                                              where al.Konferenzdatum < DateTime.Now.AddDays(-30)
                                              where al.Gesamtnote != null
                                              where al.Gesamtnote != ""
                                              select al).ToList())
            {
                bool geholt = true;
                               
                foreach (var atlantisFach in atlantisLeistung.FachAliases)
                {
                    if ((from  g in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuellUnsortiert select g.Fach).Contains(Regex.Replace(atlantisFach, @"[\d-]", string.Empty)))
                    {
                        geholt = false;
                        break;
                    }

                    // Wenn das Fach bereits geholt wurde, wird es nicht nochmal geholt

                    if ((from u in UnterrichteGeholt where u.Fach == atlantisFach select u).Any())
                    {
                        geholt = false;
                        break;
                    }
                }
                if (geholt)
                {
                    UnterrichteGeholt.Add(new Unterricht(atlantisLeistung));
                }
            }
        }

        /// <summary>
        /// Zu allen Webuntisleistungen aller Schüler wird je eine Atlantisleistung des aktuellen Abschnitts (SJ,HzJz) zugeordnet.
        /// </summary>
        /// <param name="atlantisLeistungen"></param>
        /// <param name="hzJz"></param>
        /// <param name="aktSj"></param>
        internal void GetAtlantisLeistungen(Leistungen atlantisLeistungen, string hzJz, List<string> aktSj)
        {
            foreach (var u in UnterrichteAusWebuntis)
            {
                u.GetAtlantisLeistung(atlantisLeistungen, SchlüsselExtern, hzJz, aktSj);
            }
        }

        internal void GetUnterrichte(List<Unterricht> unterrichteDerKlasse, List<Gruppe> alleGruppen)
        {
            UnterrichteAusWebuntis = new Unterrichte();

            // Unterrichte der ganzen Klasse

            var unterrichteDerKlasseOhneGruppen = (from a in unterrichteDerKlasse
                                                   where a.Startdate <= DateTime.Now
                                                   where a.Enddate >= DateTime.Now.AddMonths(-2) // Unterrichte, die 2 Monat vor Konferenz beendet wurden, zählen
                                                   where a.Gruppe == ""
                                                   select a).ToList();

            foreach (var u in unterrichteDerKlasseOhneGruppen)
            {
                // Wenn ein Lehrer zweimal mit dem selben Fach in Webuntis eingetragen ist, wird kein weiterer Unterricht angelegt.

                var gibtsSchon = (from x in UnterrichteAusWebuntis where x.Fach == u.Fach where x.Lehrkraft == u.Lehrkraft select x).FirstOrDefault();

                if (gibtsSchon == null)
                {
                    UnterrichteAusWebuntis.Add(new Unterricht(
                        u.LeistungA,
                        u.LessonNumbers[0],
                        u.Fach,
                        u.LeistungW,
                        u.Lehrkraft,
                        u.Zeile,
                        u.Periode,
                        u.Gruppe,
                        u.Klassen,
                        u.Startdate,
                        u.Enddate));
                }
                else
                {
                    gibtsSchon.LessonNumbers.Add(u.LessonNumbers[0]);
                }
            }

            // Kurse

            foreach (var gruppe in (from g in alleGruppen where g.StudentId == SchlüsselExtern select g).ToList())
            {
                var u = (from a in unterrichteDerKlasse where a.Gruppe == gruppe.Gruppenname select a).FirstOrDefault();

                // u ist z.B. null, wenn ein Kurs in der ExportLessons als (lange) abgeschlossen steht.

                if (u != null)
                {
                    // Wenn ein Lehrer zweimal mit dem selben Fach in Webuntis eingetragen ist, wird kein weiterer Unterricht angelegt.

                    var gibtsSchon = (from x in UnterrichteAusWebuntis where x.Fach == u.Fach where x.Lehrkraft == u.Lehrkraft select x).FirstOrDefault();

                    if (gibtsSchon == null)
                    {
                        UnterrichteAusWebuntis.Add(new Unterricht(
                            u.LeistungA,
                            u.LessonNumbers[0],
                            u.Fach,
                            u.LeistungW,
                            u.Lehrkraft,
                            u.Zeile,
                            u.Periode,
                            u.Gruppe,
                            u.Klassen,
                            u.Startdate,
                            u.Enddate));
                    }
                    else
                    {
                        gibtsSchon.LessonNumbers.Add(u.LessonNumbers[0]);
                    }
                }
            }
        }

        internal void HoleLeistungen(List<string> dieseFächerHolen)
        {
            for (int i = this.UnterrichteGeholt.Count - 1; i >= 0; i--)
            {
                bool holen = false;

                foreach (var dF in dieseFächerHolen)
                {
                    if (this.UnterrichteGeholt[i].Fach == dF || this.UnterrichteGeholt[i].FachnameAtlantis == dF)
                    {
                        holen = true; break;
                    }
                }
                if (holen)
                {
                    var x = (from uuu in UnterrichteAktuellAusAtlantis 
                             where uuu.Fach == this.UnterrichteGeholt[i].Fach || uuu.Fach == this.UnterrichteGeholt[i].FachnameAtlantis
                             select uuu.LeistungA).FirstOrDefault();

                    this.UnterrichteGeholt[i].Bemerkung = "Note geholt.";

                    // Die Atlantis-LeistungsID wird mit der ID des aktuellen Atlantis-Unterrichts überschrieben.

                    this.UnterrichteGeholt[i].LeistungA.Bemerkung = this.UnterrichteGeholt[i].LeistungA.LeistungId + ">" + x.LeistungId + "|";
                    this.UnterrichteGeholt[i].LeistungA.LeistungId = x.LeistungId;
                }
                else
                {
                    this.UnterrichteGeholt.RemoveAt(i);
                }
            }
        }

        /// Wenn einer Webuntis-Leistung keine Atlantis-Leistung zugeordnet werden konnte, muss trotzdem ein Unterricht der Klasse angelegt werden.
        /// Alle infragekommenden aktuellen Atlantisleistungen werden für eine spätere Auswahl angehangen.
        internal void InfragekommendeUnterrichteFürSpätereZuordnungAnlegen(Leistungen atlantisLeistungen, string hzJz, List<string> aktSj)
        {
            foreach (var u in UnterrichteAusWebuntis)
            {
                u.InfragekommendeAktuelleLeistungenHinzufügen(atlantisLeistungen, SchlüsselExtern, hzJz, aktSj);
            }
        }

        internal void QueriesBauenUpdate()
        {
            int i = 0;

            foreach (var u in UnterrichteAusWebuntis.OrderBy(x=>x.Reihenfolge))
            {
                if (u.LeistungA != null)
                {
                    if (u.LeistungW != null)
                    {
                        u.QueryBauen();

                        if (u.LeistungW.Beschreibung != null && u.LeistungW.Query != null)
                        {
                            u.LeistungW.Update();
                            i++;
                        }
                    }
                }
            }
            if (i == 0)
            {
                Global.WriteLine("A C H T U N G: Es wurde keine einzige Leistung angelegt oder verändert. Das kann daran liegen, dass das Notenblatt in Atlantis nicht richtig angelegt ist. Die Tabelle zum manuellen eintragen der Noten muss in der Notensammelerfassung sichtbar sein.");
            }
            foreach (var unterricht in UnterrichteGeholt.OrderBy(x => x.Reihenfolge))
            {
                if (unterricht.LeistungA != null)
                {
                    unterricht.LeistungW = new Leistung(unterricht.LeistungA, "geholt|");
                    unterricht.LeistungA.Gesamtpunkte = null;
                    unterricht.LeistungA.Gesamtnote = null;
                    unterricht.LeistungA.Tendenz = null;
                    unterricht.QueryBauen();
                    unterricht.LeistungW.Update();
                }
            }
        }

        /// <summary>
        /// Alle Atlantisleistungen des aktuellen Abschnitts, die nicht zu den aktuellen Webuntis-Unterrichten gehören,
        /// werden zu neuen Unterrichten, denen später die geholten Noten zugeordnet werden können.  
        /// </summary>
        /// <param name="atlantisLeistungen"></param>
        internal void UnterrichteAktuellAusAtlantisHolen(Leistungen atlantisLeistungen, string hzJz, List<string> aktSj)
        {
            UnterrichteAktuellAusAtlantis = new Unterrichte();

            foreach (var atlantisLeistung in (from al in atlantisLeistungen.OrderByDescending(x => x.Konferenzdatum)
                                              where al.SchlüsselExtern == SchlüsselExtern
                                              where al.Konferenzdatum >= DateTime.Now.Date || al.Konferenzdatum.Year == 1 
                                              where al.Schuljahr == aktSj[0] + "/" + aktSj[1]
                                              where al.HzJz == hzJz
                                              select al).ToList())
            {
                UnterrichteAktuellAusAtlantis.Add(new Unterricht(atlantisLeistung));
            }
        }
    }
}