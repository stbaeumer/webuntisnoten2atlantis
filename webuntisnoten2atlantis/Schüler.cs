// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Linq;

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
        public Unterrichte UnterrichteAktuell { get; set; }
        public Unterrichte UnterrichteGeholt { get; set; }
        public int SchlüsselExtern { get; internal set; }

        internal void GetUnterrichte(List<Unterricht> alleUnterrichte, List<Gruppe> alleGruppen)
        {
            this.UnterrichteAktuell = new Unterrichte();

            // Unterrichte der ganzen Klasse

            var unterrichteDerKlasse = (from a in alleUnterrichte                                        
                                        where a.Startdate <= DateTime.Now
                                        where a.Enddate >= DateTime.Now.AddMonths(-2) // Unterrichte, die 2 Monat vor Konferenz beendet wurden, zählen
                                        where a.Gruppe == ""
                                        select a).ToList();

            foreach (var u in unterrichteDerKlasse)
            {
                UnterrichteAktuell.Add(new Unterricht(
                    u.AtlantisLeistung,
                    u.Fach,
                    u.WebuntisLeistung,
                    u.Lehrkraft,
                    u.MarksPerLessonZeile,
                    u.Periode,
                    u.Gruppe,
                    u.Klassen,
                    u.Startdate,
                    u.Enddate));
            }            

            // Kurse

            foreach (var gruppe in (from g in alleGruppen where g.StudentId == SchlüsselExtern select g).ToList())
            {
                var u = (from a in alleUnterrichte where a.Gruppe == gruppe.Gruppenname select a).FirstOrDefault();

                // u ist z.B. null, wenn ein Kurs in der ExportLessons als (lange) abgeschlossen steht.

                if (u != null)
                {
                    UnterrichteAktuell.Add(new Unterricht(
                    u.AtlantisLeistung,
                    u.Fach,
                    u.WebuntisLeistung,
                    u.Lehrkraft,
                    u.MarksPerLessonZeile,
                    u.Periode,
                    u.Gruppe,
                    u.Klassen,
                    u.Startdate,
                    u.Enddate));
                }                
            }
        }

        internal bool ZuordnungZuAktuellemUnterrichtMöglich(Leistung atlantisLeistung)
        {
            foreach (var u in this.UnterrichteAktuell)
            {
                if (
                    atlantisLeistung.FachAliases.Contains(u.Fach) && 
                    atlantisLeistung.SchlüsselExtern == this.SchlüsselExtern &&
                    atlantisLeistung.Klasse.Substring(0,2) == this.Klasse.Substring(0,2)
                    )
                {
                    u.AtlantisLeistung = atlantisLeistung;
                    return true;
                }
            }
            return false;
        }
    }
}