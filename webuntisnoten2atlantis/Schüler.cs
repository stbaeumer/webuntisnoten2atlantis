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
        public Unterrichte UnterrichteAusWebuntis { get; set; }
        public Unterrichte UnterrichteGeholt { get; set; }
        public int SchlüsselExtern { get; internal set; }
        /// <summary>
        /// Das sind die Unterrichte, die aktuell in Webuntis nicht unterricht werden, aber in Atlantis existieren. Die werden gezogen, damit dort geholte Noten aus der Vergangenheit hineingeschrieben werden können.
        /// </summary>
        public Unterrichte UnterrichteAktuellAusAtlantis { get; internal set; }

        internal void GetUnterrichte(List<Unterricht> alleUnterrichte, List<Gruppe> alleGruppen)
        {
            this.UnterrichteAusWebuntis = new Unterrichte();

            // Unterrichte der ganzen Klasse

            var unterrichteDerKlasse = (from a in alleUnterrichte
                                        where a.Startdate <= DateTime.Now
                                        where a.Enddate >= DateTime.Now.AddMonths(-2) // Unterrichte, die 2 Monat vor Konferenz beendet wurden, zählen
                                        where a.Gruppe == ""
                                        select a).ToList();

            foreach (var u in unterrichteDerKlasse)
            {
                UnterrichteAusWebuntis.Add(new Unterricht(
                    u.AL,
                    u.LessonNumber,
                    u.Fach,
                    u.WL,
                    u.Lehrkraft,
                    u.Zeile,
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
                    UnterrichteAusWebuntis.Add(new Unterricht(
                    u.AL,
                    u.LessonNumber,
                    u.Fach,
                    u.WL,
                    u.Lehrkraft,
                    u.Zeile,
                    u.Periode,
                    u.Gruppe,
                    u.Klassen,
                    u.Startdate,
                    u.Enddate));
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
                    if (this.UnterrichteGeholt[i].Fach == dF.Split('|')[0] &&
                        this.UnterrichteGeholt[i].AL.Konferenzdatum.ToShortDateString() == dF.Split(' ')[1])
                    {
                        holen = true; break;
                    }
                }
                if (holen)
                {
                    var x = (from uuu in UnterrichteAktuellAusAtlantis where uuu.Fach == this.UnterrichteGeholt[i].Fach select uuu.AL).FirstOrDefault();

                    this.UnterrichteGeholt[i].Bemerkung = "Note geholt.";

                    // Die Atlantis-LeistungsID wird mit der ID des aktuellen Atlantis-Unterrichts überschrieben.

                    this.UnterrichteGeholt[i].AL.Bemerkung = "|" + this.UnterrichteGeholt[i].AL.LeistungId + ">" + x.LeistungId;
                    this.UnterrichteGeholt[i].AL.LeistungId = x.LeistungId;
                }
                else
                {
                    this.UnterrichteGeholt.RemoveAt(i);
                }
            }
        }
    }
}