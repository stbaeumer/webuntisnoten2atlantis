// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;

namespace webuntisnoten2atlantis
{
    public class Unterricht
    {
        public Unterricht()
        {
        }

        public Unterricht(Leistung atlantisLeistung)
        {
            AtlantisLeistung = new Leistung();
            AtlantisLeistung = atlantisLeistung;            
        }

        public Unterricht(Leistung atlantisLeistung, string fach, Leistung webuntisLeistung, string lehrkraft, int marksPerLessonZeile, int periode, string gruppe, string klassen, DateTime startdate, DateTime enddate) : this(atlantisLeistung)
        {
            this.AtlantisLeistung = atlantisLeistung;
            this.Fach = fach;
            this.WebuntisLeistung = webuntisLeistung;
            this.Lehrkraft = lehrkraft;
            this.MarksPerLessonZeile = marksPerLessonZeile;
            this.Periode = periode;
            this.Gruppe = gruppe;
            this.Klassen = klassen;
            this.Startdate = startdate;
            this.Enddate = enddate;
        }

        public int MarksPerLessonZeile { get; internal set; }
        public int LessonId { get; internal set; }
        public int LessonNumber { get; internal set; }
        public string Fach { get; internal set; }
        public string Lehrkraft { get; internal set; }
        public string Klassen { get; internal set; }
        public string Gruppe { get; internal set; }
        public int Periode { get; internal set; }
        public DateTime Startdate { get; internal set; }
        public DateTime Enddate { get; internal set; }
        public Leistung WebuntisLeistung { get; internal set; }
        public Leistung AtlantisLeistung { get; internal set; }
        public int Reihenfolge { get; internal set; }
    }
}