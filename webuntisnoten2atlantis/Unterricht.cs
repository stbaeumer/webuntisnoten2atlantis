// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;

namespace webuntisnoten2atlantis
{
    public class Unterricht
    {
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
    }
}