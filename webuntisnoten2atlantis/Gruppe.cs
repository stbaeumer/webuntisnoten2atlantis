// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;

namespace webuntisnoten2atlantis
{
    public class Gruppe
    {
        public Gruppe()
        {
        }

        public int MarksPerLessonZeile { get; internal set; }
        public int StudentId { get; internal set; }
        public string Gruppenname { get; internal set; }
        public string Fach { get; internal set; }
        public DateTime Startdate { get; internal set; }
        public DateTime Enddate { get; internal set; }
    }
}