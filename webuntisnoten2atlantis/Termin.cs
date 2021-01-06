// Published under the terms of GPLv3 Stefan Bäumer 2020.

using System;

namespace webuntisnoten2atlantis
{
    internal class Termin
    {
        public Termin()
        {
        }

        public string Bildungsgang { get; internal set; }
        public string Klasse { get; internal set; }
        public DateTime Uhrzeit { get; internal set; }
        public string Raum { get; internal set; }
        public Lehrers Lehrers { get; internal set; }
    }
}