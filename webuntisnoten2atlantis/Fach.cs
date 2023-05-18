// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;

namespace webuntisnoten2atlantis
{
    public class Fach
    {
        public string Klasse { get; private set; }
        public string Name { get; private set; }
        public string Lehrkraft { get; private set; }
        public List<string> NameAliase { get; private set; }
        public DateTime Konferenzdatum { get; }
        public string Gruppe { get; }

        public Fach(string klasse, string fach, string lehrkraft, DateTime konferenzdatum, string gruppe)
        {
            this.Klasse = klasse;
            this.Name = fach;
            this.Lehrkraft = lehrkraft;
            this.Konferenzdatum = konferenzdatum;
            this.Gruppe = gruppe;
        }

        public Fach(string name, string lehrkraft)
        {
            this.Name = name;
            Lehrkraft = lehrkraft;
        }
    }
}