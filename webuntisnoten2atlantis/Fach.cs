﻿// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;

namespace webuntisnoten2atlantis
{
    public class Fach
    {
        public string Klasse { get; private set; }
        public string Name { get; private set; }
        public string Lehrkraft { get; private set; }
        public DateTime Konferenzdatum { get; }

        public Fach(string klasse, string fach, string lehrkraft, DateTime konferenzdatum)
        {
            this.Klasse = klasse;
            this.Name = fach;
            this.Lehrkraft = lehrkraft;
            this.Konferenzdatum = konferenzdatum;
        }
    }
}