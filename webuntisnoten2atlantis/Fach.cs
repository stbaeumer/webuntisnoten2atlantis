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

        public Fach(string klasse, string fach, string lehrkraft, DateTime konferenzdatum)
        {
            this.Klasse = klasse;
            this.Name = fach;
            this.Lehrkraft = lehrkraft;
            this.Konferenzdatum = konferenzdatum;
            this.NameAliase = GetFachAliases(fach);
        }

        private List<string> GetFachAliases(string fach)
        {
            var nameAliases = new List<string>();

            if (fach == "PKG" || fach == "PK")
            {
                nameAliases.Add("PK");
                nameAliases.Add("PKG");
            }
            if (fach == "N" || fach == "NB1" || fach == "NB2" || fach == "NA1" || fach == "NA2")
            {
                nameAliases.Add("N");
                nameAliases.Add("NB1");
                nameAliases.Add("NB2");
                nameAliases.Add("NA1");
                nameAliases.Add("NA2");
            }
            if (fach == "S" || fach == "SB1" || fach == "SB2" || fach == "SA1" || fach == "SA2")
            {
                nameAliases.Add("S");
                nameAliases.Add("SB1");
                nameAliases.Add("SB2");
                nameAliases.Add("SA1");
                nameAliases.Add("SA2");
            }

            if (fach == "E" || fach == "EB1" || fach == "EB2" || fach == "EA1" || fach == "EA2")
            {
                nameAliases.Add("E");
                nameAliases.Add("EB1");
                nameAliases.Add("EB2");
                nameAliases.Add("EA1");
                nameAliases.Add("EA2");
            }

            if (fach == "KR" || fach == "ER" || fach == "REL" || fach == "KR " || fach == "ER ")
            {
                nameAliases.Add("KR");
                nameAliases.Add("ER");
                nameAliases.Add("REL");
                nameAliases.Add("KR ");
                nameAliases.Add("ER ");
            }

            return nameAliases;
        }
    }
}