// Published under the terms of GPLv3 Stefan Bäumer 2023.

namespace webuntisnoten2atlantis
{
    public class Rückmeldung
    {
        public string Lehrkraft { get; private set; }
        public string Fach { get; }
        public string Meldung { get; private set; }

        public Rückmeldung(string lehrkraft, string fach, string meldung)
        {
            this.Lehrkraft = lehrkraft;
            this.Fach = fach;
            this.Meldung = meldung;        
        }

        public override string ToString() 
        {
            return " * " + Lehrkraft.PadRight(3) + ", " + Fach.PadRight(4).Substring(0,4) + ": " + Meldung + "\n";
        }
    }
}