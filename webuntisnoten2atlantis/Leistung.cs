// Published under the terms of GPLv3 Stefan Bäumer 2020.

using System;

namespace webuntisnoten2atlantis
{
    public class Leistung
    {
        public DateTime Datum { get; internal set; }
        public string Name { get; internal set; }
        public string Klasse { get; internal set; }
        public string Fach { get; internal set; }
        /// <summary>
        /// Die Gesamtnote ist ein string, weil auch ein '-' dort stehen kann.
        /// Wenn keine Note erteilt wurde ist der Wert null.
        /// Zulässige Werte: 1,2,3,4,5,6, NULL, A, -, 
        /// </summary>
        public string Gesamtnote { get; internal set; }
        /// <summary>
        /// Wenn keine Punkte erteilt wurden, ist die Punktzahl null
        /// Zuläassige Werte: 0,1,2,...14,15,99(='-')
        /// </summary>
        public string Gesamtpunkte { get; internal set; }
        public string Bemerkung { get; internal set; }
        public string Benutzer { get; internal set; }
        public int SchlüsselExtern { get; internal set; }
        public int LeistungId { get; internal set; }
        public bool ReligionAbgewählt { get; internal set; }
        public string HzJz { get; internal set; }
        public string Anlage { get; internal set; }
        public string Zeugnistext { get; internal set; }
        public string EinheitNP { get; internal set; }
        public bool SchuelerAktivInDieserKlasse { get; internal set; }
        public DateTime Konferenzdatum { get; internal set; }
        public string Tendenz { get; internal set; }
        public int Jahrgang { get; internal set; }
        public string Schuljahr { get; internal set; }
        public string Gliederung { get; internal set; }
        public bool Abschlussklasse { get; internal set; }
        public string Beschreibung { get; internal set; }
        public bool GeholteNote { get; internal set; }

        internal bool IstAbschlussklasse()
        {
            if (Anlage.StartsWith(Properties.Settings.Default.Klassenart))
            {
                // Klassen im Jahrgang 4 sind immer Abschlussklasse

                if (Jahrgang == 4)
                {
                    return true;
                }

                if (Jahrgang == 3 && !Klasse.StartsWith("M") && !Klasse.StartsWith("E") && HzJz == "Jz")
                {
                    return true;
                }
            }
            
            return false;
        }
    }
}