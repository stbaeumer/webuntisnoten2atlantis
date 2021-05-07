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
        /// Die Gesamtnote ist ein string, weil auch ein '-' dort stehen kann. Wenn keine Note erteilt wurde ist der Wert null. Zulässige Werte: 1,2,3,4,5,6, NULL, A, -, 
        /// </summary>
        public string Gesamtnote { get; internal set; }
        /// <summary>
        /// Wenn keine Punkte erteilt wurden, ist die Punktzahl null. Zuläassige Werte sind 0,1,2,...14,15,84(=Attest),99(='-')
        /// </summary>
        public string Gesamtpunkte { get; internal set; }
        public string Bemerkung { get; internal set; }
        
        /// <summary>
        /// Jede Leistung in Webuntis wird von einer Lehrkraft eingetragen. 
        /// Nur Lehrkräfte, die eine Leistung eintragen, bekommen einen Termin für die Zeugniskonferenz gesetzt.
        /// Es ist wichtig, dass Lehrkräfte ihre Noten selbst eintragen. Wenn der Admin das für die Lehrkraft übernimt, wird er zur Lehrkraft. 
        /// </summary>
        public string Lehrkraft { get; internal set; }
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
        public bool HatBemerkung { get; internal set; }
        /// <summary>
        /// Wenn Schüler der Anlage A die Zeugnisart A01AS gesetzt haben, dann werden für sie alte Noten geholt.
        /// </summary>
        public string Zeugnisart { get; internal set; }
        public string Prüfungsart { get; internal set; }
        public DateTime Von { get; internal set; }
        public DateTime Bis { get; internal set; }
        public string FachUntis { get; internal set; }

        public bool IstAbschlussklasse()
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
            

            if ((Gliederung.StartsWith("C03") || Gliederung.StartsWith("C13")) && Jahrgang == 2 && DateTime.Now.Month > 3)
            {
                return true;
            }

            if ((Gliederung.StartsWith("B08") || Gliederung.StartsWith("C06") || this.Gliederung.StartsWith("C08")) && DateTime.Now.Month > 3)
            {
                return true;
            }


            return false;
        }

        internal string Gesamtnote2Gesamtpunkte(string gesamtnote)
        {
            if (gesamtnote.EndsWith("-"))
            {
                gesamtnote = gesamtnote.Replace("-", "");

                if (gesamtnote == "1")
                {
                    return "13";
                }
                if (gesamtnote == "2")
                {
                    return "10";
                }
                if (gesamtnote == "3")
                {
                    return "7";
                }
                if (gesamtnote == "4")
                {
                    return "4";
                }
                if (gesamtnote == "5")
                {
                    return "1";
                }
            }

            if (gesamtnote.EndsWith("+"))
            {
                gesamtnote = gesamtnote.Replace("+", "");
                if (gesamtnote == "1")
                {
                    return "15";
                }
                if (gesamtnote == "2")
                {
                    return "12";
                }
                if (gesamtnote == "3")
                {
                    return "9";
                }
                if (gesamtnote == "4")
                {
                    return "6";
                }
                if (gesamtnote == "5")
                {
                    return "3";
                }
            }

            if (gesamtnote == "1")
            {
                return "14";
            }
            if (gesamtnote == "2")
            {
                return "11";
            }
            if (gesamtnote == "3")
            {
                return "8";
            }
            if (gesamtnote == "4")
            {
                return "5";
            }
            if (gesamtnote == "5")
            {
                return "2";
            }

            return null;
        }

        internal string Gesamtnote2Tendenz(string gesamtnote)
        {
            if (gesamtnote == null)
            {
                return null;
            }
            if (gesamtnote.EndsWith("-"))
            {
                return "-";
            }

            if (gesamtnote.EndsWith("+"))
            {
                return "+";
            }

            return null;
        }

    }
}