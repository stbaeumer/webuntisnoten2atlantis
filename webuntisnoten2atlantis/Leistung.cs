// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;

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
        public string Nachname { get; internal set; }
        public string Vorname { get; internal set; }
        public DateTime Geburtsdatum { get; internal set; }
        public bool Volljährig { get; internal set; }
        public string Bereich { get; internal set; }
        public string Gesamtpunkte_12_1 { get; internal set; }
        public string Gesamtpunkte_12_2 { get; internal set; }
        public string Gesamtpunkte_13_1 { get; internal set; }
        public string Gesamtpunkte_13_2 { get; internal set; }
        public int LehrkraftAtlantisId { get; internal set; }
        public string Zielfach { get; internal set; }
        public int ZielLeistungId { get; internal set; }

        public bool IstAbschlussklasse()
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
                
        internal void Zuordnen(List<Leistung> aL, string beschreibung)
        {
            if (aL.Count > 0)
            {
                // Falls Neu oder Update

                if ((aL[0].Gesamtnote == null && aL[0].Gesamtpunkte == null && aL[0].Tendenz == null) || (aL[0].Gesamtnote != Gesamtnote || aL[0].Tendenz != Tendenz))
                {
                    // Falls Neu

                    if (aL[0].Gesamtnote == null && aL[0].Gesamtpunkte == null && aL[0].Tendenz == null)
                    {
                        this.Beschreibung = "NEU;" + this.Beschreibung;
                    }
                    else // Falls Update
                    {
                        if (aL[0].Gesamtnote != Gesamtnote || aL[0].Tendenz != Tendenz)
                        {
                            this.Beschreibung = "UPD:" + this.Beschreibung;

                            if (aL[0].Gesamtnote != Gesamtnote)
                            {
                                this.Beschreibung = (aL[0].Gesamtnote == null ? "null" : aL[0].Gesamtnote) + "->" + Gesamtnote + ";" + this.Beschreibung;
                            }
                            if (aL[0].Tendenz != Tendenz)
                            {
                                this.Beschreibung = (aL[0].Tendenz == null ? "null" : aL[0].Tendenz) + "->" + Tendenz + ";" + this.Beschreibung;
                            }
                        }
                    }
                    this.Zielfach = aL[0].Fach;
                    this.ZielLeistungId = aL[0].LeistungId;
                    this.Beschreibung += beschreibung;
                    this.LehrkraftAtlantisId = 3844; // DBA
                }
                else
                {
                    this.Beschreibung += "keine Änderung";
                }
            }            
        }
    }
}