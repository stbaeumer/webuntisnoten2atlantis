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
        public string Query { get; set; }
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
            // Nur, wenn es ein korrespondierendes Atlantis-Fach gibt

            if (aL.Count > 0)
            {
                this.Beschreibung = (aL[0].Nachname.PadRight(10)).Substring(0, 6) + " " + (aL[0].Vorname.PadRight(10)).Substring(0, 2) + "|" + aL[0].Fach.PadRight(5) + "|" + beschreibung;

                // Falls Neu oder Update

                if (
                    (aL[0].Gesamtnote == null && aL[0].Gesamtpunkte == null && aL[0].Tendenz == null)                                  // Neu
                    || (aL[0].Gesamtnote != Gesamtnote                                                                                 // UPD 
                    || (aL[0].Gesamtpunkte != Gesamtpunkte && aL[0].EinheitNP == "P") || (aL[0].Tendenz != Tendenz && aL[0].EinheitNP =="P")))     // UPD Gym
                {
                    this.Query = "UPDATE noten_einzel SET ";
                    
                    // Falls Neu

                    if (aL[0].Gesamtnote == null && aL[0].Gesamtpunkte == null && aL[0].Tendenz == null)
                    {
                        this.Beschreibung = "NEU|" + this.Beschreibung;

                        if (Gesamtpunkte != null && aL[0].EinheitNP == "P")
                        {
                            this.Beschreibung = this.Beschreibung + ("Punkte:" + (aL[0].Gesamtpunkte == null ? "null" : aL[0].Gesamtpunkte) + "->" + Gesamtpunkte).PadRight(15) + "|";
                            this.Query += "punkte=" + ("'" + Gesamtpunkte).PadLeft(3) +"'" + ", ";
                        }
                        else
                        {
                            this.Query += (" ").PadRight(14) + "  ";
                        }
                        if (Gesamtnote != null)
                        {
                            this.Beschreibung = this.Beschreibung + ("Note:" + (aL[0].Gesamtnote == null ? "null" : aL[0].Gesamtnote) + "->" + Gesamtnote).PadRight(12) + "|";
                            this.Query += "s_note='" + Gesamtnote +"'" + ", ";
                        }
                        else
                        {
                            this.Query += (" ").PadRight(10) + "  ";
                        }
                        if (Tendenz != null && aL[0].EinheitNP == "P")
                        {
                            this.Beschreibung = this.Beschreibung + ("Tendenz:" + (aL[0].Tendenz == null ? "null" : aL[0].Tendenz) + "->" + Tendenz).PadRight(15) + "|";
                            this.Query += "s_tendenz='" + Tendenz +"'" + ", ";
                        }
                        else
                        {
                            this.Query += (" ").PadRight(13) + "  ";
                        }
                    }
                    else // Falls Update
                    {
                        if (aL[0].Gesamtnote != Gesamtnote || (aL[0].Gesamtpunkte != Gesamtpunkte && EinheitNP=="P") || (aL[0].Tendenz != Tendenz && EinheitNP == "P"))
                        {
                            this.Beschreibung = "UPD|" + this.Beschreibung;

                            if (aL[0].Gesamtpunkte != Gesamtpunkte && aL[0].EinheitNP == "P")
                            {
                                this.Beschreibung = this.Beschreibung + ("Punkte:" + (aL[0].Gesamtpunkte == null ? "null" : aL[0].Gesamtpunkte) + "->" + Gesamtpunkte).PadRight(15) + "|";
                                this.Query += "punkte=" + ("'" + Gesamtpunkte).PadLeft(3) + "'" + ", ";
                            }
                            else
                            {
                                this.Query += (" ").PadRight(14) + "  ";
                            }
                            if (aL[0].Gesamtnote != Gesamtnote)
                            {
                                this.Beschreibung = this.Beschreibung + ("Note:" + (aL[0].Gesamtnote == null ? "null" : aL[0].Gesamtnote) + "->" + Gesamtnote).PadRight(12) + "|";
                                this.Query += "s_note='" + Gesamtnote +"'" + ", ";
                            }
                            else
                            {
                                this.Query += (" ").PadRight(10) + "  ";
                            }
                            if (aL[0].Tendenz != Tendenz && aL[0].EinheitNP == "P")
                            {
                                this.Beschreibung = this.Beschreibung + ("Tendenz:" + (aL[0].Tendenz == null ? "null" : aL[0].Tendenz) + "->" + Tendenz).PadRight(15) + "|";
                                this.Query += "s_tendenz='" + Tendenz +"'" + ", ";
                            }
                            else
                            {
                                this.Query += (" ").PadRight(14) + "  ";
                            }
                        }
                    }

                    this.Zielfach = aL[0].Fach;
                    this.ZielLeistungId = aL[0].LeistungId;
                    this.Beschreibung += beschreibung;
                    this.Query += "ls_id_1=3844 "; // letzter Bearbeiter
                    this.Query += "WHERE noe_id=" + aL[0].LeistungId + ";";
                }
                else
                {
                    this.Beschreibung = "   |" + this.Beschreibung + "Note bleibt: " + aL[0].Gesamtnote + (aL[0].Tendenz == null ? " " : aL[0].Tendenz);
                    this.Query += "/* Keine Änderung.";
                }
            }            
        }
    }
}