// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Linq;

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

        public bool HatBemerkung { get; internal set; }
        /// <summary>
        /// Wenn Schüler der Anlage A die Zeugnisart A01AS gesetzt haben, dann werden für sie alte Noten geholt.
        /// </summary>
        public string Zeugnisart { get; internal set; }
        public string Prüfungsart { get; internal set; }
        public string Nachname { get; set; }
        public string Vorname { get; set; }
        public DateTime Geburtsdatum { get; internal set; }
        public bool Volljährig { get; internal set; }
        public string Bereich { get; internal set; }
        public string Gesamtpunkte_12_1 { get; internal set; }
        public string Gesamtpunkte_12_2 { get; internal set; }
        public string Gesamtpunkte_13_1 { get; internal set; }
        public string Gesamtpunkte_13_2 { get; internal set; }
        public int LehrkraftAtlantisId { get; internal set; }
        public string Note { get; internal set; }
        public string Punkte { get; internal set; }
        public DateTime DatumReligionAbmeldung { get; internal set; }
        public List<string> FachAliases { get; internal set; }
        public bool IstGeholteNote { get; set; }
        public int MarksPerLessonZeile { get; internal set; }
        public int Reihenfolge { get; internal set; }
        public bool Zugeordnet { get; internal set; }
        
        // ID des Notenkopfes,
        public int NokId { get; internal set; }

        public Leistung(string name, string fach, List<string> fachAliases, string gesamtnote, string gesamtpunkte, string tendenz, DateTime datum, string nachname, string lehrkraft, int schlüsselExtern, int marksPerLessonZeile)
        {
            Name = name;
            Fach = fach;
            FachAliases = fachAliases;
            Gesamtnote = gesamtnote;
            Gesamtpunkte = gesamtpunkte;
            Tendenz = tendenz;
            Datum = datum;
            Nachname = nachname;
            Lehrkraft = lehrkraft;
            SchlüsselExtern = schlüsselExtern;
            MarksPerLessonZeile = marksPerLessonZeile;
        }

        public Leistung()
        {
        }

        public Leistung(Leistung atlantisLeistung, string bemerkung)
        {
            Klasse=atlantisLeistung.Klasse;
            Fach = atlantisLeistung.Fach;
            Gesamtnote = atlantisLeistung.Gesamtnote;
            Gesamtpunkte = atlantisLeistung.Gesamtpunkte;
            Tendenz = atlantisLeistung.Tendenz;
            Konferenzdatum = atlantisLeistung.Konferenzdatum;
            Beschreibung = atlantisLeistung.Beschreibung;
            Bemerkung = (atlantisLeistung.Bemerkung == null ? "" : atlantisLeistung.Bemerkung) + bemerkung;
            EinheitNP = atlantisLeistung.EinheitNP;
            Nachname = atlantisLeistung.Nachname;
            Vorname = atlantisLeistung.Vorname;
            LeistungId = atlantisLeistung.LeistungId;
            SchlüsselExtern = atlantisLeistung.SchlüsselExtern;
            Query = "";
        }

        public Leistung(string fach, string gesamtnote, DateTime konferenzdatum)
        {
            Fach = fach;
            Gesamtnote = gesamtnote;
            Konferenzdatum = konferenzdatum;
        }

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

        internal void GetFachAliases()
        {
            FachAliases = new List<string>();

            var nameAliases = new List<string>();

            if (Fach.StartsWith("D  G1"))
            {
                nameAliases.Add("D G1");
                nameAliases.Add("D G2");
                nameAliases.Add("D G");
                nameAliases.Add("D  G1");
                nameAliases.Add("D  G2");
                nameAliases.Add("D  G");
            }

            if (Fach.StartsWith("E  G1"))
            {
                nameAliases.Add("E G1");
                nameAliases.Add("E G2");
                nameAliases.Add("E G");
                nameAliases.Add("E  G1");
                nameAliases.Add("E  G2");
                nameAliases.Add("E  G");
            }

            if (Fach.StartsWith("M  G1"))
            {
                nameAliases.Add("M G1");
                nameAliases.Add("M G2");
                nameAliases.Add("M G");
                nameAliases.Add("M  G1");
                nameAliases.Add("M  G2");
                nameAliases.Add("M  G");
            }

            if (Fach.StartsWith("D  L1"))
            {
                nameAliases.Add("D L1");
                nameAliases.Add("D L2");
                nameAliases.Add("D L");
                nameAliases.Add("D  L1");
                nameAliases.Add("D  L2");
                nameAliases.Add("D  L");
            }

            if (Fach.StartsWith("E  L1"))
            {
                nameAliases.Add("E L1");
                nameAliases.Add("E L2");
                nameAliases.Add("E L");
                nameAliases.Add("E  L1");
                nameAliases.Add("E  L2");
                nameAliases.Add("E  L");
            }

            if (Fach.StartsWith("M  L1"))
            {
                nameAliases.Add("M L1");
                nameAliases.Add("M L2");
                nameAliases.Add("M L");
                nameAliases.Add("M  L1");
                nameAliases.Add("M  L2");
                nameAliases.Add("M  L");
            }

            if (Fach.StartsWith("D G1"))
            {
                nameAliases.Add("D G1");
                nameAliases.Add("D G2");
                nameAliases.Add("D G");
                nameAliases.Add("D  G1");
                nameAliases.Add("D  G2");
                nameAliases.Add("D  G");
            }

            if (Fach.StartsWith("E G1"))
            {
                nameAliases.Add("E G1");
                nameAliases.Add("E G2");
                nameAliases.Add("E G");
                nameAliases.Add("E  G1");
                nameAliases.Add("E  G2");
                nameAliases.Add("E  G");
            }

            if (Fach.StartsWith("M G1"))
            {
                nameAliases.Add("M G1");
                nameAliases.Add("M G2");
                nameAliases.Add("M G");
                nameAliases.Add("M  G1");
                nameAliases.Add("M  G2");
                nameAliases.Add("M  G");
            }

            if (Fach.StartsWith("D L1"))
            {
                nameAliases.Add("D L1");
                nameAliases.Add("D L2");
                nameAliases.Add("D L");
                nameAliases.Add("D  L1");
                nameAliases.Add("D  L2");
                nameAliases.Add("D  L");
            }

            if (Fach.StartsWith("E L1"))
            {
                nameAliases.Add("E L1");
                nameAliases.Add("E L2");
                nameAliases.Add("E L");
                nameAliases.Add("E  L1");
                nameAliases.Add("E  L2");
                nameAliases.Add("E  L");
            }

            if (Fach.StartsWith("M L1"))
            {
                nameAliases.Add("M L1");
                nameAliases.Add("M L2");
                nameAliases.Add("M L");
                nameAliases.Add("M  L1");
                nameAliases.Add("M  L2");
                nameAliases.Add("M  L");
            }

            if (Fach.StartsWith("S G"))
            {
                nameAliases.Add("S G1");
                nameAliases.Add("S G2");
                nameAliases.Add("S G");
                nameAliases.Add("S GB");
                nameAliases.Add("S GD");
                nameAliases.Add("S  G1");
                nameAliases.Add("S  G2");
                nameAliases.Add("S  G");
                nameAliases.Add("S  GB");
                nameAliases.Add("S  GD");
            }

            if (Fach.StartsWith("S  G"))
            {
                nameAliases.Add("S G1");
                nameAliases.Add("S G2");
                nameAliases.Add("S G");
                nameAliases.Add("S GB");
                nameAliases.Add("S GD");
                nameAliases.Add("S  G1");
                nameAliases.Add("S  G2");
                nameAliases.Add("S  G");
                nameAliases.Add("S  GB");
                nameAliases.Add("S  GD");
            }

            if (Fach.StartsWith("N G"))
            {
                nameAliases.Add("N G1");
                nameAliases.Add("N G2");
                nameAliases.Add("N G");
                nameAliases.Add("N GB");
                nameAliases.Add("N GD");
                nameAliases.Add("N  G1");
                nameAliases.Add("N  G2");
                nameAliases.Add("N  G");
                nameAliases.Add("N  GB");
                nameAliases.Add("N  GD");
            }

            if (Fach.StartsWith("N  G"))
            {
                nameAliases.Add("N G1");
                nameAliases.Add("N G2");
                nameAliases.Add("N G");
                nameAliases.Add("N GB");
                nameAliases.Add("N GD");
                nameAliases.Add("N  G1");
                nameAliases.Add("N  G2");
                nameAliases.Add("N  G");
                nameAliases.Add("N  GB");
                nameAliases.Add("N  GD");
            }

            if (Fach.StartsWith("CAD"))
            {
                nameAliases.Add("CAD1");
                nameAliases.Add("CAD2");
                nameAliases.Add("CAD3");
                nameAliases.Add("CAD/CAM");
                nameAliases.Add("CAD/ CAM");
                nameAliases.Add("CAD");
            }

            if (Fach == "PKG" || Fach == "PK")
            {
                nameAliases.Add("PK");
                nameAliases.Add("PKG");
            }
            if (Fach == "N" || Fach == "NB1" || Fach == "NB2" || Fach == "NA1" || Fach == "NA2")
            {
                nameAliases.Add("N");
                nameAliases.Add("NB1");
                nameAliases.Add("NB2");
                nameAliases.Add("NA1");
                nameAliases.Add("NA2");
            }
            if (Fach == "S" || Fach == "SB1" || Fach == "SB2" || Fach == "SA1" || Fach == "SA2")
            {
                nameAliases.Add("S");
                nameAliases.Add("SB1");
                nameAliases.Add("SB2");
                nameAliases.Add("SA1");
                nameAliases.Add("SA2");
            }

            if (Fach == "E" || Fach == "EB1" || Fach == "EB2" || Fach == "EA1" || Fach == "EA2")
            {
                nameAliases.Add("E");
                nameAliases.Add("EB1");
                nameAliases.Add("EB2");
                nameAliases.Add("EA1");
                nameAliases.Add("EA2");
            }

            if (Fach == "KR" || Fach == "ER" || Fach == "REL" || Fach == "KR " || Fach == "ER ")
            {
                nameAliases.Add("KR");
                nameAliases.Add("ER");
                nameAliases.Add("REL");
                nameAliases.Add("KR ");
                nameAliases.Add("ER ");
            }

            if (!nameAliases.Contains(Fach))
            {
                nameAliases.Add(Fach);
            }

            // Wenn das Fach einen Zähler hat, wird der Zähler entfernt.

            if (Fach != null && Fach != "" && char.IsDigit(Fach[Fach.Length - 1]))
            {
                nameAliases.Add(Fach.Substring(0,Fach.Length - 1));
            }

            FachAliases.AddRange(nameAliases);
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

        internal void Update()
        {
            try
            {
                string o = this.Query.PadRight(103, ' ') + (this.Query.Contains("Keine Änderung") ? "   " : "/* ") + this.Beschreibung;
                Global.SqlZeilen.Add((o.Substring(0, Math.Min(178, o.Length))).PadRight(180) + " */");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal bool leistungDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum(List<string> aktSj)
        {
            if (Konferenzdatum > DateTime.Now.Date.AddDays(-14) && Konferenzdatum.Date < DateTime.Now.Date)
            {
                if (Schuljahr == aktSj[0] + "/" + aktSj[1])
                {
                    return true;
                }
            }
            return false;
        }

        internal bool HatReligionAbgewählt(Leistungen atlantisLeistungen)
        {
            // Wenn bereits eine andere Leistung dieses Schülers mit einer Reliabmeldung festgestellt wurde, hat er abgewählt

            if ((from x in atlantisLeistungen where x.IstGeholteNote where x.ReligionAbgewählt select x).Any())
            {
                return true;
            }

            // Wenn eine Abmeldung mit neuerem Datum vorliegt, ist der Schüler abgemeldet.

            return this.DatumReligionAbmeldung.Year > 1 ? true : false;
        }
    }
}