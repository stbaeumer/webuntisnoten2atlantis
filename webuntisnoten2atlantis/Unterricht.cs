// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;

namespace webuntisnoten2atlantis
{
    public class Unterricht
    {
        public int Zeile { get; internal set; }
        public int LessonId { get; internal set; }
        public int LessonNumber { get; internal set; }
        public string Fach { get; internal set; }
        public string Lehrkraft { get; internal set; }
        public string Klassen { get; internal set; }
        public string Gruppe { get; internal set; }
        public int Periode { get; internal set; }
        public DateTime Startdate { get; internal set; }
        public DateTime Enddate { get; internal set; }
        public Leistung WL { get; internal set; }
        public Leistung AL { get; internal set; }
        public int Reihenfolge { get; internal set; }
        public string Bemerkung { get; internal set; }

        public Unterricht()
        {
        }

        public Unterricht(Leistung atlantisLeistung)
        {            
            if (atlantisLeistung != null)
            {
                Fach = atlantisLeistung.Fach;
                Lehrkraft = atlantisLeistung.Lehrkraft;
                Klassen = atlantisLeistung.Klasse;
            }
            AL = new Leistung();
            AL = atlantisLeistung;
        }

        public Unterricht(Leistung atlantisLeistung, int lessonNumber, string fach, Leistung webuntisLeistung, string lehrkraft, int marksPerLessonZeile, int periode, string gruppe, string klassen, DateTime startdate, DateTime enddate) : this(atlantisLeistung)
        {
            this.AL = atlantisLeistung;
            this.LessonNumber = lessonNumber;
            this.Fach = fach;
            this.WL = webuntisLeistung;
            this.Lehrkraft = lehrkraft;
            this.Zeile = marksPerLessonZeile;
            this.Periode = periode;
            this.Gruppe = gruppe;
            this.Klassen = klassen;
            this.Startdate = startdate;
            this.Enddate = enddate;
            this.Bemerkung = "";
        }

        internal void QueryBauen()
        {
            // Für eine einfachere Vergleichbarkeit wird eine leere Gesamtnote genullt

            string gesamtnote = WL.Gesamtnote != "" ? WL.Gesamtnote : null;
            string gesamtpunkte = WL.Gesamtpunkte != "" ? WL.Gesamtpunkte : null;
            string tendenz = WL.Tendenz != "" ? WL.Tendenz : null;
            
            WL.Query = "";

            WL.EinheitNP = AL.EinheitNP;
            WL.Beschreibung = AL.SchlüsselExtern + "|" + (AL.Nachname.PadRight(10)).Substring(0, 3) + " " + (AL.Vorname.PadRight(10)).Substring(0, 2) + "|"
                + AL.Klasse + "|"
                + (Lehrkraft == null ? "|" : Lehrkraft + "|")
                + AL.Fach + "|"
                + (WL.MarksPerLessonZeile == 0 ? "" : "Zeile:" + WL.MarksPerLessonZeile + "|") +
                (AL.Konferenzdatum != null && AL.Konferenzdatum.Year != 1 ? AL.Konferenzdatum.ToShortDateString() + "|" : "")
                + AL.Bemerkung;

            // Falls Neu oder Update oder zuvor geholte Noten wieder nullen

            if (
                (AL.Gesamtnote == null && AL.Gesamtpunkte == null && AL.Tendenz == null && (gesamtnote != null || gesamtpunkte != null || tendenz != null))                                  // Neu
                || (AL.Gesamtnote != gesamtnote                                                                                 // UPD 
                || (AL.Gesamtpunkte != gesamtpunkte && AL.EinheitNP == "P") || (AL.Tendenz != tendenz && AL.EinheitNP == "P"))     // UPD Gym
               )
            {
                WL.Query = "UPDATE noten_einzel SET ";

                // Falls Neu

                if (AL.Gesamtnote == null && AL.Gesamtpunkte == null && AL.Tendenz == null)
                {
                    WL.Beschreibung = "NEU|" + WL.Beschreibung;

                    if (gesamtpunkte != null && AL.EinheitNP == "P")
                    {
                        WL.Beschreibung = WL.Beschreibung + "P:[" + (AL.Gesamtpunkte == null ? "  " : AL.Gesamtpunkte.PadLeft(2)) + "]->[" + gesamtpunkte.PadLeft(2) + "]";
                        WL.Query += "punkte=" + ("'" + gesamtpunkte).PadLeft(3) + "'" + ", ";
                    }
                    else
                    {
                        WL.Query += (" ").PadRight(11) + "  ";
                    }
                    if (gesamtnote != null)
                    {
                        WL.Beschreibung = WL.Beschreibung + "N:[" + (AL.Gesamtnote == null ? " " : AL.Gesamtnote.PadRight(1)) + "]->[" + gesamtnote.PadRight(1) + "]";
                        WL.Query += ("s_note='" + gesamtnote + "'" + ", ").PadRight(11);
                    }
                    else
                    {
                        WL.Query += (" ").PadRight(10) + "  ";
                    }
                    if (WL.Tendenz != null && AL.EinheitNP == "P")
                    {
                        WL.Beschreibung = WL.Beschreibung + "T:[" + (AL.Tendenz == null ? " " : AL.Tendenz) + "]->[" + (WL.Tendenz == null ? " " : WL.Tendenz) + "]";
                        WL.Query += ("s_tendenz='" + tendenz + "'").PadRight(12) + ",  ";
                    }
                    else
                    {
                        WL.Query += (" ").PadRight(13) + "  ";
                    }
                }
                else // Falls Update
                {
                    if (AL.Gesamtnote != gesamtnote || (AL.Gesamtpunkte != gesamtpunkte && WL.EinheitNP == "P") || (AL.Tendenz != WL.Tendenz && WL.EinheitNP == "P"))
                    {
                        WL.Beschreibung = "UPD|" + WL.Beschreibung;

                        if (AL.Gesamtpunkte != gesamtpunkte && AL.EinheitNP == "P")
                        {
                            WL.Beschreibung = WL.Beschreibung + "P:[" + (AL.Gesamtpunkte == null ? "  " : AL.Gesamtpunkte.PadLeft(2)) + "]->[" + WL.Gesamtpunkte.PadLeft(2) + "]";
                            WL.Query += "punkte=" + ("'" + gesamtpunkte).PadLeft(3) + "'" + ", ";
                        }
                        else
                        {
                            WL.Query += (" ").PadRight(11) + "  ";
                        }
                        if (AL.Gesamtnote != gesamtnote)
                        {
                            WL.Beschreibung = WL.Beschreibung + "N:[" + (AL.Gesamtnote == null ? " " : AL.Gesamtnote.PadLeft(1)) + "]->[" + WL.Gesamtnote.PadLeft(1) + "]";
                            WL.Query += ("s_note='" + gesamtnote + "'" + ", ").PadRight(10);
                        }
                        else
                        {
                            WL.Query += (" ").PadRight(11) + "  ";
                        }
                        if (AL.Tendenz != tendenz && AL.EinheitNP == "P")
                        {
                            WL.Beschreibung = WL.Beschreibung + "T:[" + (AL.Tendenz == null ? " " : AL.Tendenz) + "]->[" + (tendenz == null ? " " : tendenz) + "]";
                            WL.Query += ("s_tendenz='" + tendenz + "',").PadRight(16);
                        }
                        else
                        {
                            WL.Query += (" ").PadRight(15);
                        }
                    }
                }

                WL.Zielfach = AL.Fach;
                WL.ZielLeistungId = AL.LeistungId;
                WL.Beschreibung = WL.Beschreibung + (AL.Fach != "REL" && gesamtnote == "-" ? "Zeugnisbemerkung?|" : "");

                if (WL.ReligionAbgewählt && WL.FachAliases.Contains("REL"))
                {
                    WL.Beschreibung += WL.Beschreibung + "Reli abgewählt.";
                }
                WL.Query += "ls_id_1=1337 "; // letzter Bearbeiter
                WL.Query += "WHERE noe_id=" + AL.LeistungId + ";";
            }
            else
            {
                WL.Beschreibung = "   |" + WL.Beschreibung + "Note bleibt: " + (AL.Gesamtnote + (AL.Tendenz == null ? " " : AL.Tendenz)).PadLeft(2) + (AL.EinheitNP == "P" ? "(" + AL.Gesamtpunkte.PadLeft(2) + " P)" : "");
                WL.Query += "/* KEINE ÄNDERUNG   SET punkte='" + (gesamtpunkte == null ? "  " : gesamtpunkte.PadLeft(2)) + "',".PadRight(2) + " s_note='" + (AL.Gesamtnote == null ? "" : AL.Gesamtnote.PadRight(1)) + "', s_tendenz='" + (AL.Tendenz == null ? " " : AL.Tendenz) + "',  ls_id_1=1337 WHERE noe_id=" + AL.LeistungId + "*/";
            }
        }
    }
}