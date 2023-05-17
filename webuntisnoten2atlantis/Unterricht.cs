// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;

namespace webuntisnoten2atlantis
{
    public class Unterricht
    {
        public int MarksPerLessonZeile { get; internal set; }
        public int LessonId { get; internal set; }
        public int LessonNumber { get; internal set; }
        public string Fach { get; internal set; }
        public string Lehrkraft { get; internal set; }
        public string Klassen { get; internal set; }
        public string Gruppe { get; internal set; }
        public int Periode { get; internal set; }
        public DateTime Startdate { get; internal set; }
        public DateTime Enddate { get; internal set; }
        public Leistung WebuntisLeistung { get; internal set; }
        public Leistung AtlantisLeistung { get; internal set; }
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
            AtlantisLeistung = new Leistung();
            AtlantisLeistung = atlantisLeistung;
        }

        public Unterricht(Leistung atlantisLeistung, int lessonNumber, string fach, Leistung webuntisLeistung, string lehrkraft, int marksPerLessonZeile, int periode, string gruppe, string klassen, DateTime startdate, DateTime enddate) : this(atlantisLeistung)
        {
            this.AtlantisLeistung = atlantisLeistung;
            this.LessonNumber = lessonNumber;
            this.Fach = fach;
            this.WebuntisLeistung = webuntisLeistung;
            this.Lehrkraft = lehrkraft;
            this.MarksPerLessonZeile = marksPerLessonZeile;
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

            string gesamtnote = WebuntisLeistung.Gesamtnote != "" ? WebuntisLeistung.Gesamtnote : null;
            string gesamtpunkte = WebuntisLeistung.Gesamtpunkte != "" ? WebuntisLeistung.Gesamtpunkte : null;
            string tendenz = WebuntisLeistung.Tendenz != "" ? WebuntisLeistung.Tendenz : null;
            WebuntisLeistung.Query = "";

            WebuntisLeistung.EinheitNP = AtlantisLeistung.EinheitNP;
            WebuntisLeistung.Beschreibung = AtlantisLeistung.SchlüsselExtern + "|" + (AtlantisLeistung.Nachname.PadRight(10)).Substring(0, 3) + " " + (AtlantisLeistung.Vorname.PadRight(10)).Substring(0, 2) + "|" 
                + AtlantisLeistung.Klasse.PadRight(5) + "|" 
                + (Lehrkraft == null ? "|" : Lehrkraft + "|").PadLeft(3) 
                + AtlantisLeistung.Fach.PadRight(5) 
                + (MarksPerLessonZeile == 0 ? "" : "|Zeile:" + MarksPerLessonZeile.ToString().PadLeft(4)) + "|" + AtlantisLeistung.Konferenzdatum.ToShortDateString() + "|" 
                + AtlantisLeistung.Bemerkung;

            // Falls Neu oder Update oder zuvor geholte Noten wieder nullen

            if (
                (AtlantisLeistung.Gesamtnote == null && AtlantisLeistung.Gesamtpunkte == null && AtlantisLeistung.Tendenz == null && (gesamtnote != null || gesamtpunkte != null || tendenz != null))                                  // Neu
                || (AtlantisLeistung.Gesamtnote != gesamtnote                                                                                 // UPD 
                || (AtlantisLeistung.Gesamtpunkte != gesamtpunkte && AtlantisLeistung.EinheitNP == "P") || (AtlantisLeistung.Tendenz != tendenz && AtlantisLeistung.EinheitNP == "P"))     // UPD Gym
               )
            {
                WebuntisLeistung.Query = "UPDATE noten_einzel SET ";

                // Falls Neu

                if (AtlantisLeistung.Gesamtnote == null && AtlantisLeistung.Gesamtpunkte == null && AtlantisLeistung.Tendenz == null)
                {
                    WebuntisLeistung.Beschreibung = "NEU|" + WebuntisLeistung.Beschreibung;

                    if (gesamtpunkte != null && AtlantisLeistung.EinheitNP == "P")
                    {
                        WebuntisLeistung.Beschreibung = WebuntisLeistung.Beschreibung + "P:[" + (AtlantisLeistung.Gesamtpunkte == null ? "  " : AtlantisLeistung.Gesamtpunkte.PadLeft(2)) + "]->[" + gesamtpunkte.PadLeft(2) + "]";
                        WebuntisLeistung.Query += "punkte=" + ("'" + gesamtpunkte).PadLeft(3) + "'" + ", ";
                    }
                    else
                    {
                        WebuntisLeistung.Query += (" ").PadRight(11) + "  ";
                    }
                    if (gesamtnote != null)
                    {
                        WebuntisLeistung.Beschreibung = WebuntisLeistung.Beschreibung + "N:[" + (AtlantisLeistung.Gesamtnote == null ? " " : AtlantisLeistung.Gesamtnote.PadRight(1)) + "]->[" + gesamtnote.PadRight(1) + "]";
                        WebuntisLeistung.Query += ("s_note='" + gesamtnote + "'" + ", ").PadRight(11);
                    }
                    else
                    {
                        WebuntisLeistung.Query += (" ").PadRight(10) + "  ";
                    }
                    if (WebuntisLeistung.Tendenz != null && AtlantisLeistung.EinheitNP == "P")
                    {
                        WebuntisLeistung.Beschreibung = WebuntisLeistung.Beschreibung + "T:[" + (AtlantisLeistung.Tendenz == null ? " " : AtlantisLeistung.Tendenz) + "]->[" + (WebuntisLeistung.Tendenz == null ? " " : WebuntisLeistung.Tendenz) + "]";
                        WebuntisLeistung.Query += ("s_tendenz='" + tendenz + "'").PadRight(12) + ",  ";
                    }
                    else
                    {
                        WebuntisLeistung.Query += (" ").PadRight(13) + "  ";
                    }
                }
                else // Falls Update
                {
                    if (AtlantisLeistung.Gesamtnote != gesamtnote || (AtlantisLeistung.Gesamtpunkte != gesamtpunkte && WebuntisLeistung.EinheitNP == "P") || (AtlantisLeistung.Tendenz != WebuntisLeistung.Tendenz && WebuntisLeistung.EinheitNP == "P"))
                    {
                        WebuntisLeistung.Beschreibung = "UPD|" + WebuntisLeistung.Beschreibung;

                        if (AtlantisLeistung.Gesamtpunkte != gesamtpunkte && AtlantisLeistung.EinheitNP == "P")
                        {
                            WebuntisLeistung.Beschreibung = WebuntisLeistung.Beschreibung + "P:[" + (AtlantisLeistung.Gesamtpunkte == null ? "  " : AtlantisLeistung.Gesamtpunkte.PadLeft(2)) + "]->[" + WebuntisLeistung.Gesamtpunkte.PadLeft(2) + "]";
                            WebuntisLeistung.Query += "punkte=" + ("'" + gesamtpunkte).PadLeft(3) + "'" + ", ";
                        }
                        else
                        {
                            WebuntisLeistung.Query += (" ").PadRight(11) + "  ";
                        }
                        if (AtlantisLeistung.Gesamtnote != gesamtnote)
                        {
                            WebuntisLeistung.Beschreibung = WebuntisLeistung.Beschreibung + "N:[" + (AtlantisLeistung.Gesamtnote == null ? " " : AtlantisLeistung.Gesamtnote.PadLeft(1)) + "]->[" + WebuntisLeistung.Gesamtnote.PadLeft(1) + "]";
                            WebuntisLeistung.Query += ("s_note='" + gesamtnote + "'" + ", ").PadRight(10);
                        }
                        else
                        {
                            WebuntisLeistung.Query += (" ").PadRight(11) + "  ";
                        }
                        if (AtlantisLeistung.Tendenz != tendenz && AtlantisLeistung.EinheitNP == "P")
                        {
                            WebuntisLeistung.Beschreibung = WebuntisLeistung.Beschreibung + "T:[" + (AtlantisLeistung.Tendenz == null ? " " : AtlantisLeistung.Tendenz) + "]->[" + (tendenz == null ? " " : tendenz) + "]";
                            WebuntisLeistung.Query += ("s_tendenz='" + tendenz + "',").PadRight(16);
                        }
                        else
                        {
                            WebuntisLeistung.Query += (" ").PadRight(15);
                        }
                    }
                }

                WebuntisLeistung.Zielfach = AtlantisLeistung.Fach;
                WebuntisLeistung.ZielLeistungId = AtlantisLeistung.LeistungId;
                WebuntisLeistung.Beschreibung = WebuntisLeistung.Beschreibung + (AtlantisLeistung.Fach != "REL" && gesamtnote == "-" ? "Zeugnisbemerkung?|" : "");

                if (WebuntisLeistung.ReligionAbgewählt && WebuntisLeistung.FachAliases.Contains("REL"))
                {
                    WebuntisLeistung.Beschreibung += WebuntisLeistung.Beschreibung + "Reli abgewählt.";
                }
                WebuntisLeistung.Query += "ls_id_1=1337 "; // letzter Bearbeiter
                WebuntisLeistung.Query += "WHERE noe_id=" + AtlantisLeistung.LeistungId + ";";
            }
            else
            {
                WebuntisLeistung.Beschreibung = "   |" + WebuntisLeistung.Beschreibung + "Note bleibt: " + (AtlantisLeistung.Gesamtnote + (AtlantisLeistung.Tendenz == null ? " " : AtlantisLeistung.Tendenz)).PadLeft(2) + (AtlantisLeistung.EinheitNP == "P" ? "(" + AtlantisLeistung.Gesamtpunkte.PadLeft(2) + " P)" : "");
                WebuntisLeistung.Query += "/* KEINE ÄNDERUNG   SET punkte='" + (gesamtpunkte == null ? "  " : gesamtpunkte.PadLeft(2)) + "',".PadRight(2) + " s_note='" + (AtlantisLeistung.Gesamtnote == null ? "" : AtlantisLeistung.Gesamtnote.PadRight(1)) + "', s_tendenz='" + (AtlantisLeistung.Tendenz == null ? " " : AtlantisLeistung.Tendenz) + "',  ls_id_1=1337 WHERE noe_id=" + AtlantisLeistung.LeistungId + "*/";
            }
        }
    }
}