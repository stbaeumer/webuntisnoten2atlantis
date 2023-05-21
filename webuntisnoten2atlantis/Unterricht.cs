// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;

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
        public Leistung LeistungW { get; internal set; }
        public Leistung LeistungA { get; internal set; }
        public int Reihenfolge { get; internal set; }
        public string Bemerkung { get; internal set; }
        public string KursOderAlle { get; internal set; }
        public string FachnameAtlantis { get; internal set; }

        /// <summary>
        /// Falls keine automatische Zuordnung einer Atlantisleitung zu einem Unterricht gelingt, werden alle infragkommenden 
        /// Leistungen hinzugefügt, um später manuell daraus auswählen zu können.
        /// </summary>
        public Leistungen InfragekommendeLeistungenA { get; private set; }

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
            LeistungA = new Leistung();
            LeistungA = atlantisLeistung;
        }

        public Unterricht(Leistung atlantisLeistung, int lessonNumber, string fach, Leistung webuntisLeistung, string lehrkraft, int marksPerLessonZeile, int periode, string gruppe, string klassen, DateTime startdate, DateTime enddate) : this(atlantisLeistung)
        {
            InfragekommendeLeistungenA = new Leistungen();
            this.LeistungA = atlantisLeistung;
            this.LessonNumber = lessonNumber;
            this.Fach = fach;
            this.LeistungW = webuntisLeistung;
            this.Lehrkraft = lehrkraft;
            this.Zeile = marksPerLessonZeile;
            this.Periode = periode;
            this.Gruppe = gruppe;
            this.KursOderAlle = (gruppe == "" ? "Alle" : "Kurs");
            this.Klassen = klassen;
            this.Startdate = startdate;
            this.Enddate = enddate;
            this.Bemerkung = "";
        }

        public Unterricht(string lehrkraft, string fach, string klasse, int lessonId, int reihenfolge, string gruppe, string kursOderAlle, Leistung leistungW, Leistung leistungA, int lessonNumber, string fachNameAtlantis, Leistungen infragekommendeLeistungenA)
        {
            Lehrkraft = lehrkraft;
            Fach = fach;
            Klassen = klasse;
            LessonId = lessonId;
            Reihenfolge = reihenfolge;
            Gruppe = gruppe;
            KursOderAlle = kursOderAlle;
            LeistungA = leistungA;
            LeistungW = leistungW;
            LessonNumber = lessonNumber;
            FachnameAtlantis = fachNameAtlantis;
            InfragekommendeLeistungenA = infragekommendeLeistungenA;
        }

        internal void QueryBauen()
        {
            // Für eine einfachere Vergleichbarkeit wird eine leere Gesamtnote genullt

            string gesamtnote = LeistungW.Gesamtnote != "" ? LeistungW.Gesamtnote : null;
            string gesamtpunkte = LeistungW.Gesamtpunkte != "" ? LeistungW.Gesamtpunkte : null;
            string tendenz = LeistungW.Tendenz != "" ? LeistungW.Tendenz : null;
            
            LeistungW.Query = "";

            LeistungW.EinheitNP = LeistungA.EinheitNP;
            LeistungW.Beschreibung = LeistungA.SchlüsselExtern + "|" + (LeistungA.Nachname.PadRight(10)).Substring(0, 3) + " " + (LeistungA.Vorname.PadRight(10)).Substring(0, 2) + "|"
                + LeistungA.Klasse + "|"
                + (Lehrkraft == null ? "|" : Lehrkraft + "|")
                + LeistungA.Fach + "|"
                + (LeistungW.MarksPerLessonZeile == 0 ? "" : "Zeile:" + LeistungW.MarksPerLessonZeile + "|") +
                (LeistungA.Konferenzdatum != null && LeistungA.Konferenzdatum.Year != 1 ? LeistungA.Konferenzdatum.ToShortDateString() + "|" : "")
                + LeistungA.Bemerkung;

            // Falls Neu oder Update oder zuvor geholte Noten wieder nullen

            if (
                (LeistungA.Gesamtnote == null && LeistungA.Gesamtpunkte == null && LeistungA.Tendenz == null && (gesamtnote != null || gesamtpunkte != null || tendenz != null))                                  // Neu
                || (LeistungA.Gesamtnote != gesamtnote                                                                                 // UPD 
                || (LeistungA.Gesamtpunkte != gesamtpunkte && LeistungA.EinheitNP == "P") || (LeistungA.Tendenz != tendenz && LeistungA.EinheitNP == "P"))     // UPD Gym
               )
            {
                LeistungW.Query = "UPDATE noten_einzel SET ";

                // Falls Neu

                if (LeistungA.Gesamtnote == null && LeistungA.Gesamtpunkte == null && LeistungA.Tendenz == null)
                {
                    LeistungW.Beschreibung = "NEU|" + LeistungW.Beschreibung;

                    if (gesamtpunkte != null && LeistungA.EinheitNP == "P")
                    {
                        LeistungW.Beschreibung = LeistungW.Beschreibung + "P:[" + (LeistungA.Gesamtpunkte == null ? "  " : LeistungA.Gesamtpunkte.PadLeft(2)) + "]->[" + gesamtpunkte.PadLeft(2) + "]";
                        LeistungW.Query += "punkte=" + ("'" + gesamtpunkte).PadLeft(3) + "'" + ", ";
                    }
                    else
                    {
                        LeistungW.Query += (" ").PadRight(11) + "  ";
                    }
                    if (gesamtnote != null)
                    {
                        LeistungW.Beschreibung = LeistungW.Beschreibung + "N:[" + (LeistungA.Gesamtnote == null ? " " : LeistungA.Gesamtnote.PadRight(1)) + "]->[" + gesamtnote.PadRight(1) + "]";
                        LeistungW.Query += ("s_note='" + gesamtnote + "'" + ", ").PadRight(11);
                    }
                    else
                    {
                        LeistungW.Query += (" ").PadRight(10) + "  ";
                    }
                    if (LeistungW.Tendenz != null && LeistungA.EinheitNP == "P")
                    {
                        LeistungW.Beschreibung = LeistungW.Beschreibung + "T:[" + (LeistungA.Tendenz == null ? " " : LeistungA.Tendenz) + "]->[" + (LeistungW.Tendenz == null ? " " : LeistungW.Tendenz) + "]";
                        LeistungW.Query += ("s_tendenz='" + tendenz + "'").PadRight(12) + ",  ";
                    }
                    else
                    {
                        LeistungW.Query += (" ").PadRight(13) + "  ";
                    }
                }
                else // Falls Update
                {
                    if (LeistungA.Gesamtnote != gesamtnote || (LeistungA.Gesamtpunkte != gesamtpunkte && LeistungW.EinheitNP == "P") || (LeistungA.Tendenz != LeistungW.Tendenz && LeistungW.EinheitNP == "P"))
                    {
                        LeistungW.Beschreibung = "UPD|" + LeistungW.Beschreibung;

                        if (LeistungA.Gesamtpunkte != gesamtpunkte && LeistungA.EinheitNP == "P")
                        {
                            LeistungW.Beschreibung = LeistungW.Beschreibung + "P:[" + (LeistungA.Gesamtpunkte == null ? "  " : LeistungA.Gesamtpunkte.PadLeft(2)) + "]->[" + LeistungW.Gesamtpunkte.PadLeft(2) + "]";
                            LeistungW.Query += "punkte=" + ("'" + gesamtpunkte).PadLeft(3) + "'" + ", ";
                        }
                        else
                        {
                            LeistungW.Query += (" ").PadRight(11) + "  ";
                        }
                        if (LeistungA.Gesamtnote != gesamtnote)
                        {
                            LeistungW.Beschreibung = LeistungW.Beschreibung + "N:[" + (LeistungA.Gesamtnote == null ? " " : LeistungA.Gesamtnote.PadLeft(1)) + "]->[" + LeistungW.Gesamtnote.PadLeft(1) + "]";
                            LeistungW.Query += ("s_note='" + gesamtnote + "'" + ", ").PadRight(10);
                        }
                        else
                        {
                            LeistungW.Query += (" ").PadRight(11) + "  ";
                        }
                        if (LeistungA.Tendenz != tendenz && LeistungA.EinheitNP == "P")
                        {
                            LeistungW.Beschreibung = LeistungW.Beschreibung + "T:[" + (LeistungA.Tendenz == null ? " " : LeistungA.Tendenz) + "]->[" + (tendenz == null ? " " : tendenz) + "]";
                            LeistungW.Query += ("s_tendenz='" + tendenz + "',").PadRight(16);
                        }
                        else
                        {
                            LeistungW.Query += (" ").PadRight(15);
                        }
                    }
                }

                LeistungW.Beschreibung = LeistungW.Beschreibung + (LeistungA.Fach != "REL" && gesamtnote == "-" ? "Zeugnisbemerkung?|" : "");

                if (LeistungW.ReligionAbgewählt && LeistungW.FachAliases.Contains("REL"))
                {
                    LeistungW.Beschreibung += LeistungW.Beschreibung + "Reli abgewählt.";
                }
                LeistungW.Query += "ls_id_1=1337 "; // letzter Bearbeiter
                LeistungW.Query += "WHERE noe_id=" + LeistungA.LeistungId + ";";
            }
            else
            {
                LeistungW.Beschreibung = "   |" + LeistungW.Beschreibung + "Note bleibt: " + (LeistungA.Gesamtnote + (LeistungA.Tendenz == null ? " " : LeistungA.Tendenz)).PadLeft(2) + (LeistungA.EinheitNP == "P" ? "(" + LeistungA.Gesamtpunkte.PadLeft(2) + " P)" : "");
                LeistungW.Query += "/* KEINE ÄNDERUNG   SET punkte='" + (gesamtpunkte == null ? "  " : gesamtpunkte.PadLeft(2)) + "',".PadRight(2) + " s_note='" + (LeistungA.Gesamtnote == null ? "" : LeistungA.Gesamtnote.PadRight(1)) + "', s_tendenz='" + (LeistungA.Tendenz == null ? " " : LeistungA.Tendenz) + "',  ls_id_1=1337 WHERE noe_id=" + LeistungA.LeistungId + "*/";
            }
        }
    }
}