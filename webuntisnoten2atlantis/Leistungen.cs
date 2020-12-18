// Published under the terms of GPLv3 Stefan Bäumer 2019.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;

namespace webuntisnoten2atlantis
{
    public class Leistungen : List<Leistung>
    {
        public Leistungen(string datei)
        {
            var leistungen = new Leistungen();

            using (StreamReader reader = new StreamReader(datei))
            {
                string überschrift = reader.ReadLine();

                Console.Write("Leistungsdaten aus Webuntis ".PadRight(71, '.'));

                int i = 1;

                while (true)
                {
                    string line = reader.ReadLine();                    
                    try
                    {
                        if (line != null)
                        {
                            Leistung leistung = new Leistung();
                            var x = line.Split('\t');
                            i++;

                            if (x.Length != 10)
                            {
                                Console.WriteLine("MarksPerLesson.csv: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.csv korrigiert werden.");
                                Console.ReadKey();
                                DateiÖffnen(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MarksPerLesson.csv");
                                throw new Exception("MarksPerLesson.csv: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.csv korrigiert werden.");
                            }

                            leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            leistung.Name = x[1];
                            leistung.Klasse = x[2];
                            leistung.Fach = x[3];                            
                            leistung.Gesamtpunkte = x[9].Split('.')[0];
                            leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);
                            leistung.Bemerkung = x[6];
                            leistung.Benutzer = x[7];
                            leistung.SchlüsselExtern = Convert.ToInt32(x[8]);
                            leistungen.Add(leistung);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Console.WriteLine((" " + leistungen.Count.ToString()).PadLeft(30, '.'));
            }
            this.AddRange((from l in leistungen select l).OrderBy(x => x.Anlage).ThenBy(x => x.Klasse));
        }

        public Leistungen(string connetionstringAtlantis, string aktSj)
        {
            Global.Output.Add("/* ****************************************************************************** */");
            Global.Output.Add("/* Diese Datei enthält alle Noten und Fehlzeiten aus Webuntis.Sie können alle No- */");
            Global.Output.Add("/* ten und Fehlzeiten aus Webuntis nach Atlantis importieren, indem Sie diese Da- */");
            Global.Output.Add("/* tei in Atlantis unter Funktionen>SQL-Anweisung hochladen.                      */");
            Global.Output.Add("/* Published under the terms of GPLv3. Hoping for the best!                       */");
            Global.Output.Add("/* " + (System.Security.Principal.WindowsIdentity.GetCurrent().Name + " " + DateTime.Now.ToString()).PadRight(78) + " */");
            Global.Output.Add("/* ****************************************************************************** */");
            Global.Output.Add("");

            try
            {
                var typ = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 8) ? "JZ" : "HZ";

                Console.Write(("Leistungsdaten aus Atlantis (" + typ + ")").PadRight(71, '.'));

                using (OdbcConnection connection = new OdbcConnection(connetionstringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.noten_einzel.noe_id AS LeistungId,
DBA.noten_einzel.fa_id,
DBA.noten_einzel.kurztext AS Fach,
DBA.noten_einzel.zeugnistext AS Zeugnistext,
DBA.noten_einzel.s_note AS Note,
DBA.noten_einzel.punkte AS Punkte,
DBA.noten_einzel.s_einheit AS Einheit,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.pu_id AS SchlüsselExtern,
DBA.schue_sj.s_religions_unterricht AS Religion,
DBA.schue_sj.dat_austritt AS ausgetreten,
DBA.schue_sj.vorgang_akt_satz_jn AS SchuelerAktivInDieserKlasse,
(substr(schue_sj.s_berufs_nr,4,5)) AS Fachklasse,
DBA.klasse.s_klasse_art AS Anlage,
DBA.noten_kopf.s_typ_nok AS HzJz,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE vorgang_schuljahr = '" + aktSj + "' AND s_art_fach <> 'U' AND schue_sj.s_typ_vorgang = 'A' AND (s_typ_nok = 'JZ' OR s_typ_nok = 'HZ') ORDER BY DBA.klasse.s_klasse_art ASC , DBA.klasse.klasse ASC; ", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        if (typ == theRow["HzJz"].ToString())
                        {
                            DateTime austrittsdatum = theRow["ausgetreten"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["ausgetreten"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                            Leistung leistung = new Leistung();

                            try
                            {
                                // Wenn der Schüler nicht in diesem Schuljahr ausgetreten ist ...

                                if (!(austrittsdatum > new DateTime(DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1, 8, 1) && austrittsdatum < DateTime.Now))
                                {
                                    leistung.LeistungId = Convert.ToInt32(theRow["LeistungId"]);
                                    leistung.ReligionAbgewählt = theRow["Religion"].ToString() == "N";
                                    leistung.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                                    leistung.Klasse = theRow["Klasse"].ToString();
                                    leistung.Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();
                                    leistung.Gesamtnote = theRow["Note"].ToString();
                                    leistung.Gesamtpunkte = theRow["Punkte"].ToString();
                                    leistung.EinheitNP = theRow["Einheit"].ToString() == "" ? "N" : theRow["Einheit"].ToString();
                                    leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString());
                                    leistung.HzJz = theRow["HzJz"].ToString();
                                    leistung.Anlage = theRow["Anlage"].ToString();
                                    leistung.Zeugnistext = theRow["Zeugnistext"].ToString();
                                    leistung.SchuelerAktivInDieserKlasse = theRow["SchuelerAktivInDieserKlasse"].ToString() == "J";                                    
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Fehler beim Einlesen der Atlantis-Leistungsdatensätze: ENTER" + ex);
                                Console.ReadKey();
                            }
                                this.Add(leistung);                                                  
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
        }

        internal void SprachenZuordnen(Leistungen atlantisLeistungen)
        {
            // Bei Sprachen-Fächern ...

            foreach (var we in (from t in this where (t.Fach.Split(' ')[0] == "EK" || t.Fach.Split(' ')[0] == "E" || t.Fach.Split(' ')[0] == "NL" || t.Fach.Split(' ')[0] == "LA" || t.Fach.Split(' ')[0] == "S" ) select t).ToList())
            {
                // ... wenn keine 1:1-Zuordnung möglich ist ...

                if (!(from a in atlantisLeistungen where a.Fach == we.Fach where a.Klasse == we.Klasse select a).Any())
                {
                    // wird versucht die Niveaustufe abzuschneiden

                    if (!(from a in atlantisLeistungen where a.Fach.Replace("A1","").Replace("A2", "").Replace("B1", "").Replace("B2", "").Replace("KA2", "") == we.Fach where a.Klasse == we.Klasse select a).Any())
                    {
                        // wird versucht den Kurs abzuschneiden

                        if ((we.Fach.Replace("  ", " ").Split(' ')).Count() > 1)
                        {                        
                            foreach (var a in atlantisLeistungen)
                            {
                                if (a.Klasse == we.Klasse)
                                {
                                    if (a.Fach.Split(' ')[0] == we.Fach.Split(' ')[0])
                                    {
                                        if ((a.Fach.Replace("  ", " ").Split(' ')).Count() > 1)
                                        {
                                            if (a.Fach.Replace("  ", " ").Split(' ')[1].Substring(0, 1) == we.Fach.Replace("  ", " ").Split(' ')[1].Substring(0, 1))
                                            {
                                                we.Fach = a.Fach;
                                            }
                                            else
                                            {
                                                Console.WriteLine("Das Fach " + we.Fach + " in der Klasse " + we.Klasse + " kann keinem Atlantis-Fach zugeordnet werden. ENTER");
                                                Console.ReadKey();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        we.Fach = (from a in atlantisLeistungen where a.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "").Replace("KA2", "") == we.Fach where a.Klasse == we.Klasse select a.Fach).FirstOrDefault();
                    }
                }
                else
                {
                    we.Fach = (from a in atlantisLeistungen where a.Fach == we.Fach where a.Klasse == we.Klasse select a.Fach).FirstOrDefault();
                }
            }
        }
        
        internal void BindestrichfächerZuordnen(Leistungen atlantisLeistungen)
        {
            // Bei Webuntis-Bindestrich-Fächern (z.B. VEF-GUP) 

            foreach (var we in (from t in this where t.Fach.Contains("-") select t).ToList())
            {
                // ... wenn keine 1:1-Zuordnung möglich ist ...

                if (!(from a in atlantisLeistungen where a.Fach == we.Fach select a).Any())
                {
                    // ... wird versucht das Fach aufzutrennen und dann zuzuordnen.

                    foreach (var fach in we.Fach.Split('-'))
                    {
                        var atlantisLeistung = (from a in atlantisLeistungen
                                                where a.Klasse == we.Klasse
                                                where a.Fach == fach
                                                where a.SchlüsselExtern == we.SchlüsselExtern
                                                select a).FirstOrDefault();

                        if (atlantisLeistung != null)
                        {
                            var x = (from wee in this where wee.Fach == we.Fach where wee.SchlüsselExtern == atlantisLeistung.SchlüsselExtern select wee).FirstOrDefault();
                            x.Fach = fach;
                        }
                    }
                }
            }
        }

        internal void Religionsabwähler(Leistungen atlantisLeistungen)
        {
            // Für die verschiednenen Klassen ...

            foreach (var klasse in (from k in this select k.Klasse).Distinct().ToList())
            {
                // ... wenn Religion in der Klasse unterrichtet wird ...

                if (
                    (from t in this
                     where (t.Fach == "KR" || t.Fach == "ER" || t.Fach == "REL" || t.Fach.StartsWith("KR ") || t.Fach.StartsWith("ER "))
                     where t.Klasse == klasse
                     select t).Any())
                {
                    // ... wird bei allen Atlantis-Reli-Abwählern ...

                    foreach (var aLeistung in (from a in atlantisLeistungen
                                                 where a.Klasse == klasse
                                                 where (a.Fach == "KR" || a.Fach == "ER" || a.Fach == "REL" || a.Fach.StartsWith("KR ") || a.Fach.StartsWith("ER "))
                                                 where a.ReligionAbgewählt
                                                 select a).ToList())
                    {
                        var wf = ((from w in this
                                  where w.Klasse == klasse
                                  where w.SchlüsselExtern == aLeistung.SchlüsselExtern
                                  where (w.Fach == "KR" || w.Fach == "ER" || w.Fach == "REL" || w.Fach.StartsWith("KR ") || w.Fach.StartsWith("ER "))
                                  select w).FirstOrDefault());
                        if (wf == null)
                        {
                            // ... ein neuer Webuntis-Datensatz hinzugefügt und ein '-' gesetzt.

                            Leistung leistung = new Leistung();
                            leistung.Datum = new DateTime();
                            leistung.Name = aLeistung.Name;
                            leistung.Klasse = aLeistung.Klasse;
                            leistung.Fach = aLeistung.Fach;
                            leistung.Gesamtnote = "-";
                            leistung.Bemerkung = "";
                            leistung.Benutzer = "";
                            leistung.SchlüsselExtern = aLeistung.SchlüsselExtern;
                            this.Add(leistung);
                        }
                        else
                        {
                            wf.Gesamtnote = "-";
                        }                        
                    }
                }
            }            
        }
        
        internal void ReligionKorrigieren()
        {
            foreach (var leistung in this)
            {
                if (leistung.Fach == "KR" || leistung.Fach == "ER" || leistung.Fach.StartsWith("KR ") || leistung.Fach.StartsWith("ER "))
                {
                    leistung.Fach = "REL";
                }
            }             
        }

        internal void WeitereFächerZuordnen(Leistungen atlantisLeistungen)
        {
            // Für die verschiednenen Klassen ...

            foreach (var klasse in (from k in this select k.Klasse).Distinct())
            {
                // ... und deren verschiedene Fächer aus Webuntis, in denen auch Noten gegeben wurden (also kein FU) ... 

                foreach (var webuntisFach in (from t in this where t.Klasse == klasse
                                              where t.Gesamtnote != ""
                                              where t.Fach != "KR" && t.Fach != "ER" && !t.Fach.StartsWith("KR ") && !t.Fach.StartsWith("ER ") && !t.Fach.Contains("-")
                                              select t.Fach).Distinct())
                {
                    // ... wenn es in Atlantis kein entsprechendes Fach gibt ...

                    if (!(from a in atlantisLeistungen where a.Fach == webuntisFach select a).Any())
                    {
                        // Bei Nicht-Bindestrich-Fächern wird geschaut, ob es ein Fach denselben Anfangsbuchstaben gibt, ...

                        var afa = (from a in atlantisLeistungen where a.Klasse == klasse where a.Fach.Substring(0, 1) == webuntisFach.Substring(0, 1) select a.Fach).Distinct().ToList();

                        var af = (from a in atlantisLeistungen where a.Klasse == klasse where a.Fach.Substring(0, 1) == webuntisFach.Substring(0, 1) select a).FirstOrDefault();

                        // ... und dass seinerseits in Webuntis nicht vorkommt ...

                        if (afa.Count == 1 && !(from w in this where w.Fach == af.Fach select w).Any())
                        {
                            string o = "/* Klasse: " + klasse + ". Das Webuntisfach " + webuntisFach + " wird dem Atlantisfach " + af.Fach + " zugeordnet.";
                            Global.Output.Add((o.Substring(0, Math.Min(82, o.Length))).PadRight(82) + "*/");

                            foreach (var leistung in this)
                            {
                                if (leistung.Fach == webuntisFach && leistung.Klasse == klasse)
                                {
                                    leistung.Fach = af.Fach;
                                }
                            }
                        }
                        else
                        {
                            string o = "/* Achtung: In der " + klasse + " kann das Fach " + webuntisFach + " keinem Atlantisfach zugeordnet werden!";
                            Global.Output.Add((o.Substring(0, Math.Min(82, o.Length))).PadRight(82) + "*/");
                        }
                    }
                }
            }
        }
        
        public Leistungen()
        {
        }

        internal void Add(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {
            int i = 0;

            Global.PrintMessage("Neu anzulegende Leistungen in Atlantis:");

            try
            {
                foreach (var klasse in interessierendeKlassen)
                {
                    foreach (var a in (from t in this where t.Klasse == klasse where t.SchuelerAktivInDieserKlasse select t).ToList())
                    {
                        var w = (from webuntisLeistung in webuntisLeistungen
                                 where webuntisLeistung.Fach == a.Fach
                                 where webuntisLeistung.Klasse == a.Klasse
                                 where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                 where a.Gesamtnote == ""
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null && w.Gesamtnote != "")
                        {
                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + w.Gesamtnote + "|" + a.Name, "UPDATE noten_einzel SET s_note='" + w.Gesamtnote + "' WHERE noe_id=" + a.LeistungId + ";");

                            if (a.EinheitNP == "P")
                            {
                                UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + w.Gesamtpunkte + "|" + a.Name, "UPDATE noten_einzel SET punkte='" + w.Gesamtpunkte + "' WHERE noe_id=" + a.LeistungId + ";");
                            }
                            i++;
                        }                       
                    }
                }
                
                if (i == 0)
                {
                    UpdateLeistung("                               ***keine***", "");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string Gesamtpunkte2Gesamtnote(string gesamtpunkte)
        {
            string gesamtnote = "";

            if (gesamtpunkte == "0")
            {
                gesamtnote = "6";
            }
            if (gesamtpunkte == "1")
            {
                gesamtnote = "5";
            }
            if (gesamtpunkte == "2")
            {
                gesamtnote = "5";
            }
            if (gesamtpunkte == "3")
            {
                gesamtnote = "5";
            }
            if (gesamtpunkte == "4")
            {
                gesamtnote = "4";
            }
            if (gesamtpunkte == "5")
            {
                gesamtnote = "4";
            }
            if (gesamtpunkte == "6")
            {
                gesamtnote = "4";
            }
            if (gesamtpunkte == "7")
            {
                gesamtnote = "3";
            }
            if (gesamtpunkte == "8")
            {
                gesamtnote = "3";
            }
            if (gesamtpunkte == "9")
            {
                gesamtnote = "3";
            }
            if (gesamtpunkte == "10")
            {
                gesamtnote = "2";
            }
            if (gesamtpunkte == "11")
            {
                gesamtnote = "2";
            }
            if (gesamtpunkte == "12")
            {
                gesamtnote = "2";
            }
            if (gesamtpunkte == "13")
            {
                gesamtnote = "1";
            }
            if (gesamtpunkte == "14")
            {
                gesamtnote = "1";
            }
            if (gesamtpunkte == "15")
            {
                gesamtnote = "1";
            }
            if (gesamtpunkte == "84")
            {
                gesamtnote = "A";
            }
            if (gesamtpunkte == "99")
            {
                gesamtnote = "-";
            }
            return gesamtnote;
        }

        internal void Update(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {
            int i = 0;

            Global.PrintMessage("Zu ändernde Noten in Atlantis:");

            try
            {
                foreach (var a in this)
                {
                    if (a.SchuelerAktivInDieserKlasse)
                    {
                        var w = (from webuntisLeistung in webuntisLeistungen
                                 where webuntisLeistung.Fach == a.Fach
                                 where webuntisLeistung.Klasse == a.Klasse
                                 where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                 where a.Gesamtnote != ""
                                 where a.Gesamtnote != webuntisLeistung.Gesamtnote
                                 where webuntisLeistung.Gesamtnote != ""
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null)
                        {
                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Gesamtnote + ">" + w.Gesamtnote + "|" + a.Name, "UPDATE noten_einzel SET s_note='" + w.Gesamtnote + "' WHERE noe_id=" + a.LeistungId + ";");

                            if (a.EinheitNP == "P")
                            {
                                UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Gesamtpunkte + ">" + w.Gesamtpunkte + "|" + a.Name, "UPDATE noten_einzel SET punkte='" + w.Gesamtpunkte + "' WHERE noe_id=" + a.LeistungId + ";");
                            }

                            i++;
                        }
                    }                    
                }
                if (i == 0)
                {
                    UpdateLeistung("                               ***keine***", "");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }

        internal void Delete(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {
            Global.PrintMessage("Zu löschende Noten in Atlantis:");

            int i = 0;

            try
            {
                foreach (var a in this)
                {
                    if (a.SchuelerAktivInDieserKlasse)
                    {
                        // Wenn es zu einem Atlantis-Datensatz keine Entsprechung in Webuntis gibt, ... 
                        
                        var webuntisLeistung = (from w in webuntisLeistungen
                                                where w.Fach == a.Fach
                                                where w.Klasse == a.Klasse
                                                where w.SchlüsselExtern == a.SchlüsselExtern
                                                where w.Gesamtpunkte != ""
                                                select w).FirstOrDefault();

                        if (webuntisLeistung == null)
                        {
                            if (a.Gesamtnote != "")
                            {
                                UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Name, "UPDATE noten_einzel SET s_note=NULL WHERE noe_id=" + a.LeistungId + ";");
                                if (a.EinheitNP == "P")
                                {
                                    UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Name, "UPDATE noten_einzel SET punkte=NULL WHERE noe_id=" + a.LeistungId + ";");
                                }
                                i++;
                            }
                        }
                    }
                }
                if (i == 0)
                {
                    UpdateLeistung("                               ***keine***", "");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal List<string> GetIntessierendeKlassen(Leistungen alleWebuntisLeistungen)
        {
            int auswahl;

            do
            {
                Console.WriteLine("");
                Console.WriteLine(" Bitte (in denen auch Gesamtnoten eingetragen sind) auswählen:");
                Console.WriteLine("");
                Console.WriteLine(" 1. Alle Klassen");
                Console.WriteLine(" 2. Alle Vollzeitklassen");
                Console.WriteLine(" 3. Alle Teilzeitklassen");
                Console.WriteLine(" 4. Bestimmte Klassen");
                
                Console.WriteLine("");
                Console.Write(" Ihre Auswahl " + (Properties.Settings.Default.Auswahl > 0 && Properties.Settings.Default.Auswahl < 5 ? "[ " + Properties.Settings.Default.Auswahl + " ] : " : ": "));
                
                var key = Console.ReadKey();
                

                if (int.TryParse(key.KeyChar.ToString(), out auswahl))
                {
                    if (Convert.ToInt32(key.KeyChar.ToString()) > 0 && Convert.ToInt32(key.KeyChar.ToString()) < 5)
                    {
                        Properties.Settings.Default.Auswahl = Convert.ToInt32(key.KeyChar.ToString());
                        Properties.Settings.Default.Save();
                        auswahl = Convert.ToInt32(key.KeyChar.ToString());
                        //Console.WriteLine(auswahl);
                    }
                }
                if (key.Key == ConsoleKey.Enter)
                {
                    auswahl = Properties.Settings.Default.Auswahl;
                }
                if (auswahl < 1 || auswahl > 4)
                {                    
                    Console.WriteLine("");
                    Console.Write(" ... Ungültige Auswahl! ");
                    Console.WriteLine("");
                }
            } while (auswahl < 1 || auswahl > 4);

            List<string> alleVerschiedenenUntisKlassenMitNoten = (from k in alleWebuntisLeistungen where k.Gesamtnote != "" where k.Klasse != null select k.Klasse).Distinct().ToList();
            List<string> alleVerschiedenenAtlantisKlassen = (from k in this select k.Klasse).Distinct().ToList();
            List<string> interessierendeKlassen = new List<string>();
                        
            foreach (var klasse in alleVerschiedenenUntisKlassenMitNoten)
            {
                if (auswahl == 1)
                {
                    interessierendeKlassen.Add(klasse);
                }
                if (auswahl == 2)
                {
                    if ((from a in this where a.Klasse == klasse where !a.Anlage.StartsWith("A") select a).Any())
                    {
                        interessierendeKlassen.Add(klasse);
                    }
                }
                if (auswahl == 3)
                {
                    if ((from a in this where a.Klasse == klasse where a.Anlage.StartsWith("A") select a).Any())
                    {
                        interessierendeKlassen.Add(klasse);
                    }
                }
            }

            if (auswahl == 4)
            {
                interessierendeKlassen.AddRange(KlassenAuswählen(alleVerschiedenenUntisKlassenMitNoten));
            }

            AusgabeSchreiben(interessierendeKlassen);

            return interessierendeKlassen;
        }

        private List<string> KlassenAuswählen(List<string> alleVerschiedenenUntisKlassenMitNoten)
        {
            List<string> klassenliste = new List<string>();
            string klassenlisteString = "";
            string eingabestring = "";

            do
            {
                Console.WriteLine("");
                Console.WriteLine("");
                Console.WriteLine("     Geben Sie einen oder mehrere Klassennamen oder auch deren Anfangsbuchstaben ein.");
                Console.WriteLine("");
                Console.Write("     Wählen Sie" + (Properties.Settings.Default.Klassenwahl == "" ? " : " : " [ " + Properties.Settings.Default.Klassenwahl + " ] : "));

                string eingaben = Console.ReadLine().ToUpper();

                if (eingaben == "")
                {
                    eingaben = Properties.Settings.Default.Klassenwahl;
                }
         
                foreach (var eingabe in eingaben.ToUpper().Split(','))
                {
                    foreach (var klasse in alleVerschiedenenUntisKlassenMitNoten)
                    {
                        if (klasse.StartsWith(eingabe))
                        {                            
                            klassenliste.Add(klasse);
                            klassenlisteString += klasse + ",";
                            if (!eingabestring.Contains("," +eingabe) && !eingabestring.StartsWith(eingabe))
                            {
                                eingabestring += eingabe + ",";
                            }                            
                        }
                    }                    
                }
                                
                Console.WriteLine("");

                if (klassenliste.Count == 0)
                {
                    Console.Write(" ... Ungültige Auswahl ! ");
                }
            } while (klassenliste.Count == 0);

            Properties.Settings.Default.Klassenwahl = eingabestring.TrimEnd(',');
            Properties.Settings.Default.Save();

            return klassenliste;
        }

        private void AusgabeSchreiben(List<string> interessierendeKlassen)
        {
            int z = 0;

            do
            {
                var zeile = "";

                try
                {
                    while ((zeile + interessierendeKlassen[z] + ", ").Length <= 78)
                    {
                        zeile += interessierendeKlassen[z] + ", ";
                        z++;
                    }
                }
                catch (Exception)
                {
                    z++;
                    zeile.TrimEnd(',');
                }

                zeile = zeile.TrimEnd(' ');

                string o = "/* " + zeile.TrimEnd(',');
                Global.Output.Add((o.Substring(0, Math.Min(82, o.Length))).PadRight(82) + "*/");

            } while (z < interessierendeKlassen.Count);
        }

        private void UpdateLeistung(string message, string updateQuery)
        {
            try
            {
                string o = updateQuery + "/*" + message;
                Global.Output.Add((o.Substring(0, Math.Min(82, o.Length))).PadRight(82) + "*/");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void ErzeugeSqlDatei(string outputSql)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(outputSql, true, Encoding.Default))
                {
                    foreach (var o in Global.Output)
                    {
                        writer.WriteLine(o);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            DateiÖffnen(outputSql);
        }

        private void DateiÖffnen(string pfad)
        {
            try
            {
                System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Notepad++\Notepad++.exe", pfad);
            }
            catch (Exception)
            {
                System.Diagnostics.Process.Start("Notepad.exe", pfad);
            }
        }
    }
}