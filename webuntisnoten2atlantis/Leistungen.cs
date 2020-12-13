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
                                Console.WriteLine("In der Zeile " + i + " stimmt die Anzahl der Spalten nicht.");
                                Console.ReadKey();
                                DateiÖffnen(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MarksPerLesson.csv");
                                throw new Exception("In der Zeile " + i + " stimmt die Anzahl der Spalten nicht.");
                            }

                            leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            leistung.Name = x[1];
                            leistung.Klasse = x[2];
                            leistung.Fach = x[3];                            
                            leistung.Gesamtnote = x[9].Split('.')[0];
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
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.pu_id AS SchlüsselExtern,
DBA.schue_sj.s_religions_unterricht AS Religion,
(substr(schue_sj.s_berufs_nr,4,5)) AS Fachklasse,
DBA.klasse.s_klasse_art AS Anlage,
DBA.noten_kopf.s_typ_nok AS HzJz,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE vorgang_schuljahr = '" + aktSj + "' AND s_art_fach <> 'U' AND (s_typ_nok = 'JZ' OR s_typ_nok = 'HZ') ORDER BY DBA.klasse.s_klasse_art ASC , DBA.klasse.klasse ASC; ", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        if (typ == theRow["HzJz"].ToString())
                        {
                            Leistung leistung = new Leistung();
                            try
                            {
                                leistung.LeistungId = Convert.ToInt32(theRow["LeistungId"]);
                                leistung.ReligionAbgewählt = theRow["Religion"].ToString() == "N" ? true : false;
                                leistung.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                                leistung.Klasse = theRow["Klasse"].ToString();
                                leistung.Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();
                                leistung.Gesamtnote = theRow["Note"].ToString();
                                leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString());                                
                                leistung.HzJz = theRow["HzJz"].ToString();
                                leistung.Anlage = theRow["Anlage"].ToString();
                                leistung.Zeugnistext = theRow["Zeugnistext"].ToString();
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
                        // Bei Nicht-Bindestrich-Fächern wird geschaut, ob es ein Fach mit demselben Anfangsbuchstaben gibt, ...

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
                    foreach (var a in (from t in this where t.Klasse == klasse select t).ToList())
                    {
                        var w = (from webuntisLeistung in webuntisLeistungen
                                 where webuntisLeistung.Fach == a.Fach
                                 where webuntisLeistung.Klasse == a.Klasse
                                 where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                 where a.Gesamtnote == ""
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null && w.Gesamtnote != "")
                        {
                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|>" + w.Gesamtnote + "|" + a.Name, "UPDATE noten_einzel SET s_note='" + w.Gesamtnote + "' WHERE noe_id=" + a.LeistungId + ";");

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

        public void Punkte2NoteInAnlageNichtD(Leistungen atlantisLeistungen)
        {
            foreach (var leistung in (from t in this where t.Gesamtnote != "" select t).ToList())
            {
                if (!(from a in atlantisLeistungen where a.Klasse == leistung.Klasse select a.Anlage).FirstOrDefault().StartsWith("D"))
                {
                    if (leistung.Gesamtnote == "0")
                    {
                        leistung.Gesamtnote = "6";
                    }
                    if (leistung.Gesamtnote == "1")
                    {
                        leistung.Gesamtnote = "5";
                    }
                    if (leistung.Gesamtnote == "2")
                    {
                        leistung.Gesamtnote = "5";
                    }
                    if (leistung.Gesamtnote == "3")
                    {
                        leistung.Gesamtnote = "5";
                    }
                    if (leistung.Gesamtnote == "4")
                    {
                        leistung.Gesamtnote = "4";
                    }
                    if (leistung.Gesamtnote == "5")
                    {
                        leistung.Gesamtnote = "4";
                    }
                    if (leistung.Gesamtnote == "6")
                    {
                        leistung.Gesamtnote = "4";
                    }
                    if (leistung.Gesamtnote == "7")
                    {
                        leistung.Gesamtnote = "3";
                    }
                    if (leistung.Gesamtnote == "8")
                    {
                        leistung.Gesamtnote = "3";
                    }
                    if (leistung.Gesamtnote == "9")
                    {
                        leistung.Gesamtnote = "3";
                    }
                    if (leistung.Gesamtnote == "10")
                    {
                        leistung.Gesamtnote = "2";
                    }
                    if (leistung.Gesamtnote == "11")
                    {
                        leistung.Gesamtnote = "2";
                    }
                    if (leistung.Gesamtnote == "12")
                    {
                        leistung.Gesamtnote = "2";
                    }
                    if (leistung.Gesamtnote == "13")
                    {
                        leistung.Gesamtnote = "1";
                    }
                    if (leistung.Gesamtnote == "14")
                    {
                        leistung.Gesamtnote = "1";
                    }
                    if (leistung.Gesamtnote == "15")
                    {
                        leistung.Gesamtnote = "1";
                    }
                }
            }
        }

        internal void Update(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {
            int i = 0;

            Global.PrintMessage("Zu ändernde Noten in Atlantis:");

            try
            {
                foreach (var a in this)
                {   
                    var w = (from webuntisLeistung in webuntisLeistungen
                                where webuntisLeistung.Fach == a.Fach
                                where webuntisLeistung.Klasse == a.Klasse
                                where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                where a.Gesamtnote != ""
                                where a.Gesamtnote != webuntisLeistung.Gesamtnote
                                select webuntisLeistung).FirstOrDefault();

                    if (w != null)
                    {
                        UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Gesamtnote + ">" + w.Gesamtnote + "|" + a.Name, "UPDATE noten_einzel SET s_note='" + w.Gesamtnote + "' WHERE noe_id=" + a.LeistungId + ";");
                        i++;
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
                    // Wenn es zu einem Atlantis-Datensatz keine Entsprechung in Webuntis gibt, ... 

                    var webuntisLeistung = (from w in webuntisLeistungen
                                            where w.Fach == a.Fach
                                            where w.Klasse == a.Klasse
                                            where w.SchlüsselExtern == a.SchlüsselExtern
                                            select w).FirstOrDefault();

                    if (webuntisLeistung == null)
                    {
                        if (a.Gesamtnote != "")
                        {
                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Name, "UPDATE noten_einzel SET s_note=NULL WHERE noe_id=" + a.LeistungId + ";");
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

        internal List<string> GetIntessierendeKlassen(Leistungen alleWebuntisLeistungen)
        {
            List<string> alleKlassenVerschiedenenKlassen = (from k in this where k.Gesamtnote != "" select k.Klasse).Distinct().ToList();
                            
            List<string> interessierendeKlassen = new List<string>();

            try
            {
                Console.WriteLine("");
                Console.WriteLine("Sie müssen nun eine Auswahl der gewünschte(n) Klasse(n) treffen:.");
                Console.WriteLine("");
                Console.WriteLine("     '*'               gibt alle Noten aus aus allen Klassen mit Notenblatt aus.");
                Console.WriteLine("     '" + alleKlassenVerschiedenenKlassen[0] + "'          gibt alle Noten der Klasse " + alleKlassenVerschiedenenKlassen[0] + " aus.");
                if (alleKlassenVerschiedenenKlassen.Count > 1)
                {
                    Console.WriteLine("     '" + (alleKlassenVerschiedenenKlassen[0] + "," + alleKlassenVerschiedenenKlassen[1] + "'").PadRight(16) + " gibt alle Noten beider Klassen aus.");
                }                
                Console.WriteLine("     '" + alleKlassenVerschiedenenKlassen[0].Substring(0,1) + "'               gibt die Noten aller Klassen aus, die mit '" + alleKlassenVerschiedenenKlassen[0].Substring(0,1) + "' beginnen.");
                Console.WriteLine("     '" + alleKlassenVerschiedenenKlassen[0].Substring(0, 1) + ",G'             gibt aller Klassen aus, die mit '" + alleKlassenVerschiedenenKlassen[0].Substring(0, 1) + "' oder 'G' beginnen.");
                Console.WriteLine("");
                Console.Write(" Wählen Sie  " + (Properties.Settings.Default.Klassenwahl == "" ? " : " : "[ " + Properties.Settings.Default.Klassenwahl + " ] : "));
                
                string eingabe = Console.ReadLine().ToUpper();

                if (eingabe == "")
                {
                    eingabe = Properties.Settings.Default.Klassenwahl;
                }

                Properties.Settings.Default.Klassenwahl = "";
                Properties.Settings.Default.Save();
                Console.WriteLine("");

                if (eingabe != "")
                {
                    if (eingabe == "*")
                    {
                        interessierendeKlassen.AddRange(alleKlassenVerschiedenenKlassen);
                        Properties.Settings.Default.Klassenwahl = "*";
                        Properties.Settings.Default.Save();
                        if (alleKlassenVerschiedenenKlassen.Count > 0)
                        {
                            string alleKlassenString = "";

                            foreach (var iK in interessierendeKlassen)
                            {
                                alleKlassenString += iK + ",";
                            }

                            Console.WriteLine("Insgesamt ausgewählte Klassen (in denen auch Noten angelegt sind): " + interessierendeKlassen.Count);
                            Console.WriteLine("Die ausgewählten Klassen sind: " + alleKlassenString.TrimEnd(',') );

                            var anlagen = "";

                            foreach (var item in interessierendeKlassen)
                            {
                                foreach (var an in (from a in alleWebuntisLeistungen where a.Klasse == item select a.Anlage).ToList())
                                {
                                    if (!anlagen.Contains(an.Substring(0,3)))
                                    {
                                        anlagen += an.Substring(0,3) + ",";
                                    }
                                } 
                            }

                            Console.WriteLine("");
                            Console.WriteLine("Sie können nun die Verarbeitung mit ENTER starten oder die Anlagen einschränken");
                            Console.WriteLine("");

                            var iKNeu = new List<string>();

                            do
                            {
                                Console.Write(" Wählen Sie. '*' wählt alle Anlagen aus  " + (Properties.Settings.Default.Anlagen == "" ? "[ " + anlagen.TrimEnd(',') + " ] : " : "[ " + Properties.Settings.Default.Anlagen + "] : "));

                                var eingabeAnlagen = Console.ReadLine();
                                
                                if (eingabeAnlagen != "")
                                {   
                                    alleKlassenString = "";

                                    foreach (var item in interessierendeKlassen)
                                    {
                                        if ((from a in alleWebuntisLeistungen where a.Klasse == item where eingabeAnlagen.ToUpper().Split(',').Contains(a.Anlage.Substring(0, 3)) select a).Any() || eingabeAnlagen == "*")
                                        {
                                            iKNeu.Add(item);
                                            alleKlassenString += item + ",";
                                        }
                                    }

                                    if (iKNeu.Count > 0)
                                    {
                                        interessierendeKlassen = new List<string>();
                                        interessierendeKlassen = iKNeu;
                                        Console.WriteLine("");
                                        Console.WriteLine("Insgesamt ausgewählte Klassen (in denen auch Noten angelegt sind): " + interessierendeKlassen.Count + " ...");
                                        Console.WriteLine("Die ausgewählten Klassen sind: " + alleKlassenString.TrimEnd(','));
                                        Console.WriteLine("Die Verarbeitung startet ...");
                                        Properties.Settings.Default.Anlagen = anlagen;
                                        Properties.Settings.Default.Save();
                                    }
                                    else
                                    {
                                        Console.WriteLine("");
                                        Console.WriteLine("Es konnte keine einzige Klasse zu Ihrer Auswahl '" + eingabeAnlagen.ToUpper() + "' zugeordnet werden. Wiederholen Sie ...");
                                        Console.WriteLine("");
                                    }                                    
                                }
                            } while (iKNeu.Count == 0);                            
                        }
                        else
                        {
                            Console.WriteLine("Sie wollen die Leistungsdaten aller Klassen auslesen, aber keine einzige Klasse hat Leistungsdaten.");
                        }
                    }
                    else
                    {
                        var interessierendeKlassenString = "";
                        var klassenOhneNotenblattString = "";
                        var klassenOhneLeistungsdatensätzeString = "";

                        foreach (var klasse in alleKlassenVerschiedenenKlassen)
                        {
                            // Wenn die Klasse zur Eingabe passt ...

                            if ((from k in eingabe.Split(',') where klasse.StartsWith(k) select k).FirstOrDefault() != null)
                            {
                                // ... und auch Noten in Webuntis erfasst sind:

                                if ((from w in alleWebuntisLeistungen where w.Klasse == klasse select w).Any())
                                {
                                    interessierendeKlassen.Add(klasse);
                                    interessierendeKlassenString += klasse + ",";
                                    Properties.Settings.Default.Klassenwahl += klasse + ",";
                                    Properties.Settings.Default.Save();
                                }
                                else
                                {
                                    klassenOhneLeistungsdatensätzeString += klasse + ",";
                                }
                            }
                        }

                        Properties.Settings.Default.Klassenwahl = Properties.Settings.Default.Klassenwahl.TrimEnd(',');
                        Properties.Settings.Default.Save();

                        if (klassenOhneLeistungsdatensätzeString != "")
                        {
                            Global.Output.Add("/* [!] Folgende Klassen passen zu Ihrem Suchmuster. Allerdings liegen in Webuntis keine Leistungsdatensätze vor:\n    " + klassenOhneLeistungsdatensätzeString.TrimEnd(',') + "\n*/");
                        }

                        foreach (var w in (from w in alleWebuntisLeistungen select w.Klasse).Distinct())
                        {
                            if (!(from x in this where x.Klasse == w select x).Any())
                            {
                                if ((from i in interessierendeKlassen where i == w select i).Any())
                                {
                                    klassenOhneNotenblattString += w + ",";
                                }
                            }
                        }

                        if (klassenOhneNotenblattString != "")
                        {
                            Global.Output.Add("/* [!] Folgende Klassen passen zu Ihrem Suchmuster. Allerdings ist kein Notenblatt angelegt in Atlantis zugerdnet:\n    " + klassenOhneNotenblattString.TrimEnd(',') + "\n */");
                        }

                        if (interessierendeKlassenString == "")
                        {
                            Console.WriteLine("Es ist keine einzige Klasse bereit zur Verarbeitung. Ist kein Notenblatt angelegt? Ist der Klassenname korrekt?");
                        }
                        else
                        {
                            if (Properties.Settings.Default.Klassenwahl == "*")
                            {
                        
                            }
                            else
                            {
                                Console.WriteLine("Die Verarbeitung startet für ausgewählte Klasse" + (interessierendeKlassen.Count > 1 ? "n " : " ") + interessierendeKlassenString.TrimEnd(',') + " ...");
                            }                            
                        }
                        Console.WriteLine("");
                    }
                                        
                    return interessierendeKlassen;
                }
                else
                {
                    Console.WriteLine("");
                    Console.WriteLine("Ihre Auswahl: Alle " + (from a in this select a.Klasse).Distinct().Count().ToString().PadLeft(2) + " Klassen, in denen ein Notenblatt und ein Zeugnisformular angelegt ist.");
                    return alleKlassenVerschiedenenKlassen;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
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