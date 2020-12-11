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
            using (StreamReader reader = new StreamReader(datei))
            {
                string überschrift = reader.ReadLine();

                Console.Write("Leistungsdaten aus Webuntis ".PadRight(70, '.'));

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
                            this.Add(leistung);                           
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
                Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
            }
        }

        public Leistungen(string connetionstringAtlantis, string aktSj)
        {
            Global.Output.Add("/* ****************************************************************************** */");
            Global.Output.Add("/* Diese Datei enthält alle Noten und Fehlzeiten aus Webuntis.                    */");
            Global.Output.Add("/* Sie können alle Noten und Fehlzeiten aus Webuntis nach Atlantis importieren,   */");
            Global.Output.Add("/* indem Sie diese Datei in Atlantis unter Funktionen>SQL-Anweisung hochladen.    */");
            Global.Output.Add("/* Published under the terms of GPLv3. Hoping for the best!                       */");
            Global.Output.Add("/* " + (System.Security.Principal.WindowsIdentity.GetCurrent().Name + " " + DateTime.Now.ToString()).PadRight(78) + " */");
            Global.Output.Add("/* ****************************************************************************** */");
            Global.Output.Add("");

            try
            {
                Console.Write("Leistungsdaten aus Atlantis ".PadRight(70, '.'));

                using (OdbcConnection connection = new OdbcConnection(connetionstringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.noten_einzel.noe_id AS LeistungId,
DBA.noten_einzel.fa_id,
DBA.noten_einzel.kurztext AS Fach,
DBA.noten_einzel.zeugnistext,
DBA.noten_einzel.s_note AS Note,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.pu_id AS SchlüsselExtern,
DBA.schue_sj.s_religions_unterricht AS Religion,
(substr(schue_sj.s_berufs_nr,4,5)) AS Fachklasse,
DBA.noten_kopf.s_art_nok AS Zeugnisart,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE vorgang_schuljahr = '" + aktSj + "' AND s_art_fach = 'UF'; ", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
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
                            leistung.Zeugnisart = theRow["Zeugnisart"].ToString();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Fehler beim Einlesen der Atlantis-Leistungsdatensätze: ENTER");
                            Console.ReadKey();
                        }

                        this.Add(leistung);
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

        internal List<Leistung> Filter(List<string> zeugnisart, Leistungen alleAtlantisLeistungenGefiltert)
        {
            Leistungen leistungen = new Leistungen();

            foreach (var leistung in this)
            {
                leistung.Zeugnisart = (from a in alleAtlantisLeistungenGefiltert where a.SchlüsselExtern == leistung.SchlüsselExtern where a.Fach == leistung.Fach select a.Zeugnisart).FirstOrDefault();

                if (zeugnisart.Contains(leistung.Zeugnisart))
                {                       
                    leistungen.Add(leistung);
                }
            }

            return leistungen;
        }

        internal void BindestrichfächerZuordnen(Leistungen atlantisLeistungen)
        {
            // Bei Webuntis-Bindestrich-Fächern (z.B. VEF-GUP) 

            foreach (var we in (from t in this where t.Fach.Contains("-") select t).ToList())
            {
                // ... wenn keine 1:1-Zuordnung möglich ist ...

                if (!(from a in atlantisLeistungen where a.Fach == we.Fach select a).Any())
                {
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
                // ... wenn Religion unterrichtet wird ...

                if ((from t in this where (t.Fach == "KR" || t.Fach == "ER" || t.Fach == "REL") where t.Klasse == klasse select t).Any())
                {
                    // ... wird bei allen Atlantis-Reli-Abwählern ...

                    foreach (var aLeistung in (from a in atlantisLeistungen
                                                 where a.Klasse == klasse
                                                 where (a.Fach == "KR" || a.Fach == "ER" || a.Fach == "REL")
                                                 where a.ReligionAbgewählt
                                                 select a).ToList())
                    {
                        var wf = ((from w in this
                                  where w.Klasse == klasse
                                  where w.SchlüsselExtern == aLeistung.SchlüsselExtern
                                  where (w.Fach == "KR" || w.Fach == "ER" || w.Fach == "REL")
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
                if (leistung.Fach == "KR" || leistung.Fach == "ER")
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
                // ... unde deren verschiedene Fächer aus Webuntis, in denen auch Noten gegeben wurden (also kein FU) ... 

                foreach (var webuntisFach in (from t in this where t.Klasse == klasse
                                              where t.Gesamtnote != ""
                                              where t.Fach != "KR" && t.Fach != "ER" && !t.Fach.Contains("-")
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
                            string o = "/* Achtung: In der " + webuntisFach + " kann das Fach " + webuntisFach + " keinem Atlantisfach zugeordnet werden!";
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
                foreach (var a in this)
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

        public void Punkte2NoteInAnlageD(Leistungen atlantisLeistungen)
        {
            foreach (var leistung in (from t in this where t.Gesamtnote != "" select t).ToList())
            {
                if (!(from a in atlantisLeistungen where a.Klasse == leistung.Klasse select a.Zeugnisart).FirstOrDefault().StartsWith("D"))
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

        internal List<string> GetIntessierendeKlassen(Leistungen alleWebuntisLeistungen, List<string> zeugnisart)
        {
            List<string> alleKlassen = (from k in this where zeugnisart.Contains(k.Zeugnisart) select k.Klasse).Distinct().ToList();
            string alleKlassenString = "";

            foreach (var kl in alleKlassen)
            {
                alleKlassenString += kl + ",";
            }

            alleKlassenString = alleKlassenString.TrimEnd(',');
            alleKlassenString = alleKlassenString.Substring(0, Math.Min(alleKlassenString.Length, 40));
                
            List<string> interessierendeKlassen = new List<string>();

            try
            {
                Console.WriteLine("Geben Sie die gewünschte(n) Klasse(n) (kommagetrennt) ein.");
                Console.WriteLine("Oder '*' für alle Klassen der gewählten Zeugnisart(en), bei denen zuvor ein Notenblatt angelegt wurde.");
                Console.WriteLine("Beispiel: HBW20A oder HB,HG oder GW18A,GG oder '*') " + (Properties.Settings.Default.Klassenwahl == "" ? "" : "[" + Properties.Settings.Default.Klassenwahl + "] "));
                
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
                        interessierendeKlassen.AddRange(alleKlassen);
                        Properties.Settings.Default.Klassenwahl = "*";
                        Properties.Settings.Default.Save();
                    }
                    else
                    {
                        var interessierendeKlassenString = "";
                        var klassenOhneNotenblattString = "";
                        var klassenOhneLeistungsdatensätzeString = "";

                        foreach (var klasse in alleKlassen)
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
                            Global.Output.Add("/* [!] Folgende Klassen passen zu Ihrem Suchmuster. Allerdings ist kein Notenblatt bzw. Zeugnisformular in Atlantis zugerdnet:\n    " + klassenOhneNotenblattString.TrimEnd(',') + "\n */");
                        }

                        if (interessierendeKlassenString == "")
                        {
                            Console.WriteLine("Es ist keine einzige Klasse bereit zur Verarbeitung. Ist kein Notenblatt angelegt? Ist der Klassenname korrekt? Ist keine Zeugnisart zugeordnet?");

                        }
                        else
                        {
                            Console.WriteLine("Die Verarbeitung startet für ausgewählte Klasse(n): " + interessierendeKlassenString.TrimEnd(',') + " ...");
                        }
                        Console.WriteLine("");
                    }
                                        
                    return interessierendeKlassen;
                }
                else
                {
                    Console.WriteLine("");
                    Console.WriteLine("Ihre Auswahl: Alle " + (from a in this select a.Klasse).Distinct().Count().ToString().PadLeft(2) + " Klassen, in denen ein Notenblatt und ein Zeugnisformular angelegt ist.");
                    return alleKlassen;
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