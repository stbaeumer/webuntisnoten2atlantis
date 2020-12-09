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

                int i = 0;

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
                                throw new Exception();
                            }

                            leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            leistung.Name = x[1];
                            leistung.Klasse = x[2];
                            leistung.Fach = x[3] == "KR" || x[3] == "ER" ? "REL" : x[3];                            
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

        public Leistungen(string connetionstringAtlantis, string aktSj, Schlüssels schlüssels, List<string> zeugnisart)
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
DBA.noten_kopf.s_art_nok AS Zeugnisart,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE vorgang_schuljahr = '" + aktSj + "' AND s_art_fach = 'UF'; ", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        if (zeugnisart.Contains(theRow["Zeugnisart"].ToString()))
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

        internal Leistungen Filter(List<string> interessierendeKlassen)
        {
            Leistungen leistungen = new Leistungen();
            int i = 0;
            foreach (var leistung in this)
            {
                if (interessierendeKlassen.Contains(leistung.Klasse))
                {
                    if (!(from l in leistungen
                          where l.Klasse == leistung.Klasse
                          where l.Fach == leistung.Fach
                          where l.Name == leistung.Name
                          select l).Any())
                    {
                        leistungen.Add(leistung);
                        i++;
                    }                    
                }
            }
            return leistungen;
        }

        public Leistungen()
        {
        }

        internal void Add(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {
            Global.PrintMessage("Neu anzulegende Leistungen in Atlantis:");

            try
            {
                foreach (var a in this)
                {
                    if (interessierendeKlassen.Contains(a.Klasse))
                    {
                        var w = (from webuntisLeistung in webuntisLeistungen
                                 where webuntisLeistung.Fach == a.Fach
                                 where webuntisLeistung.Klasse == a.Klasse
                                 where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                 where a.Gesamtnote == ""
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null && w.Gesamtnote != "")
                        {
                            string note = a.Zeugnisart.StartsWith("D") ? w.Gesamtnote : Punkte2Note(w.Gesamtnote);

                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|0>" + note + "|" + a.Name, "UPDATE noten_einzel SET s_note=" + note + " WHERE noe_id=" + a.LeistungId + ";");
                        }
                        else
                        {
                            // Wenn es um Religion geht und der Schüler abgewählt hat und Religion in der Klasse unterrichtet wird ...

                            if (a.Fach == "REL" && a.ReligionAbgewählt && a.Gesamtnote == null && (from at in this
                                                                                             where at.Klasse == a.Klasse
                                                                                             where at.Fach == "REL"
                                                                                             where at.Gesamtnote != ""
                                                                                             where at.Gesamtnote != "-"
                                                                                             select at).Any())
                            {
                                // ... wird '-' gesetzt.

                                UpdateLeistung(a.Klasse + "|" + a.Fach + "|0>'-'" + "|" + a.Name, "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");
                            }
                        }
                    }                    
                }

                // Prüfen, ob ein Fach in Webuntis eine Note bekommen hat, zu dem es in Atlantis kein entsprechendes Fach gibt.

                foreach (var w in webuntisLeistungen)
                {
                    if (!(from a in this
                          where a.Klasse == w.Klasse
                          where a.Fach == w.Fach
                          select a).Any())
                    {
                        UpdateLeistung("","/* ACHTUNG: Das Fach " + w.Fach + " gibt es in Atlantis nicht! */");
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string Punkte2Note(string gesamtnote)
        {
            if (gesamtnote == "0")
            {
                return "6";
            }
            if (gesamtnote == "1")
            {
                return "5";
            }
            if (gesamtnote == "2")
            {
                return "5";
            }
            if (gesamtnote == "3")
            {
                return "5";
            }
            if (gesamtnote == "4")
            {
                return "4";
            }
            if (gesamtnote == "5")
            {
                return "4";
            }
            if (gesamtnote == "6")
            {
                return "4";
            }
            if (gesamtnote == "7")
            {
                return "3";
            }
            if (gesamtnote == "8")
            {
                return "3";
            }
            if (gesamtnote == "9")
            {
                return "3";
            }
            if (gesamtnote == "10")
            {
                return "2";
            }
            if (gesamtnote == "11")
            {
                return "2";
            }
            if (gesamtnote == "12")
            {
                return "2";
            }
            if (gesamtnote == "13")
            {
                return "1";
            }
            if (gesamtnote == "14")
            {
                return "1";
            }
            if (gesamtnote == "15")
            {
                return "1";
            }
            return gesamtnote;
        }

        internal void Update(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {
            Global.PrintMessage("Zu ändernde Noten in Atlantis:");

            try
            {
                foreach (var a in this)
                {
                    if (interessierendeKlassen.Contains(a.Klasse))
                    {
                        var w = (from webuntisLeistung in webuntisLeistungen
                                 where webuntisLeistung.Fach == a.Fach
                                 where webuntisLeistung.Klasse == a.Klasse
                                 where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                 where a.Gesamtnote != ""
                                 where a.Gesamtnote != (a.Zeugnisart.StartsWith("D") ? webuntisLeistung.Gesamtnote : Punkte2Note(webuntisLeistung.Gesamtnote))
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null && a.Fach != "REL")
                        {
                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Gesamtnote + ">" + w.Gesamtnote + "|" + a.Name, "UPDATE noten_einzel SET s_note=" + w.Gesamtnote + " WHERE noe_id=" + a.LeistungId + ";");
                        }

                        // Wenn es um Religion geht und der Schüler abgewählt hat und Religion in der Klasse unterrichtet wird ...

                        if (w != null && a.Fach == "REL" && a.ReligionAbgewählt && a.Gesamtnote == null && (from at in this
                                                                                                      where at.Klasse == a.Klasse
                                                                                                      where at.Fach == "REL"
                                                                                                      where at.Gesamtnote != ""
                                                                                                      where at.Gesamtnote != "-"
                                                                                                      select at).Any())
                        {
                            // ... wird '-' gesetzt.
                            try
                            {
                                UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Gesamtnote + ">" + w.Gesamtnote + "|" + a.Name, "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                        }
                    }                    
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
            try
            {
                foreach (var a in this)
                {
                    if (interessierendeKlassen.Contains(a.Klasse))
                    {
                        // Wenn es zu einem Atlantis-Datensatz keine Entsprechung in Webuntis gibt, ... 

                        var webuntisLeistung = (from w in webuntisLeistungen
                                                where w.Fach == a.Fach
                                                where w.Klasse == a.Klasse
                                                where w.SchlüsselExtern == a.SchlüsselExtern
                                                select w).FirstOrDefault();

                        if (webuntisLeistung == null)
                        {
                            // ... wird der Datensatz gelöscht, sofern es sich nicht um REL handelt und in Atlantis eine Note gesetzt ist.

                            if (a.Fach != "REL" && a.Gesamtnote != "")
                            {
                                UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Name, "UPDATE noten_einzel SET s_note=NULL WHERE noe_id=" + a.LeistungId + ";");
                            }

                            // Wenn es um Religion geht, der Schüler Religion abgewählt hat und kein '-' gesetzt ist und Religion in der Klasse unterrichtet wird ...

                            if (a.Fach == "REL" && a.ReligionAbgewählt && a.Gesamtnote != "-" && (from at in this
                                                                                            where at.Klasse == a.Klasse
                                                                                            where at.Fach == "REL"
                                                                                            where at.Gesamtnote != ""
                                                                                            where at.Gesamtnote != "-"
                                                                                            select at).Any())
                            {
                                // ... wird '-' gesetzt.

                                UpdateLeistung(a.Klasse + "|" + a.Fach + "|''>'-'|" + a.Name, "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");
                            }
                        }
                    }   
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }

        internal List<string> GetIntessierendeKlassen(Leistungen alleWebuntisLeistungen)
        {
            List<string> alleKlassen = (from k in this select k.Klasse).Distinct().ToList();
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
                Console.WriteLine("****************************************************************************************************");
                Console.WriteLine("*                                                                                                  *");
                Console.WriteLine("*  Geben Sie die interessierende(n) Klasse(n) ein (z. B. HH oder HHU oder HHU1). Dann ENTER.       *");
                Console.WriteLine("*  Oder ENTER drücken, um alle " + (from a in this select a.Klasse).Distinct().Count().ToString().PadLeft(3) + " Klassen mit angelegtem Notenblatt und                           *");
                Console.WriteLine("*  zugewiesenem Zeugnisformular zu wählen. Die Klassen sind:                                       *");
                Console.WriteLine("*  " + (alleKlassenString).PadRight(86) + "          *");
                Console.WriteLine("*                                                                                                  *");
                Console.WriteLine("****************************************************************************************************");
                Console.WriteLine("");
                Console.Write("Ihre Wahl: " + (Properties.Settings.Default.Klassenwahl == "" ? "" : "[" + Properties.Settings.Default.Klassenwahl + "]"));

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
                        Console.WriteLine("Es ist keine einzige Klasse bereit zur Verarbeitung.");
                        throw new Exception("Es ist keine einzige Klasse bereit zur Verarbeitung.");
                    }
                    else
                    {
                        Console.WriteLine("Die Verarbeitung startet für ausgewählte Klasse(n): " + interessierendeKlassenString.TrimEnd(',') + " ...");
                    }
                    Console.WriteLine("");
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
                        
            try
            {
                System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Notepad++\Notepad++.exe", outputSql);
            }
            catch (Exception)
            {
                System.Diagnostics.Process.Start("Notepad.exe", outputSql);
            }
        }
    }
}