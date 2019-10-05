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

                while (true)
                {
                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            Leistung leistung = new Leistung();
                            var x = line.Split('\t');

                            leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            leistung.Name = x[1];
                            leistung.Klasse = x[2];
                            leistung.Fach = x[3];
                            leistung.Prüfungsart = x[4];
                            leistung.Note = x[5].Substring(0, 1);
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

        public Leistungen(string connetionstringAtlantis, string aktSj, string prüfungsart, Schlüssels schlüssels)
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
SELECT DBA.schue_sj.vorgang_schuljahr,
DBA.klasse.klasse AS Klasse,
DBA.schue_sj.pu_id,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.noten_kopf.nok_id,
DBA.noten_kopf.position,
DBA.noten_kopf.s_typ_nok,
DBA.noten_kopf.s_art_nok AS Prüfungsart,
DBA.noten_kopf.dat_zeugnis,
DBA.noten_kopf.dat_notenkonferenz,
DBA.noten_kopf.fehlstunden_anzahl,
DBA.noten_einzel.noe_id AS LeistungId,
DBA.noten_einzel.nok_id,
DBA.noten_einzel.pu_id AS SchlüsselExtern,
DBA.noten_einzel.fa_id,
DBA.noten_einzel.kurztext AS Fach,
DBA.noten_einzel.s_note AS Note,
DBA.schue_sj.s_religions_unterricht AS Religion,
DBA.noten_kopf.pu_id,
DBA.noten_kopf.pj_id,
DBA.schue_sj.pj_id,
DBA.noten_einzel.position_1
FROM(((DBA.schue_sj JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_kopf ON DBA.schueler.pu_id = DBA.noten_kopf.pu_id ) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id
WHERE /*schue_sj.pu_id = '149565' AND*/ 
schue_sj.vorgang_schuljahr = '" + aktSj + @"' AND 
s_art_fach = 'UF' AND 
schue_sj.pj_id = noten_kopf.pj_id
ORDER BY 
DBA.schue_sj.vorgang_schuljahr ASC ,
DBA.klasse.klasse ASC ,
DBA.schueler.name_1 ASC ,
DBA.schueler.name_2 ASC,
DBA.noten_einzel.position_1 ASC; ", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        if (schlüssels.RichtigePrüfungsart(theRow["Prüfungsart"].ToString(), prüfungsart))
                        {
                            Leistung leistung = new Leistung()
                            {
                                LeistungId = Convert.ToInt32(theRow["LeistungId"]),
                                ReligionAbgewählt = theRow["Religion"].ToString() == "N" ? true : false,
                                Name = theRow["Nachname"] + " " + theRow["Vorname"],
                                Klasse = theRow["Klasse"].ToString(),
                                Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString(),
                                Prüfungsart = prüfungsart,
                                Note = theRow["Note"].ToString(),
                                SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString())
                            };
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
        
        public Leistungen()
        {
        }

        internal void Add(List<Leistung> webuntisLeistungen)
        {
            Global.PrintMessage("Neu anzulegende Leistungen in Atlantis:");

            try
            {
                foreach (var a in this)
                {
                    var w = (from webuntisLeistung in webuntisLeistungen
                             where webuntisLeistung.Fach == a.Fach
                             where webuntisLeistung.Klasse == a.Klasse
                             where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                             where a.Note == ""
                             select webuntisLeistung).FirstOrDefault();

                    if (w != null)
                    {
                        UpdateLeistung(a.Klasse + "|" + a.Fach + "|0>" + w.Note + "|" + a.Name, "UPDATE noten_einzel SET s_note=" + w.Note + " WHERE noe_id=" + a.LeistungId + ";");   
                    }
                    else
                    {
                        // Wenn es um Religion geht und der Schüler abgewählt hat und Religion in der Klasse unterrichtet wird ...

                        if (a.Fach == "REL" && a.ReligionAbgewählt && a.Note == null && (from at in this
                                                                                         where at.Klasse == a.Klasse
                                                                                         where at.Fach == "REL"
                                                                                         where at.Note != ""
                                                                                         where at.Note != "-"
                                                                                         select at).Any())
                        {
                            // ... wird '-' gesetzt.
                            
                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|0>'-'"+ "|" + a.Name, "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");                            
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
        
        internal void Update(List<Leistung> webuntisLeistungen)
        {
            Global.PrintMessage("Zu ändernde Noten in Atlantis:");

            try
            {
                foreach (var a in this)
                {
                    var w = (from webuntisLeistung in webuntisLeistungen
                             where webuntisLeistung.Fach == a.Fach
                             where webuntisLeistung.Klasse == a.Klasse
                             where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                             where a.Note != ""
                             where a.Note != webuntisLeistung.Note
                             select webuntisLeistung).FirstOrDefault();

                    if (w != null && a.Fach != "REL")
                    {
                        UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Note + ">" + w.Note + "|" + a.Name, "UPDATE noten_einzel SET s_note=" + w.Note + " WHERE noe_id=" + a.LeistungId + ";");
                    }
                    
                    // Wenn es um Religion geht und der Schüler abgewählt hat und Religion in der Klasse unterrichtet wird ...

                    if (a.Fach == "REL" && a.ReligionAbgewählt && a.Note == null && (from at in this
                                                                                        where at.Klasse == a.Klasse
                                                                                        where at.Fach == "REL"
                                                                                        where at.Note != ""
                                                                                        where at.Note != "-"
                                                                                        select at).Any())
                    {
                        // ... wird '-' gesetzt.
                        
                        UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Note + ">" + w.Note + "|" + a.Name, "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");                        
                    }                    
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }
        
        internal void Delete(List<Leistung> webuntisLeistungen)
        {
            Global.PrintMessage("Zu löschende Noten in Atlantis:");
            try
            {
                foreach (var a in this)
                {
                    // Wenn es zu einem Atlantis-Datensatz keine Entsprechung in Webuntis gibt, ... 

                    if (!(from w in webuntisLeistungen
                          where w.Fach == a.Fach
                          where w.Klasse == a.Klasse
                          where w.SchlüsselExtern == a.SchlüsselExtern
                          select w).Any())
                    {
                        // ... wird der Datensatz gelöscht, sofern es sich nicht um REL handelt und in Atlantis eine Note gesetzt ist.

                        if (a.Fach != "REL" && a.Note != "")
                        {
                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|" + a.Name, "UPDATE noten_einzel SET s_note=NULL WHERE noe_id=" + a.LeistungId + ";");
                        }

                        // Wenn es um Religion geht, der Schüler Religion abgewählt hat und kein '-' gesetzt ist und Religion in der Klasse unterrichtet wird ...

                        if (a.Fach == "REL" && a.ReligionAbgewählt && a.Note != "-" && (from at in this
                                                                                        where at.Klasse == a.Klasse
                                                                                        where at.Fach == "REL"
                                                                                        where at.Note != ""
                                                                                        where at.Note != "-"
                                                                                        select at).Any())
                        {
                            // ... wird '-' gesetzt.
                            
                            UpdateLeistung(a.Klasse + "|" + a.Fach + "|''>'-'|" + a.Name, "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");                            
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
            List<string> interessierendeKlassen = new List<string>();

            try
            {
                Console.WriteLine("****************************************************************************************************");
                Console.WriteLine("*                                                                                                  *");
                Console.WriteLine("*  Geben Sie die interessierende(n) Klasse(n) ein (z. B. HH oder HHU oder HHU1). Dann ENTER.       *");
                Console.WriteLine("*  Oder ENTER drücken, um alle " + (from a in this select a.Klasse).Distinct().Count().ToString().PadLeft(3) + " Klassen mit angelegtem Notenblatt und                           *");
                Console.WriteLine("*  zugewiesenem " + alleWebuntisLeistungen[0].Prüfungsart + "-Zeugnisformular zu wählen:                                       *");
                Console.WriteLine("*                                                                                                  *");
                Console.WriteLine("****************************************************************************************************");
                Console.WriteLine("");
                Console.Write("Ihre Wahl: ");

                string eingabe = Console.ReadLine().ToUpper();

                Console.WriteLine("");

                if (eingabe != "")
                {
                    var interessierendeKlassenString = "";
                    var klassenOhneNotenblattString = "";
                    var klassenOhneLeistungsdatensätzeString = "";

                    foreach (var klasse in alleKlassen)
                    {
                        // Wenn die Klasse zur Eingabe passt ...

                        if ((from k in eingabe.Split(',') where klasse.Contains(k) select k).FirstOrDefault() != null)
                        {
                            // ... und auch Noten in Webuntis erfasst sind:

                            if ((from w in alleWebuntisLeistungen where w.Klasse == klasse select w).Any())
                            {
                                interessierendeKlassen.Add(klasse);
                                interessierendeKlassenString += klasse + ",";
                            }
                            else
                            {
                                klassenOhneLeistungsdatensätzeString += klasse + ",";
                            }
                        }
                    }

                    if (klassenOhneLeistungsdatensätzeString != "")
                    {
                        Global.Output.Add("/* [!] Folgende Klassen passen zu Ihrem Suchmuster. Allerdings liegen in Webuntis keine Leistungsdatensätze vor:\n    " + klassenOhneLeistungsdatensätzeString.TrimEnd(',') + "\n*/");
                    }

                    foreach (var w in (from w in alleWebuntisLeistungen select w.Klasse).Distinct())
                    {
                        if (!(from x in this where x.Klasse == w select x).Any())
                        {
                            klassenOhneNotenblattString += w + ",";
                        }
                    }

                    if (klassenOhneNotenblattString != "")
                    {
                        Global.Output.Add("/* [!] Folgende Klassen passen zu Ihrem Suchmuster. Allerdings ist kein Notenblatt bzw. Zeugnisformular in Atlantis zugerdnet:\n    " + klassenOhneNotenblattString.TrimEnd(',') + "\n */");
                    }

                    if (interessierendeKlassenString == "")
                    {
                        Console.WriteLine("Es ist keine einzige Klasse bereit zur Verarbeitung.");
                    }
                    else
                    {
                        Console.WriteLine("Die Verarbeitung startet für ausgewählte Klassen: " + interessierendeKlassenString.TrimEnd(',') + " ...");
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