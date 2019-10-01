// published under the terms of GPLv3 Stefan Bäumer 2019

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace webuntisnoten2atlantis
{
    public class Leistungen : List<Leistung>
    {   
        public List<string> output = new List<string>();

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
                        Console.WriteLine("Hoppla. Da ist etwas schiefgelaufen. Die Verarbeitung wird hier abgebrochen.\n\n" + ex);
                        Console.ReadKey();
                        Environment.Exit(0);
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
            output.Add("/* Alle Noten aus Webuntis werden mit dieser Datei in Atlantis eingefügt */");
            output.Add("/* Hoping for the best! */");
            output.Add("/* " + System.Security.Principal.WindowsIdentity.GetCurrent().Name + " " + DateTime.Now.ToString() + " */");
            output.Add("");

            try
            {
                Console.Write("Schüler mit Abwesenheiten aus Atlantis ".PadRight(70, '.'));
                
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
                Console.WriteLine("Hoppla. Da ist etwas schiefgelaufen. Die Verarbeitung wird hier abgebrochen.\n\n" + ex);
                Console.ReadKey();
                Environment.Exit(0);
            }
            Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
        }
        
        public Leistungen()
        {
        }

        internal void NeuZuSetzendeNoten(Leistungen webuntisLeistungen)
        {
            output.Add("");
            output.Add("/* Neu einzutragende Noten in Atlantis: */");
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
                        Console.Write("[ADD] " + a.Klasse.PadRight(6) + a.Name.PadRight(30) + a.Fach.PadRight(7) + w.Note);

                        UpdateLeistung(a.Name, a.Klasse, a.Fach, w.Note, "UPDATE noten_einzel SET s_note=" + w.Note + " WHERE noe_id=" + a.LeistungId + ";");

                        Console.WriteLine(" ... ok");
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

                            Console.Write("[ADD] " + a.Klasse.PadRight(6) + a.Name.PadRight(30) + a.Fach.PadRight(7) + "-");

                            UpdateLeistung(a.Name, a.Klasse, a.Fach, "-", "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");

                            Console.WriteLine(" ... ok");
                        }
                    }
                }

                // Prüfen, ob ein Fach in Webuntis eine Note bekommen hat, zu dem es in Atlantis keinentsprechendes Fach gibt.

                foreach (var w in webuntisLeistungen)
                {
                    if (!(from a in this where a.Klasse == w.Klasse where a.Fach == w.Fach select a).Any())
                    {
                        UpdateLeistung(w.Name, w.Klasse, w.Fach, "-", "/*ACHTUNG: Das Fach " + w.Fach + " gibt es in Atlantis nicht!*/");
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void ZuÄnderendeNoten(Leistungen webuntisLeistungen)
        {
            output.Add("");
            output.Add("/* Zu ändernde Noten in Atlantis: */");
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
                        Console.Write("[UPD] " + a.Klasse.PadRight(6) + a.Name.PadRight(25) + a.Note + ">" + w.Note);

                        UpdateLeistung(a.Name, a.Klasse, a.Fach, a.Note + ">" + w.Note, "UPDATE noten_einzel SET s_note=" + w.Note + " WHERE noe_id=" + a.LeistungId + ";");

                        Console.WriteLine(" ... ok");
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

                        Console.Write("[UPD] " + a.Klasse.PadRight(6) + a.Name.PadRight(30) + a.Fach.PadRight(7) + "-");

                        UpdateLeistung(a.Name, a.Klasse, a.Fach, a.Note + "-", "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");

                        Console.WriteLine(" ... ok");
                    }                    
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }
        
        internal void ZuLöschendeNoten(Leistungen webuntisLeistungen)
        {
            output.Add("");
            output.Add("/* Zu löschende Noten in Atlantis: */");
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
                            Console.Write("[DEL] " + a.Klasse.PadRight(6) + a.Name.PadRight(30) + a.Fach.PadRight(7) + a.Note);

                            UpdateLeistung(a.Name, a.Klasse, a.Fach, "", "UPDATE noten_einzel SET s_note=NULL WHERE noe_id=" + a.LeistungId + ";");

                            Console.WriteLine(" ... ok");
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

                            Console.Write("[DEL] " + a.Klasse.PadRight(6) + a.Name.PadRight(30) + a.Fach.PadRight(7) + "-");

                            UpdateLeistung(a.Name, a.Klasse, a.Fach, "", "UPDATE noten_einzel SET s_note='-' WHERE noe_id=" + a.LeistungId + ";");

                            Console.WriteLine(" ... ok");

                        }                       
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }
        
        private void UpdateLeistung(string name, string klasse, string fach, string meldung, string updateQuery)
        {
            try
            {
                string o = updateQuery + "/*" + klasse + "," + fach + "," + meldung + "," + name;
                output.Add(o.Substring(0, Math.Min(82, o.Length - 1)) + "*/");
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
                    foreach (var o in output)
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