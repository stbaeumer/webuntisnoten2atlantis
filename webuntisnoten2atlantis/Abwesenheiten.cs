// Published under the terms of GPLv3 Stefan Bäumer 2019.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;

namespace webuntisnoten2atlantis
{
    public class Abwesenheiten : List<Abwesenheit>
    {
        public Abwesenheiten(string inputAbwesenheitenCsv)
        {
            using (StreamReader reader = new StreamReader(inputAbwesenheitenCsv))
            {
                string überschrift = reader.ReadLine();

                Console.Write("Abwesenheiten aus Webuntis ".PadRight(70, '.'));

                while (true)
                {
                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            Abwesenheit abwesenheit = new Abwesenheit();
                            var x = line.Split('\t');
                            abwesenheit.StudentId = Convert.ToInt32(x[2]);
                            abwesenheit.Name = x[3] + " " + x[4];
                            abwesenheit.Klasse = x[5];
                            abwesenheit.StundenAbwesend = Convert.ToInt32(x[7])/45;
                            abwesenheit.StundenAbwesendUnentschuldigt = Convert.ToInt32(x[8])/45;

                            this.Add(abwesenheit);
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

        public Abwesenheiten(string connetionstringAtlantis, string aktSj)
        {
            try
            {
                Console.Write("Abwesenheiten aus Atlantis ".PadRight(70, '.'));

                using (OdbcConnection connection = new OdbcConnection(connetionstringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.schue_sj.vorgang_schuljahr,
DBA.klasse.klasse AS Klasse,
DBA.schue_sj.pu_id AS StudentId,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.noten_kopf.nok_id AS NotenkopfId,
DBA.noten_kopf.s_art_nok AS Zeugnisart,
DBA.noten_kopf.fehlstunden_anzahl AS Fehlstunden,
DBA.noten_kopf.fehlstunden_ents_unents AS FehlstundenUnentschuldigt
FROM((DBA.schue_sj JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_kopf ON DBA.schueler.pu_id = DBA.noten_kopf.pu_id
WHERE 
schue_sj.vorgang_schuljahr = '" + aktSj + @"' AND 
schue_sj.pj_id = noten_kopf.pj_id
ORDER BY DBA.schue_sj.vorgang_schuljahr ASC ,
DBA.klasse.klasse ASC ,
DBA.schueler.name_1 ASC ,
DBA.schueler.name_2 ASC ", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        Abwesenheit abwesenheit = new Abwesenheit();
                        abwesenheit.StudentId = Convert.ToInt32(theRow["StudentId"]);
                        abwesenheit.NotenkopfId = Convert.ToInt32(theRow["NotenkopfId"]);
                        abwesenheit.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                        abwesenheit.Klasse = theRow["Klasse"].ToString();
                        abwesenheit.StundenAbwesend = theRow["Fehlstunden"].ToString() == "" ? 0 : Convert.ToDouble(theRow["Fehlstunden"]);
                        abwesenheit.StundenAbwesendUnentschuldigt = theRow["FehlstundenUnentschuldigt"].ToString() == "" ? 0 : Convert.ToDouble(theRow["FehlstundenUnentschuldigt"]);
                        abwesenheit.Zeugnisart = theRow["Zeugnisart"].ToString();
                        this.Add(abwesenheit);
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

        internal List<Abwesenheit> Filter(List<string> zeugnisart)
        {
            Abwesenheiten abwesenheiten = new Abwesenheiten();

            foreach (var abwesenheit in this)
            {
                if (zeugnisart.Contains(abwesenheit.Zeugnisart))
                {
                    abwesenheiten.Add(abwesenheit);
                }
            }

            return abwesenheiten;
        }        

        internal List<Abwesenheit> Filter(List<string> zeugnisart, Abwesenheiten alleAtlantisAbwesenheitenGefiltert)
        {
            Abwesenheiten abwesenheiten = new Abwesenheiten();

            foreach (var abwesenheit in this)
            {
                abwesenheit.Zeugnisart = (from a in alleAtlantisAbwesenheitenGefiltert where a.StudentId == abwesenheit.StudentId select a.Zeugnisart).FirstOrDefault();

                if (zeugnisart.Contains(abwesenheit.Zeugnisart))
                {
                    abwesenheiten.Add(abwesenheit);
                }
            }

            return abwesenheiten;
        }

        public Abwesenheiten()
        {
        }

        internal void Add(List<Abwesenheit> webuntisAbwesenheiten)
        {
            int i = 0;

            Global.PrintMessage("Neu einzutragende Abwesenheiten in Atlantis:");

            try
            {
                foreach (var a in this)
                {  
                    var w = (from webuntisAbwesenheit in webuntisAbwesenheiten
                             where webuntisAbwesenheit.StudentId == a.StudentId
                             where webuntisAbwesenheit.Klasse == a.Klasse
                             where a.StundenAbwesend == 0
                             where webuntisAbwesenheit.StundenAbwesend > 0
                             select webuntisAbwesenheit).FirstOrDefault();

                    if (w != null)
                    {
                        UpdateAbwesenheit(a.Klasse + ",0->" + w.StundenAbwesend + "|" + a.Name, "UPDATE noten_kopf SET fehlstunden_anzahl=" + w.StundenAbwesend.ToString().PadLeft(3) + " WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
                }
                foreach (var a in this)
                {
                    var w = (from webuntisAbwesenheit in webuntisAbwesenheiten
                             where webuntisAbwesenheit.StudentId == a.StudentId
                             where webuntisAbwesenheit.Klasse == a.Klasse
                             where a.StundenAbwesendUnentschuldigt == 0
                             where webuntisAbwesenheit.StundenAbwesendUnentschuldigt > 0
                             select webuntisAbwesenheit).FirstOrDefault();

                    if (w != null)
                    {
                        UpdateAbwesenheit(a.Klasse + ",0->" + w.StundenAbwesendUnentschuldigt.ToString() + "|" + a.Name, "UPDATE noten_kopf SET fehlstunden_ents_unents=" + w.StundenAbwesendUnentschuldigt.ToString().PadLeft(3) + " WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
                }
                if (i == 0)
                {
                    UpdateAbwesenheit("                               ***keine***", "");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void Update(List<Abwesenheit> webuntisAbwesenheiten)
        {
            int i = 0;

            Global.PrintMessage("Zu ändernde Abwesenheiten in Atlantis:");

            try
            {
                foreach (var a in this)
                {
                    var w = (from webuntisAbwesenheit in webuntisAbwesenheiten
                             where webuntisAbwesenheit.Klasse == a.Klasse
                             where webuntisAbwesenheit.StudentId == a.StudentId
                             where a.StundenAbwesend != 0
                             where webuntisAbwesenheit.StundenAbwesend != 0
                             where a.StundenAbwesend != webuntisAbwesenheit.StundenAbwesend
                             select webuntisAbwesenheit).FirstOrDefault();

                    if (w != null)
                    {
                        UpdateAbwesenheit(a.Klasse + "|" + a.StundenAbwesend + "->" + w.StundenAbwesend + "|" + a.Name, "UPDATE noten_kopf SET fehlstunden_anzahl=" + w.StundenAbwesend.ToString() + " WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
                }
                foreach (var a in this)
                {
                    var w = (from webuntisAbwesenheit in webuntisAbwesenheiten
                             where webuntisAbwesenheit.Klasse == a.Klasse
                             where webuntisAbwesenheit.StudentId == a.StudentId
                             where a.StundenAbwesendUnentschuldigt != 0
                             where webuntisAbwesenheit.StundenAbwesendUnentschuldigt != 0
                             where a.StundenAbwesendUnentschuldigt != webuntisAbwesenheit.StundenAbwesendUnentschuldigt
                             select webuntisAbwesenheit).FirstOrDefault();
                    if (w != null)
                    {
                        UpdateAbwesenheit(a.Klasse + "|" + a.StundenAbwesendUnentschuldigt+"->"+ w.StundenAbwesendUnentschuldigt + "|" + a.Name, "UPDATE noten_kopf SET fehlstunden_ents_unents=" + w.StundenAbwesendUnentschuldigt + " WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
                }
                if (i == 0)
                {
                    UpdateAbwesenheit("                               ***keine***", "");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void Delete(List<Abwesenheit> webuntisAbwesenheiten)
        {
            int i = 0;

            Global.PrintMessage("Zu löschende Abwesenheiten in Atlantis:");

            try
            {
                foreach (var a in this)
                {
                    if ((from w in webuntisAbwesenheiten
                          where w.StudentId == a.StudentId
                          where w.Klasse == a.Klasse
                          where a.StundenAbwesend > 0
                          where w.StundenAbwesend == 0
                          select w).Any())
                    {
                        UpdateAbwesenheit(a.Klasse + "|" + a.StundenAbwesend +"->0" + "|" + a.Name, "UPDATE noten_kopf SET fehlstunden_anzahl=0 WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
                    if ((from w in webuntisAbwesenheiten
                          where w.StudentId == a.StudentId
                          where w.Klasse == a.Klasse
                          where a.StundenAbwesendUnentschuldigt > 0
                          where w.StundenAbwesendUnentschuldigt == 0
                          select w).Any())
                    {
                        UpdateAbwesenheit(a.Klasse + "|" + a.StundenAbwesend + "->0" + "|" + a.Name, "UPDATE noten_kopf SET fehlstunden_ents_unents=0 WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
                }
                if (i == 0)
                {
                    UpdateAbwesenheit("                               ***keine***", "");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void UpdateAbwesenheit(string message, string updateQuery)
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
    }
}