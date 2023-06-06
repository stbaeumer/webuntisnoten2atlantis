// Published under the terms of GPLv3 Stefan Bäumer 2023.

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
        public Abwesenheiten(string inputAbwesenheitenCsv, string interessierendeKlasse, Leistungen alleWebuntisleistungen)
        {
            using (StreamReader reader = new StreamReader(inputAbwesenheitenCsv))
            {
                string überschrift = reader.ReadLine();

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
                            abwesenheit.StundenAbwesend = Convert.ToInt32(x[7]) / 45;
                            abwesenheit.StundenAbwesendUnentschuldigt = Convert.ToInt32(x[8]) / 45; // unenentschuldigt oder offen

                            if (interessierendeKlasse == abwesenheit.Klasse)
                            {
                                this.Add(abwesenheit);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        throw new Exception("Die Datei " + inputAbwesenheitenCsv + " kann nicht gelesen werden.");
                    }

                    if (line == null)
                    {
                        break;
                    }
                }
                Global.WriteLine("Abwesenheiten aus Webuntis ".PadRight(Global.PadRight, '.') + this.Count.ToString().PadLeft(4));

                // Für Schüler in Webuntis ohne Abwesenheit wird kein Datensatz exportiert.
                // Damit ein möglicherweise händisch angelegter Datenatz in Fucklantis überschrieben wird,
                // muss hier für jeden Schüler ohne Datensatz einer angelegt werden:

                foreach (var item in alleWebuntisleistungen)
                {
                    if (!(from t in this where t.StudentId == item.SchlüsselExtern select t).Any())
                    {
                        Abwesenheit abwesenheit = new Abwesenheit();                        
                        abwesenheit.StudentId = item.SchlüsselExtern;
                        abwesenheit.Name = item.Name;
                        abwesenheit.Klasse = item.Klasse;
                        abwesenheit.StundenAbwesend = 0;
                        abwesenheit.StundenAbwesendUnentschuldigt = 0; // unenentschuldigt oder offen
                        this.Add(abwesenheit);
                    }
                }
            }
        }

        public Abwesenheiten(string connetionstringAtlantis, List<string> aktSj, string interessierendeKlasse)
        {
            try
            {
                var typ = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

                Global.WriteLine("Die Atlantis-Abwesenheiten & -Leistungen beziehen sich auf den Abschnitt: ".PadRight(Global.PadRight, '.') + "  " + typ);

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
DBA.noten_kopf.s_typ_nok AS HzJz,
DBA.noten_kopf.s_art_nok AS Zeugnisart,
DBA.noten_kopf.fehlstunden_anzahl AS Fehlstunden,
DBA.noten_kopf.fehlstunden_ents_unents AS FehlstundenUnentschuldigt
FROM((DBA.schue_sj JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_kopf ON DBA.schueler.pu_id = DBA.noten_kopf.pu_id
WHERE 
schue_sj.vorgang_schuljahr = '" + aktSj[0] + "/" + aktSj[1] + @"' AND schue_sj.pj_id = noten_kopf.pj_id AND schue_sj.s_typ_vorgang = 'A'
ORDER BY DBA.schue_sj.vorgang_schuljahr ASC ,
DBA.klasse.klasse ASC ,
DBA.schueler.name_1 ASC ,
DBA.schueler.name_2 ASC ", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        if (typ == theRow["HzJz"].ToString() || theRow["HzJz"].ToString() == "GO")
                        {
                            Abwesenheit abwesenheit = new Abwesenheit();
                            abwesenheit.StudentId = Convert.ToInt32(theRow["StudentId"]);
                            abwesenheit.NotenkopfId = Convert.ToInt32(theRow["NotenkopfId"]);
                            abwesenheit.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                            abwesenheit.Klasse = theRow["Klasse"].ToString();
                            abwesenheit.StundenAbwesend = theRow["Fehlstunden"].ToString() == "" ? 0 : Convert.ToDouble(theRow["Fehlstunden"]);
                            abwesenheit.StundenAbwesendUnentschuldigt = theRow["FehlstundenUnentschuldigt"].ToString() == "" ? 0 : Convert.ToDouble(theRow["FehlstundenUnentschuldigt"]);
                            abwesenheit.Zeugnisart = theRow["Zeugnisart"].ToString();
                            abwesenheit.HzJz = theRow["HzJz"].ToString();

                            if (interessierendeKlasse == abwesenheit.Klasse)
                            {
                                this.Add(abwesenheit);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            Global.WriteLine("Abwesenheiten aus Atlantis ".PadRight(Global.PadRight, '.') + this.Count.ToString().PadLeft(4));            
        }

        public Abwesenheiten()
        {
        }

        internal void Add(List<Abwesenheit> webuntisAbwesenheiten)
        {
            int outputIndex = Global.SqlZeilen.Count();

            int i = 0;

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
                        UpdateAbwesenheit(a.Klasse.PadRight(6) + "(" + a.HzJz + ")" + "|0->" + w.StundenAbwesend.ToString().PadRight(3) + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 12)) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_kopf SET fehlstunden_anzahl=" + w.StundenAbwesend.ToString().PadLeft(3) + " WHERE nok_id=" + a.NotenkopfId + ";");
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
                        UpdateAbwesenheit(a.Klasse.PadRight(6) + "(" + a.HzJz + ")" + "|0->" + w.StundenAbwesendUnentschuldigt.ToString().PadRight(3) + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 12)) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_kopf SET fehlstunden_ents_unents=" + w.StundenAbwesendUnentschuldigt.ToString().PadLeft(3) + " WHERE nok_id=" + a.NotenkopfId + ";");
                        Global.Rückmeldungen.AddRückmeldung(new Rückmeldung("Klassenleitung", "", "(Unentschuldigte) Fehlstunden so ok?"));
                        
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void Update(List<Abwesenheit> webuntisAbwesenheiten)
        {
            int outputIndex = Global.SqlZeilen.Count();

            int i = 0;

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
                        UpdateAbwesenheit(a.Klasse.PadRight(6) + "(" + a.HzJz + ")" + "|" + a.StundenAbwesend.ToString().PadLeft(3) + "->" + w.StundenAbwesend.ToString().PadLeft(3) + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 12)) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_kopf SET fehlstunden_anzahl=" + w.StundenAbwesend.ToString().PadRight(3) + " WHERE nok_id=" + a.NotenkopfId + ";");
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
                        UpdateAbwesenheit(a.Klasse.PadRight(6) + "(" + a.HzJz + ")" + "|" + a.StundenAbwesendUnentschuldigt.ToString().PadLeft(3) + "->" + w.StundenAbwesendUnentschuldigt.ToString().PadLeft(3) + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 12)) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_kopf SET fehlstunden_ents_unents=" + w.StundenAbwesendUnentschuldigt.ToString().PadRight(3) + " WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void Delete(List<Abwesenheit> webuntisAbwesenheiten)
        {
            int outputIndex = Global.SqlZeilen.Count();
            int i = 0;

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
                        UpdateAbwesenheit(a.Klasse.PadRight(6) + "|" + a.StundenAbwesend.ToString().PadLeft(3) + "->0" + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 12)) + "(" + a.HzJz + ")" + a.Beschreibung, "UPDATE noten_kopf SET fehlstunden_anzahl=0 WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
                    if ((from w in webuntisAbwesenheiten
                         where w.StudentId == a.StudentId
                         where w.Klasse == a.Klasse
                         where a.StundenAbwesendUnentschuldigt > 0
                         where w.StundenAbwesendUnentschuldigt == 0
                         select w).Any())
                    {
                        UpdateAbwesenheit(a.Klasse.PadRight(6) + "|" + a.StundenAbwesend.ToString().PadLeft(3) + "->0" + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 12)) + "(" + a.HzJz + ")" + a.Beschreibung, "UPDATE noten_kopf SET fehlstunden_ents_unents=0 WHERE nok_id=" + a.NotenkopfId + ";");
                        i++;
                    }
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
                string o = updateQuery.PadRight(103) + "/* " + message;
                Global.SqlZeilen.Add((o.Substring(0, Math.Min(125, o.Length))).PadRight(Global.PadRight + 31) + "*/");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}