using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;

namespace webuntisnoten2atlantis
{
    public class Abwesenheiten : List<Abwesenheit>
    {
        public List<string> output = new List<string>();

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
                            abwesenheit.StundenAbwesend = Convert.ToInt32(x[9]);
                            abwesenheit.StundenAbwesendUnentschuldigt = Convert.ToInt32(x[10]);

                            this.Add(abwesenheit);
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

        public Abwesenheiten(string connetionstringAtlantis, string aktSj, string prüfungsart, Schlüssels schlüssels)
        {
            try
            {
                Console.Write("Leistungsdaten aus Atlantis ".PadRight(70, '.'));

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
DBA.noten_kopf.s_art_nok AS Prüfungsart,
DBA.noten_kopf.fehlstunden_anzahl AS Fehlstunden,
DBA.noten_kopf.fehlstunden_ents_unents AS FehlstundenUnentschuldigt
FROM((DBA.schue_sj JOIN DBA.schueler ON DBA.schue_sj.pu_id = DBA.schueler.pu_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_kopf ON DBA.schueler.pu_id = DBA.noten_kopf.pu_id
WHERE /*schue_sj.pu_id = '152113' AND */
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
                        if (schlüssels.RichtigePrüfungsart(theRow["Prüfungsart"].ToString(), prüfungsart))
                        {
                            Abwesenheit abwesenheit = new Abwesenheit();
                            abwesenheit.StudentId = Convert.ToInt32(theRow["StudentId"]);
                            abwesenheit.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                            abwesenheit.Klasse = theRow["Klasse"].ToString();
                            abwesenheit.StundenAbwesend = theRow["Fehlstunden"].ToString() == "" ? 0 : Convert.ToDouble(theRow["Fehlstunden"]);
                            abwesenheit.StundenAbwesendUnentschuldigt = theRow["FehlstundenUnentschuldigt"].ToString() == "" ? 0 : Convert.ToDouble(theRow["FehlstundenUnentschuldigt"]);
                            this.Add(abwesenheit);
                        };                        
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
    }
}