// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;

namespace webuntisnoten2atlantis
{
    public class Lehrers : List<Lehrer>
    {
        public Lehrers()
        {
        }

        public Lehrers(string connetionstringAtlantis, List<string> aktSj)
        {
            var typ = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

            try
            {
                using (OdbcConnection connection = new OdbcConnection(connetionstringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.lehr_sc.ls_id As Id,
DBA.lehrer.le_kuerzel As Kuerzel,
DBA.adresse.email As Mail
FROM(DBA.lehr_sc JOIN DBA.lehrer ON DBA.lehr_sc.le_id = DBA.lehrer.le_id) JOIN DBA.adresse ON DBA.lehrer.le_id = DBA.adresse.le_id
WHERE vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 0) + "/" + (Convert.ToInt32(aktSj[1]) - 0) + @"'; ", connection);
                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        Lehrer lehrer = new Lehrer();
                        try
                        {
                            lehrer.Kuerzel = theRow["Kuerzel"].ToString();
                            lehrer.AtlantisId = Convert.ToInt32(theRow["Id"]);
                            lehrer.Mail = theRow["Mail"].ToString();
                        }
                        catch (Exception)
                        {
                        }
                        if (lehrer.Mail.Contains("@berufskolleg-borken.de"))
                        {
                            this.Add(lehrer);
                        }                        
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            Console.WriteLine(("Lehrer*innen aus Atlantis (" + typ + ") ").PadRight(Global.PadRight, '.') + this.Count.ToString().PadLeft(4));            
        }        
    }
}