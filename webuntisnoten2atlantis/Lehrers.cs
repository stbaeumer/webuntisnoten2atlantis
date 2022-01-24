// Published under the terms of GPLv3 Stefan Bäumer 2021.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;

namespace webuntisnoten2atlantis
{
    internal class Lehrers : List<Lehrer>
    {
        public Lehrers()
        {
        }

        public Lehrers(string connetionstringAtlantis, List<string> aktSj)
        {
            try
            {
                var typ = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

                Console.Write(("Lehrer aus Atlantis (" + typ + ")").PadRight(71, '.'));

                using (OdbcConnection connection = new OdbcConnection(connetionstringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.adresse.email AS Mail,
DBA.lehrer.le_kuerzel AS Kuerzel
FROM DBA.lehrer JOIN DBA.adresse ON DBA.lehrer.le_id = DBA.adresse.le_id;", connection);
                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        Lehrer lehrer = new Lehrer();
                        try
                        {
                            lehrer.Kuerzel = theRow["Kuerzel"].ToString();
                            lehrer.Mail = theRow["Mail"].ToString();
                        }
                        catch (Exception ex)
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
            Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));            
        }        
    }
}