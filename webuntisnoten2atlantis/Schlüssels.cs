// Published under the terms of GPLv3 Stefan Bäumer 2019.

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Data.Odbc;

namespace webuntisnoten2atlantis
{
    public class Schlüssels : List<Schlüssel>
    {
        private string wert;

        public Schlüssels(string connectionStringAtlantis)
        {
            try
            {
                Console.Write("Schlüssel aus Atlantis ".PadRight(70, '.'));

                using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.schluessel.aufloesung AS Auflösung,
DBA.schluessel.wert AS Wert
FROM DBA.schluessel
WHERE kennzeichen = 'NOK-ART'
; ", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        Schlüssel schlüssel = new Schlüssel();
                        schlüssel.Wert = theRow["Wert"].ToString(); // C03HJ
                        schlüssel.Auflösung = theRow["Auflösung"].ToString(); // Halbjaheszeugnis                        
                        this.Add(schlüssel);
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

        internal bool RichtigePrüfungsart(string wert, string auflösung)
        {
            return (from x in this
                    where x.Auflösung == auflösung
                    where x.Wert == wert
                    select x).Any();
        }
    }
}