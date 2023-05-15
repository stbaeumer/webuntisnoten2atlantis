// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.IO;

namespace webuntisnoten2atlantis
{
    public class Unterrichte : List<Unterricht>
    {
        public Unterrichte()
        {
        }

        public Unterrichte(string sourceExportLessons)
        {
            using (StreamReader reader = new StreamReader(sourceExportLessons))
            {
                var überschrift = reader.ReadLine();
                int i = 1;

                while (true)
                {
                    i++;
                    Unterricht unterricht = new Unterricht();

                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            var x = line.Split('\t');

                            unterricht = new Unterricht();
                            unterricht.MarksPerLessonZeile = i;
                            unterricht.LessonId = Convert.ToInt32(x[0]);
                            unterricht.LessonNumber = Convert.ToInt32(x[1]) / 100;
                            unterricht.Fach = x[2];
                            unterricht.Lehrkraft = x[3];
                            unterricht.Klassen = x[4];
                            unterricht.Gruppe = x[5];
                            unterricht.Periode = Convert.ToInt32(x[6]);
                            unterricht.Startdate = DateTime.ParseExact(x[7], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            unterricht.Enddate = DateTime.ParseExact(x[8], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            this.Add(unterricht);
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

                Global.WriteLine(("Alle Unterrichte ").PadRight(Global.PadRight - 2, '.') + this.Count.ToString().PadLeft(6));
            }
        }

        public Unterrichte(IEnumerable<Unterricht> collection, string interessierendeKlasse) : base(collection)
        {
            foreach (var item in collection)
            {

            }
        }

        internal object GetKlassenunterrichte(string interessierendeKlasse)
        {
            throw new NotImplementedException();
        }
    }
}