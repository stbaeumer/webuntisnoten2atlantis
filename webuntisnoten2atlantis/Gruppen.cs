// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.IO;

namespace webuntisnoten2atlantis
{
    public class Gruppen: List<Gruppe>
    {
        public Gruppen(string sourceStudentgroupStudents)
        {
            using (StreamReader reader = new StreamReader(sourceStudentgroupStudents))
            {
                var überschrift = reader.ReadLine();
                int i = 1;

                while (true)
                {
                    i++;
                    var gruppe = new Gruppe();

                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            var x = line.Split('\t');

                            gruppe = new Gruppe();
                            gruppe.MarksPerLessonZeile = i;
                            gruppe.StudentId = Convert.ToInt32(x[0]);
                            gruppe.Gruppenname = x[3];
                            gruppe.Fach = x[4];
                            try
                            {
                                gruppe.Startdate = DateTime.ParseExact(x[5], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            }
                            catch (Exception)
                            {
                            }

                            try
                            {
                                gruppe.Enddate = DateTime.ParseExact(x[6], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            }
                            catch (Exception)
                            {
                            }
                            
                            this.Add(gruppe);
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

                Global.WriteLine(("Alle Gruppen ").PadRight(Global.PadRight - 2, '.') + this.Count.ToString().PadLeft(6));
            }
        }
    }
}