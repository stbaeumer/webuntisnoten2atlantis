// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntisnoten2atlantis
{
    public class Zuordnungen : List<Zuordnung>
    {
        public List<string> AlleZulässigenAtlantisZielFächerAuflisten(Leistung dieseLeistung, Leistungen atlantisleistungen, string aktSj, string interessierendeKlasse)
        {
            List<string> alle = new List<string>();

            var i = 1;

            foreach (var quellklasse in (from z in atlantisleistungen where interessierendeKlasse == z.Klasse select z.Klasse).Distinct().ToList())
            {
                foreach (var atlantisfach in (from a in atlantisleistungen
                                              where a.Klasse == quellklasse
                                              where a.Schuljahr == aktSj
                                              select a.Fach).Distinct().ToList())
                {

                    var zielLeistung = (from t in this
                                        where t.Quellklasse == dieseLeistung.Klasse
                                        where t.Quellfach == dieseLeistung.Fach
                                        where t.Zielfach == atlantisfach
                                        select t).FirstOrDefault();

                    Console.WriteLine(i.ToString().PadLeft(3) + ". " + (quellklasse + "|" + atlantisfach).PadRight(13) + (zielLeistung != null ? " <<<<= " + dieseLeistung.Fach : "")); i++;
                    alle.Add(atlantisfach);
                }
            }
            return alle;
        }
    }
}