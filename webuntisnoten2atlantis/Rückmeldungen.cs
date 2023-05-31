// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntisnoten2atlantis
{
    public class Rückmeldungen :List<Rückmeldung>
    {
        public void AddRückmeldung(Rückmeldung rückmeldung)
        {
            if (rückmeldung.Meldung!= "")
            {
                if (!(from t in this
                      where t.Lehrkraft == rückmeldung.Lehrkraft
                      where t.Meldung == rückmeldung.Meldung
                      where t.Fach == rückmeldung.Fach
                      select t).Any())
                {
                    this.Add(rückmeldung);
                }
            }
        }
    }
}