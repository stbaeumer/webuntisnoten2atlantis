// Published under the terms of GPLv3 Stefan Bäumer 2023.

using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace webuntisnoten2atlantis
{
    internal class Termine : List<Termin>
    {
        public Termine()
        {
        }

        internal void ToOutlook(ExchangeService service, string mail)
        {
            Console.WriteLine("Konferenzen werden zu einer gemeinsamen Teams-Besprechung gruppiert werden.");

            List<string> konferenzGruppen = new List<string>();

            konferenzGruppen.Add("BW,HBW");
            konferenzGruppen.Add("BS,HBG,FS");
            konferenzGruppen.Add("BT,HBT,FM");
            konferenzGruppen.Add("G");

            Console.WriteLine("Es werden Gruppen gebildet.");

            Termin gruppenkonferenz = new Termin();

            foreach (var kg in konferenzGruppen)
            {
                Termin konf = new Termin();
                konf.Lehrers = new Lehrers();
                konf.Uhrzeit = DateTime.Now.AddYears(1);
                konf.Klasse = kg;

                foreach (var kgm in kg.Split(','))
                {
                    foreach (var k in this)
                    {
                        if (k.Klasse.StartsWith(kgm))
                        {
                            konf.Bildungsgang = konf.Bildungsgang == null ? k.Bildungsgang + "," : !konf.Bildungsgang.Contains(k.Bildungsgang) ? konf.Bildungsgang += k.Bildungsgang + "," : konf.Bildungsgang;
                            konf.Uhrzeit = k.Uhrzeit < konf.Uhrzeit ? k.Uhrzeit : konf.Uhrzeit;

                            foreach (var l in k.Lehrers)
                            {
                                if (!(from le in konf.Lehrers where le.Kuerzel == l.Kuerzel select l).Any())
                                {
                                    konf.Lehrers.Add(l);
                                }
                            }
                        }
                    }
                }
                konf.ToOutlook(service, mail);
            }
        } 
    
        internal void Lehrers(Lehrers lehrers, Leistungen alleWebuntisLeistungen)
        {
            foreach (var konferenz in this)
            {
                var le = (from l in lehrers where System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().Split('\\')[1] == l.Kuerzel select l).FirstOrDefault();
                konferenz.Lehrers = alleWebuntisLeistungen.LehrerDieserKlasse(konferenz, lehrers);                
            }
        }

        internal void TerminkollisionenFinden()
        {
            foreach (var konferenz in this)
            {
                foreach (var le in konferenz.Lehrers)
                {
                    var x = (from k in this
                             where k.Klasse != konferenz.Klasse  // Wenn eine andere Konferenz, ...
                             where k.Lehrers.Contains(le)        // ... in der dieser Lehrer Mitglied ist ...
                             where (konferenz.Uhrzeit <= k.Uhrzeit && k.Uhrzeit < konferenz.Uhrzeit.AddMinutes(10)) // ... zeitgleich oder währenddessen beginnt ...
                             select k).ToList();

                    if (x != null)
                    {
                        foreach (var konf in x)
                        {
                            Global.PrintMessage(Global.Output.Count, "Terminkollision: "  + le.Kuerzel + ": " + konferenz.Klasse + "#" + konf.Klasse);
                        }
                    }
                }
            }
        }
    }
}