﻿// Published under the terms of GPLv3 Stefan Bäumer 2023.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace webuntisnoten2atlantis
{
    public class Leistungen : List<Leistung>
    {
        public Leistungen(string targetMarksPerLesson, Lehrers lehrers, string klasse)
        {
            var leistungen = new Leistungen();
            var leistungenOhneWidersprüchlicheNoten = new Leistungen();

            Global.VerschiedeneKlassenAusMarkPerLesson = new List<string>();
            
            using (StreamReader reader = new StreamReader(targetMarksPerLesson))
            {
                string überschrift = reader.ReadLine();

                //Console.Write(" Leistungsdaten aus Webuntis ".PadRight(71, '.'));

                int i = 1;

                Leistung leistung = new Leistung();

                while (true)
                {
                    string line = reader.ReadLine();
                    try
                    {                        
                        if (line != null)
                        {
                            var x = line.Split('\t');

                            // Es wird nur die interessierende Klasse ausgewertet

                            if (klasse == "" || x[2] == klasse)
                            {
                                i++;

                                if (x.Length == 10)
                                {
                                    leistung = new Leistung();
                                    leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                    leistung.Name = x[1];
                                    leistung.Klasse = x[2];
                                    leistung.Fach = x[3];
                                    leistung.Gesamtpunkte = x[9].Split('.')[0] == "" ? null : x[9].Split('.')[0];
                                    leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);
                                    leistung.Tendenz = Gesamtpunkte2Tendenz(leistung.Gesamtpunkte);
                                    leistung.Bemerkung = x[6];
                                    leistung.Lehrkraft = x[7];
                                    leistung.LehrkraftAtlantisId = (from l in lehrers where l.Kuerzel == leistung.Lehrkraft select l.AtlantisId).FirstOrDefault();
                                    leistung.SchlüsselExtern = Convert.ToInt32(x[8]);

                                    if (leistungen.LeistungHinzufügen(leistung))
                                    {
                                        leistungen.Add(leistung);
                                    }
                                }

                                // Wenn in den Bemerkungen eine zusätzlicher Umbruch eingebaut wurde:

                                if (x.Length == 7)
                                {
                                    leistung = new Leistung();
                                    leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                    leistung.Name = x[1];
                                    leistung.Klasse = x[2];
                                    leistung.Fach = x[3];
                                    leistung.Bemerkung = x[6];
                                    //Console.WriteLine("[!] Achtung: In den Zeilen " + (i - 1) + "-" + i + " hat vermutlich die Lehrkraft eine Bemerkung mit einem Zeilen-");
                                    //Console.Write("      umbruch eingebaut. Es wird nun versucht trotzdem korrekt zu importieren ... ");
                                }

                                if (x.Length == 4)
                                {
                                    leistung.Lehrkraft = x[1];
                                    leistung.SchlüsselExtern = Convert.ToInt32(x[2]);
                                    leistung.Gesamtpunkte = x[3].Split('.')[0];
                                    leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);
                                    if (klasse != "")
                                    {
                                        Global.AufConsoleSchreiben("[!] Achtung: Zeilen " + ((i - 1) + " - " + i).PadRight(7) + ": Lehrer hat Bemerkung mit einem Zeilenumbruch eingebaut. ... Korrigiert.");
                                    }

                                    if (leistung.Klasse == klasse && leistungen.LeistungHinzufügen(leistung))
                                    {
                                        leistungen.Add(leistung);
                                    }
                                }

                                if (x.Length < 4)
                                {
                                    Global.AufConsoleSchreiben("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
                                    Console.ReadKey();
                                    //OpenFiles(new List<string>() { targetMarksPerLesson });
                                    throw new Exception("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
                                }
                            }
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
                                
                var leistungenVonNeuNachAlt = leistungen.OrderByDescending(x => x.Datum).ToList();
                

                // Für jede Leistung aus Webuntis wird geprüft, ...

                foreach (var item in (from l in leistungenVonNeuNachAlt select l).ToList())
                {
                    // ... ob es bereits eine Leistung desselben Lehrers neueren Datums im selben Fach gibt, ...

                    if ((from l in leistungen 
                         where l.Datum > item.Datum 
                         where l.Lehrkraft == item.Lehrkraft 
                         where Regex.Match(l.Fach, @"^[^0-9]*").Value == Regex.Match(item.Fach, @"^[^0-9]*").Value                         
                         where l.SchlüsselExtern == item.SchlüsselExtern select l).Any())
                    {
                        // ... wenn ja, dann wird die ältere Leistung verworfen.

                        if (klasse!="")
                        {
                            Global.AufConsoleSchreiben(item.Name + "(" + item.Klasse + "|" + Regex.Match(item.Fach, @"^[^0-9]*").Value + "|" + item.Lehrkraft + ") hat widersprechende Noten bekommen. Die ältere v. " + item.Datum.ToShortDateString() + "(" + item.Gesamtnote + item.Tendenz + ") wird verworfen.");
                        }                        
                    }
                    else
                    {
                        leistungenOhneWidersprüchlicheNoten.Add(item);
                    }
                }

                if (klasse != "")
                {
                    Global.AufConsoleSchreiben(("Leistungsdaten für Klasse " + klasse + " aus Webuntis ").PadRight(Global.PadRight, '.') + leistungenOhneWidersprüchlicheNoten.Count.ToString().PadLeft(4));
                }                
            }
            this.AddRange((from l in leistungenOhneWidersprüchlicheNoten select l).OrderBy(x => x.Anlage).ThenBy(x => x.Klasse));
        }

        public string GetMöglicheKlassen()
        {
            var möglicheKlassen = (from k in this.OrderBy(xx => xx.Klasse) where k.Klasse != null select k.Klasse).Distinct().ToList();

            var möglicheKlassenString = "\nMögliche Klassen aus der Webuntis-Datei:\n ";
            int i = 0;

            foreach (var ik in möglicheKlassen)
            {
                if (ik.Length > 3)
                {
                    if (i % 17 == 0)
                    {
                        möglicheKlassenString += "\n";
                    }
                    möglicheKlassenString += ik + " ";
                }
                i++;
            }
            return möglicheKlassenString.TrimEnd(' ');
        }

        private bool LeistungHinzufügen(Leistung leistung)
        {
            // Wenn ein Webuntisdatensatz keine Gesamtnote enthält, wird er nicht angelegt

            if (leistung.Gesamtnote == null)
            {
                return false;
            }

            // Wenn es noch keine WebuntisLeistung zu SchülerID und Fach und selber Gesamtnote gibt, wird sie hinzugefügt.
            // Bei widersprechenden Gesamtnoten wird später gelöscht.

            foreach (var wl in this)
            {
                if (leistung.SchlüsselExtern == wl.SchlüsselExtern)
                {
                    if (leistung.Fach == wl.Fach)
                    {
                        if (leistung.Gesamtnote == wl.Gesamtnote)
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        internal Leistungen WidersprechendeGesamtnotenKorrigieren(Leistungen atlantisleistungen)
        {
            var leistungen = new Leistungen();

            var liste = new List<string>(); // enthält Klasse|Fachstamm|Lehrkraft

            var entfernteLeistungen = ""; // Klasse|Fach|Lehrkraft

            foreach (var wLeistung in this)
            {
                bool entferneLeistung = true;

                // Wenn ein Lehrer sich selbst in der Gesamtnote im selben Fach widerspricht

                if ((from t in this where t.Fach == wLeistung.Fach where t.SchlüsselExtern == wLeistung.SchlüsselExtern select t).Distinct().ToList().Count > 1)
                {
                    Global.AufConsoleSchreiben("Achtung: " + wLeistung.Lehrkraft + " hat dem Schüler " + wLeistung.SchlüsselExtern + " in " + wLeistung.Fach + " unterschiedliche Noten gegeben!");
                    Global.AufConsoleSchreiben("Nur die Leistung mit dem jüngsten Datum bleibt erhalten.");

                    var xx = (from t in this.OrderByDescending(x => x.Datum) where t.Fach == wLeistung.Fach where t.SchlüsselExtern == wLeistung.SchlüsselExtern select t).Distinct().ToList();
                    Console.ReadKey();
                }


                if (esGibtDoppelteFächerZuDieser(wLeistung)) // Z.B. GPF1 und GPF2
                {
                    if (esGibtWidersprechendeNoten(wLeistung)) // Z.B. Note 7.0 in GPF1 und Note 10.0 in GPF
                    {
                        // Es wird nur dann nachgefragt, wenn für diese Klasse und dieses Fach noch kein Eintrag existiert

                        if (!(from l in liste where l.StartsWith(wLeistung.Klasse + "|" + Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value) select l).Any())
                        {
                            liste.Add(NachfragenWelchesFachZuWählenIst(wLeistung));
                        }

                        // Nur wenn diese wLeistung dem Muster der liste entspricht, wird sie hinzugefügt

                        if (liste.Contains(wLeistung.Klasse + "|" + Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value + "|" + wLeistung.Lehrkraft))
                        {
                            leistungen.Add(wLeistung); entferneLeistung = false; 
                        }
                    }
                    else
                    {
                        leistungen.Add(wLeistung); entferneLeistung = false;  // ohne Notenwiderspruch verbleibt die Leistung in den Webuntisleistungen
                    }
                }
                else
                {   
                    leistungen.Add(wLeistung); entferneLeistung = false; // ohne Dopplungen verbleibt die Leistung in den Webuntisleistungen
                }

                // Wenn die Leistung nicht hinzugefügt wurde, wird sie protokolliert

                if (entferneLeistung)
                {
                    if (!entfernteLeistungen.Contains(wLeistung.Klasse + "|" + wLeistung.Fach + "|" + wLeistung.Lehrkraft)) 
                    {
                        entfernteLeistungen += wLeistung.Klasse + "|" + wLeistung.Fach + "|" + wLeistung.Lehrkraft + ",";
                    }
                }
            }

            if (this.Count > leistungen.Count)
            {
                Global.AufConsoleSchreiben("Aufgrund von widersprechenden Noten sind " + (this.Count - leistungen.Count) + " Leistungen ("+ entfernteLeistungen.TrimEnd(',') +") entfernt worden.");
            }

            return leistungen;
        }

        private string NachfragenWelchesFachZuWählenIst(Leistung wLeistung)
        {
            var ähnlicheFächer = (from w in this
                                  where w.Klasse == wLeistung.Klasse
                                  where w.SchlüsselExtern == wLeistung.SchlüsselExtern
                                  where Regex.Match(w.Fach, @"^[^0-9]*").Value == Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value
                                  select w.Fach).Distinct().ToList();

            var ähnlicheFächerMitUnterschiedlichenNoten = (from w in this
                                                           where w.Klasse == wLeistung.Klasse
                                                           where w.SchlüsselExtern == wLeistung.SchlüsselExtern
                                                           where ähnlicheFächer.Contains(w.Fach)
                                                           select w).ToList();

            Global.AufConsoleSchreiben("Widersprechende Gesamtnoten im Fach " + Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value + "* beim Schüler " + wLeistung.SchlüsselExtern + ":");

            for (int i = 0; i < ähnlicheFächerMitUnterschiedlichenNoten.Count; i++)
            {
                Global.AufConsoleSchreiben((i + 1) + ". " + ähnlicheFächerMitUnterschiedlichenNoten[i].SchlüsselExtern.ToString().PadRight(7) + ähnlicheFächerMitUnterschiedlichenNoten[i].Name.ToString().PadRight(25) + ähnlicheFächerMitUnterschiedlichenNoten[i].Lehrkraft.PadRight(4) + ähnlicheFächerMitUnterschiedlichenNoten[i].Fach.PadRight(6) + "Note: " + ähnlicheFächerMitUnterschiedlichenNoten[i].Gesamtnote + ähnlicheFächerMitUnterschiedlichenNoten[i].Tendenz + " (Punkte: " + ähnlicheFächerMitUnterschiedlichenNoten[i].Gesamtpunkte +")");
            }

            bool wiederholen = true;

            int index = 0;

            do
            {
                Console.WriteLine("Bitte eine Zahl zwischen 1 und " + ähnlicheFächerMitUnterschiedlichenNoten.Count + " eingeben, um generell zwischen den genannten LuL in der " + wLeistung.Klasse + " im Fach " + Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value + "* zu wählen:");

                ConsoleKeyInfo eingabe;

                eingabe = Console.ReadKey();
                 
                if (char.IsDigit(eingabe.KeyChar))
                {
                    index = int.Parse(eingabe.KeyChar.ToString()); // use Parse if it's a Digit

                    if (index > 0 && index <= ähnlicheFächerMitUnterschiedlichenNoten.Count)
                    {
                        wiederholen = false;
                    }
                }
            } while (wiederholen);

            Global.AufConsoleSchreiben("Sie haben " + index + " gewählt.");

            return wLeistung.Klasse + "|" + Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value + "|" + ähnlicheFächerMitUnterschiedlichenNoten[index-1].Lehrkraft;
        }

        private bool esGibtWidersprechendeNoten(Leistung wLeistung)
        {
            var ähnlicheFächer = (from w in this
                                  where w.Klasse == wLeistung.Klasse
                                  where w.SchlüsselExtern == wLeistung.SchlüsselExtern
                                  where Regex.Match(w.Fach, @"^[^0-9]*").Value == Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value
                                  select w.Fach).Distinct().ToList();

            var ähnlicheFächerMitUnterschiedlichenNoten = (from w in this
                                                           where w.Klasse == wLeistung.Klasse
                                                           where w.SchlüsselExtern == wLeistung.SchlüsselExtern
                                                           where ähnlicheFächer.Contains(wLeistung.Fach)
                                                           select w.Gesamtpunkte).Distinct().ToList();

            if (ähnlicheFächerMitUnterschiedlichenNoten.Count > 1)
            {
                return true;
            }
            return false;
        }

        private bool esGibtDoppelteFächerZuDieser(Leistung wLeistung)
        {
            var ähnlicheFächer = (from w in this
                                  where w.Klasse == wLeistung.Klasse
                                  where w.SchlüsselExtern == wLeistung.SchlüsselExtern
                                  where Regex.Match(w.Fach, @"^[^0-9]*").Value == Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value
                                  select w.Fach).Distinct().ToList();

            if (ähnlicheFächer.Count > 1)
            {
                return true;
            }
            return false;
        }

        private List<string> GetDoppelteFächerOhneZählerMitWidersprechendenNoten(Leistung wLeistung)
        {
            var fächerOhneZähler = (from w in this where w.Klasse == wLeistung.Klasse select Regex.Match(w.Fach, @"^[^0-9]*").Value).Distinct().ToList();

            var fächer = (from w in this where w.Klasse == wLeistung.Klasse select w.Fach).Distinct().ToList();

            var doppelteFächer = new List<string>();

            foreach (var fach in fächerOhneZähler)
            {
                var anz = 0;

                foreach (var item in fächer)
                {
                    if (Regex.Match(item, @"^[^0-9]*").Value == fach)
                    {
                        anz++;
                    }
                }
                if (anz>1)
                {
                    if (!doppelteFächer.Contains(fach))
                    {
                        doppelteFächer.Add(fach);
                    }                    
                }
            }
            return doppelteFächer;
        }

        internal void ErzeugeSerienbriefquelleFehlendeTeilleistungen(Leistungen webuntisLeistungen)
        {
            List<Leistung> fehlendeTeilleistungen = new List<Leistung>();

            foreach (var atlantisLeistung in (from t in this where t.Fach != "" where !(t.Zeugnistext.StartsWith("Abschluss")) select t).ToList())
            {
                // Wenn es für eine Atlantisleistung keine Entsprechung in Untis gibt, wird der Datensatz hinzugefügt.

                if (!(from w in webuntisLeistungen where w.Klasse == atlantisLeistung.Klasse where w.Fach == atlantisLeistung.Fach select w).Any())
                {
                    fehlendeTeilleistungen.Add(atlantisLeistung);
                }
            }
        }

        internal void Gym12NotenInDasGostNotenblattKopieren(List<string> interessierendeKlassen, List<string> aktSj)
        {
            foreach (var item in this)
            {
                // Wenn es sich um die Leistung einer Gym 12 oder 13 handelt ...

                if (item.Klasse.StartsWith("G") && (item.Klasse.Contains((Convert.ToInt32(aktSj[1]) - 2).ToString()) || item.Klasse.Contains((Convert.ToInt32(aktSj[1]) - 3).ToString())))
                {
                    // ... wird die Leistung in das Notenblatt der GO geschrieben.

                    throw new NotImplementedException();

                }
            }
        }

        internal void ErzeugeSerienbriefquelleNichtversetzer()
        {
            var l = new Leistungen();

            foreach (var leistung in this)
            {
                if (leistung.SchlüsselExtern == 153203)
                {
                    l.Add(leistung);
                }
            }
        }

        internal void AtlantisLeistungenZuordnenUndQueryBauen(Leistungen atlantisLeistungen, string aktuellesSchuljahr, string interessierendeKlasse)
        {
            var kombinationAusKlasseUndFachWirdNurEinmalAufgerufen = new List<string>();

            Zuordnungen gespeicherteZuordnungen = new Zuordnungen();

            // Es wird versucht jedes Webuntis-Fach einem Atlantis-Fach zuzuordnen

            foreach (var webuntisLeistung in this)
            {
                // Zielleistungen dieses Schülers im aktuellen Abschnitt werden ermittelt.

                var atlantisZielLeistungen = (from a in atlantisLeistungen
                                              where (a.Konferenzdatum >= DateTime.Now || a.Konferenzdatum.Year == 1) // eine Leistung des aktuellen Abschnitts
                                              where a.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
                                              where a.Schuljahr == aktuellesSchuljahr
                                              where a.HzJz == Global.HzJz
                                              orderby a.LeistungId descending // Die höchste ID ist die ID des aktuellen Abschnitts
                                              select a).ToList();

                if (webuntisLeistung.Fach == "DV")
                {
                    string aaa = "";
                }
                // Wenn es den Schüler überhaupt nicht in Atlantis gibt, dann wird die Verarbeitung abgebrochen

                if (atlantisZielLeistungen.Count > 0)
                {
                    // Im Idealfall stimmen die Fachnamen überein.

                    if (webuntisLeistung.Query == null)
                    {
                        var aL = (from a in atlantisZielLeistungen
                                  where a.Fach == webuntisLeistung.Fach // Prüfung auf Fachnamensgleichheit
                                  select a).ToList();

                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, "");
                    }

                    // Suffix entfernen

                    if (webuntisLeistung.Query == null)
                    {
                        if (char.IsDigit(webuntisLeistung.Fach.Last()))
                        {
                            var aL = (from a in atlantisZielLeistungen
                                      where a.Fach == Regex.Match(webuntisLeistung.Fach, @"^[^0-9]*").Value
                                      select a).ToList();

                            webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->" + Regex.Match(webuntisLeistung.Fach, @"^[^0-9]*").Value + "|");
                        }
                    }

                    // Religion wird immer zu REL

                    if (webuntisLeistung.Query == null)
                    {
                        if (webuntisLeistung.Fach == "KR" || webuntisLeistung.Fach == "ER" || webuntisLeistung.Fach.StartsWith("KR ") || webuntisLeistung.Fach.StartsWith("ER "))
                        {
                            var aL = (from a in atlantisZielLeistungen
                                      where a.Fach == "REL"
                                      select a).ToList();

                            webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->REL|");
                        }
                    }

                    // EK -> E

                    if (webuntisLeistung.Query == null)
                    {
                        if (webuntisLeistung.Fach.StartsWith("EK") && webuntisLeistung.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "E")
                        {
                            var aL = (from a in atlantisZielLeistungen
                                      where a.Fach == "E"
                                      select a).ToList();

                            webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->E|");
                        }
                        if (webuntisLeistung.Fach.StartsWith("EB1"))
                        {
                            var aL = (from a in atlantisZielLeistungen
                                      where a.Fach.StartsWith("E")
                                      where a.Fach.Contains("B1")
                                      select a).ToList();

                            webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->EB1|");
                        }
                    }

                    // N -> NKA1

                    if (webuntisLeistung.Query == null)
                    {
                        if (webuntisLeistung.Fach.StartsWith("N"))
                        {
                            if (webuntisLeistung.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "NK")
                            {
                                var aL = (from a in atlantisZielLeistungen
                                          where a.Fach == "N"
                                          select a).ToList();

                                webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->N|");
                            }
                        }
                    }

                    // Kurse mit Leerzeichen

                    if (webuntisLeistung.Query == null)
                    {
                        var aL = (from a in atlantisZielLeistungen
                                  where (webuntisLeistung.Fach.Contains(" ") && a.Fach.Contains(" ")) // Wenn beide Leerzeichen enthalten
                                  where (webuntisLeistung.Fach.Split(' ')[0] == a.Fach.Split(' ')[0]) // Wenn der erste Teil identisch ist
                                  where (webuntisLeistung.Fach.Substring(webuntisLeistung.Fach.LastIndexOf(' ') + 1).Substring(0, 1) == a.Fach.Substring(a.Fach.LastIndexOf(' ') + 1).Substring(0, 1)) // Wenn der zweite Teil identisch anfängt
                                  select a).ToList();

                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->" + (aL.Count > 0 ? aL[0].Fach : "") + "|");

                        // Kurse mit Leerzeichen

                        // Wenn in Atlantis die Anzahl der Zielfächer größer als 1 ist, dann wird eine weitere Leistung hinzugefügt 

                        if (aL.Count > 1)
                        {
                            Leistung neueLeistung = new Leistung();

                            neueLeistung.Abschlussklasse = webuntisLeistung.Abschlussklasse;
                            neueLeistung.Anlage = webuntisLeistung.Anlage;
                            neueLeistung.Bemerkung = webuntisLeistung.Bemerkung;
                            neueLeistung.Beschreibung = webuntisLeistung.Fach + "->" + aL[1] + ",";
                            neueLeistung.Datum = webuntisLeistung.Datum;
                            neueLeistung.EinheitNP = webuntisLeistung.EinheitNP;
                            neueLeistung.Fach = aL[1].Fach;
                            neueLeistung.Zielfach = aL[1].Fach;
                            neueLeistung.GeholteNote = webuntisLeistung.GeholteNote;
                            neueLeistung.Gesamtnote = webuntisLeistung.Gesamtnote;
                            neueLeistung.Gesamtpunkte = webuntisLeistung.Gesamtpunkte;
                            neueLeistung.Gliederung = webuntisLeistung.Gliederung;
                            neueLeistung.HatBemerkung = webuntisLeistung.HatBemerkung;
                            neueLeistung.HzJz = webuntisLeistung.HzJz;
                            neueLeistung.Jahrgang = webuntisLeistung.Jahrgang;
                            neueLeistung.Klasse = webuntisLeistung.Klasse;
                            neueLeistung.Konferenzdatum = webuntisLeistung.Konferenzdatum;
                            neueLeistung.Lehrkraft = webuntisLeistung.Lehrkraft;
                            neueLeistung.LeistungId = webuntisLeistung.LeistungId;
                            neueLeistung.Name = webuntisLeistung.Name;
                            neueLeistung.ReligionAbgewählt = webuntisLeistung.ReligionAbgewählt;
                            neueLeistung.SchlüsselExtern = webuntisLeistung.SchlüsselExtern;
                            neueLeistung.SchuelerAktivInDieserKlasse = webuntisLeistung.SchuelerAktivInDieserKlasse;
                            neueLeistung.Schuljahr = webuntisLeistung.Schuljahr;
                            neueLeistung.Tendenz = webuntisLeistung.Tendenz;
                            neueLeistung.Zeugnisart = webuntisLeistung.Zeugnisart;
                            neueLeistung.Zeugnistext = webuntisLeistung.Zeugnistext;
                            this.Add(neueLeistung);
                        }
                    }

                    // Wenn keine automatische Zuordnung vorgenommen wurde:

                    if (webuntisLeistung.Query == null && !kombinationAusKlasseUndFachWirdNurEinmalAufgerufen.Contains(webuntisLeistung.Klasse + webuntisLeistung.Fach))
                    {
                        Global.AufConsoleSchreiben("Das Webuntis-Fach **" + webuntisLeistung.Fach + "** kann keinem Atlantis-Fach zugeordnet werden. Bitte manuell zuordnen:");

                        var alleZulässigenAtlantisZielfächer = gespeicherteZuordnungen.AlleZulässigenAtlantisZielFächerAuflisten(webuntisLeistung, atlantisLeistungen, aktuellesSchuljahr);

                        Global.AufConsoleSchreiben(" Wollen Sie die Zuordnung für **" + webuntisLeistung.Fach + "** manuell vornehmen?");

                        bool wiederholen = true;

                        int index = 0;

                        ConsoleKeyInfo eingabe;

                        do
                        {
                            Global.AufConsoleSchreiben(" Bitte Ziffer 1 bis " + alleZulässigenAtlantisZielfächer.Count + " eingeben oder ENTER falls keine Änderung gewünscht ist: ");

                            eingabe = Console.ReadKey();

                            if (char.IsDigit(eingabe.KeyChar))
                            {
                                index = int.Parse(eingabe.KeyChar.ToString()); // use Parse if it's a Digit

                                if (index > 0 && index <= alleZulässigenAtlantisZielfächer.Count)
                                {
                                    wiederholen = false;
                                }
                            }
                        } while (wiederholen);

                        // Die Zuordnung wird für alle anderen Schüler ebenfalls vorgenommen

                        foreach (var w in (from w in this
                                           where w.Fach == webuntisLeistung.Fach
                                           where w.Klasse == interessierendeKlasse
                                           select w).ToList())
                        {
                            var aL = (from a in atlantisLeistungen
                                      where (a.Konferenzdatum >= DateTime.Now || a.Konferenzdatum.Year == 1) // eine Leistung des aktuellen Abschnitts
                                      where a.Schuljahr == aktuellesSchuljahr
                                      where a.Klasse == interessierendeKlasse
                                      where a.Fach == alleZulässigenAtlantisZielfächer[index -1]
                                      where a.SchlüsselExtern == w.SchlüsselExtern
                                      orderby a.LeistungId descending // Die höchste ID ist die ID des aktuellen Abschnitts
                                      select a).ToList();

                            if (aL.Count > 0)
                            {
                                webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->" + (aL.Count > 0 ? aL[0].Fach : "") + "|manuell zugeordnet|");
                            }
                            else
                            {
                                webuntisLeistung.Query = "/* Das Fach " + webuntisLeistung.Fach + " wurde keinem Atlantisfach zugeordnet.";
                                webuntisLeistung.Beschreibung = "";
                            }
                        }

                        kombinationAusKlasseUndFachWirdNurEinmalAufgerufen.Add(webuntisLeistung.Klasse + webuntisLeistung.Fach);
                    }
                }
            }
                                    
            Global.AufConsoleSchreiben("Schülerleistungen aus Webuntis Atlantisleistungen erfolgreich zugeordnet ".PadRight(110, '.') + this.Count.ToString().PadLeft(4));
        }
            
        internal void FehlendeSchülerInAtlantis(Leistungen atlantisLeistungen)
        {
            List<string> keinSchülerInAtlantis = new List<string>();

            foreach (var we in this)
            {
                var atlantisFächerDiesesSchuelers = (from a in atlantisLeistungen
                                                     where a.Klasse == we.Klasse
                                                     where a.SchlüsselExtern == we.SchlüsselExtern
                                                     select a).ToList();

                if (atlantisFächerDiesesSchuelers.Count == 0 && !keinSchülerInAtlantis.Contains(we.Name + "|" + we.Klasse))
                {
                    keinSchülerInAtlantis.Add(we.Name + "|" + we.Klasse);
                }
            }
            if (keinSchülerInAtlantis.Count > 0)
            {
                AusgabeSchreiben("Achtung: Es gibt Leistungsdatensätze in Webuntis, die keinem Schüler in Atlantis zugeordnet werden können:", keinSchülerInAtlantis);
            }
        }

        private string Gesamtpunkte2Tendenz(string gesamtpunkte)
        {
            string tendenz = null;

            if (gesamtpunkte == "1")
            {
                tendenz = "-";
            }
            if (gesamtpunkte == "3")
            {
                tendenz = "+";
            }
            if (gesamtpunkte == "4")
            {
                tendenz = "-";
            }
            if (gesamtpunkte == "6")
            {
                tendenz = "+";
            }
            if (gesamtpunkte == "7")
            {
                tendenz = "-";
            }
            if (gesamtpunkte == "9")
            {
                tendenz = "+";
            }
            if (gesamtpunkte == "10")
            {
                tendenz = "-";
            }
            if (gesamtpunkte == "12")
            {
                tendenz = "+";
            }
            if (gesamtpunkte == "13")
            {
                tendenz = "-";
            }
            if (gesamtpunkte == "15")
            {
                tendenz = "+";
            }

            if (tendenz == "")
            {
                return null;
            }
            return tendenz;
        }

        internal void FehlendeZeugnisbemerkungBeiStrich(Leistungen webuntisLeistungen, List<string> interessierendeKlassen)
        {
            List<string> fehlendeBemerkungen = new List<string>();

            try
            {
                foreach (var klasse in interessierendeKlassen)
                {
                    foreach (var a in (from t in this
                                       where t.Klasse == klasse
                                       where !t.Anlage.StartsWith("A")
                                       where t.SchuelerAktivInDieserKlasse
                                       where t.Fach != "REL"
                                       select t).OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name).ToList())
                    {
                        var w = (from webuntisLeistung in webuntisLeistungen
                                 where webuntisLeistung.Fach == a.Fach
                                 where webuntisLeistung.Klasse == a.Klasse
                                 where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                 where a.Gesamtnote == null
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null && w.Gesamtnote != null)
                        {
                            // Ein '-' in Religion wird nur dann gesetzt, wenn bereits andere Schüler eine Gesamtnote bekommen haben

                            if (!(w.Fach == "REL" && w.Gesamtnote == "-" && (from we in webuntisLeistungen
                                                                             where we.Klasse == w.Klasse
                                                                             where we.Fach == "REL"
                                                                             where (we.Gesamtnote != null && we.Gesamtnote != "-")
                                                                             select we).Count() == 0))
                            {
                                if (w.Gesamtnote == "-" && (a.Bemerkung == null || w.Bemerkung == null))
                                {
                                    fehlendeBemerkungen.Add(w.Name + "(" + w.Klasse + "," + w.Fach + "," + w.Lehrkraft + ")");
                                }
                            }
                        }
                    }
                }
                if (fehlendeBemerkungen.Count > 0)
                {
                    AusgabeSchreiben("Es gibt " + fehlendeBemerkungen.Count + " Schüler*innen in Vollzeitklassen, die einen '-' in einem Nicht-Reli-Fach als Noten bekommen, ohne dass eine entsprechende Zeugnisbemerkung vorliegt:", fehlendeBemerkungen);
                }
                Global.AufConsoleSchreiben(("Fehlende Zeugnisbemerkung bei Strich ").PadRight(Global.PadRight, '.') + fehlendeBemerkungen.Count.ToString().PadLeft(4));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal Leistungen NotenVergangenerAbschnitteZiehen(Leistungen webuntisLeistungen, List<string> interessierendeKlassen, List<string> aktSj)
        {
            var digits = new[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

            try
            {
                Leistungen leistungen = new Leistungen();

                List<string> leistungenDieserListe = new List<string>();

                Global.AufConsoleSchreiben("");

                Global.AufConsoleSchreiben("Folgende aktuelle Leistungen aus webuntis und Leistungen vergangener Abschnitte liegen vor:");
                Global.AufConsoleSchreiben("");
                

                var alleVerschiedenenInteressierendenSchüler = (from t in this where t.SchlüsselExtern != 0 select t.SchlüsselExtern).Distinct().ToList();

                var alleVerschiedenenKlassenInteressierenderSchüler = (from t in this select t.Klasse).Distinct().ToList();

                var alleVerschiedenenAlteKonferenzdaten = (from xs in this.OrderByDescending(xs => xs.Konferenzdatum)
                                                           where xs.GeholteNote
                                                           where (xs.Gesamtnote != null && xs.Gesamtnote!= "")
                                                           select xs.Konferenzdatum).Distinct().ToList();

                var alleLeistungsdatenDerVergangenheit = (from xs in this.OrderByDescending(xs => xs.Konferenzdatum).ThenBy(xs => xs.Fach)
                                                          where xs.GeholteNote
                                                          where (xs.Gesamtnote != null && xs.Gesamtnote != "")
                                                          select xs).ToList();

                // Kopfzeile wird geschrieben

                var xx = "SuS   |Klasse|aktueller Abschnitt                     ";
                foreach (var item in alleVerschiedenenAlteKonferenzdaten)
                {
                    xx += "| " + item.ToShortDateString() + "     ";
                }

                Global.AufConsoleSchreiben(xx);

                // Wenn bereits eine Leistung neueren Datums existiert, wird die alte nicht angelegt

                var fachExistiertSchon = new List<string>();

                // alle Schüler werden gelistet.

                foreach (var schülerId in alleVerschiedenenInteressierendenSchüler)
                {
                    var aktuelleFächer = (from s in webuntisLeistungen where s.SchlüsselExtern == schülerId select s.Fach).Distinct().ToList();
                    var alteFächer = (from s in this where s.SchlüsselExtern == schülerId where !aktuelleFächer.Contains(s.Fach) select s.Fach).Distinct().ToList();
                    var aktuelleKlasse = (from s in webuntisLeistungen where s.SchlüsselExtern == schülerId select s.Klasse).FirstOrDefault();

                    var zeile = schülerId.ToString().PadLeft(6) + "|" + aktuelleKlasse.PadRight(6) + "|";

                    foreach (var fa in aktuelleFächer)
                    {
                        var note = (from s in webuntisLeistungen.OrderByDescending(x=>x.Datum) where s.SchlüsselExtern == schülerId where s.Fach == fa select s.Gesamtnote).FirstOrDefault();

                        zeile += fa + "(" + note + "),";

                        if (!fachExistiertSchon.Contains(schülerId + fa))
                        {
                            fachExistiertSchon.Add(schülerId + fa);
                        }
                    }

                    zeile = zeile.TrimEnd(',').PadRight(54);
                                        
                    foreach (var konferenzdatum in alleVerschiedenenAlteKonferenzdaten)
                    {
                        var faa = "";
                        foreach (var fach in (from s in this where s.SchlüsselExtern == schülerId where !aktuelleFächer.Contains(s.Fach) where konferenzdatum == s.Konferenzdatum select s.Fach).Distinct().ToList())
                        {
                            if (!fachExistiertSchon.Contains(schülerId + fach))
                            {
                                fachExistiertSchon.Add(schülerId + fach);
                                
                                var l = (from s in this
                                         where s.SchlüsselExtern == schülerId
                                         where !aktuelleFächer.Contains(s.Fach)
                                         where konferenzdatum == s.Konferenzdatum
                                         where s.Fach == fach
                                         select s).ToList();
                                leistungen.Add(l[0]);
                                faa += fach + "(" + l[0].Gesamtnote + "),";
                            }                            
                        }
                        zeile = zeile + "|" + (faa.TrimEnd(',')).PadRight(20);                        
                    }
                    Global.AufConsoleSchreiben(zeile);
                }

                Global.AufConsoleSchreiben(" ");

                // Wenn es keine alten Konferenzdaten gibt, wird auch nicht nachgefragt

                if (alleVerschiedenenAlteKonferenzdaten.Count > 0)
                {
                    ConsoleKeyInfo x;

                    do
                    {
                        Global.AufConsoleSchreiben(" ");
                        Global.AufConsoleSchreiben("Wenn später doch noch Noten über Webuntis eingelesen werden, werden die geholten Noten überschrieben.");
                        Global.AufConsoleSchreiben("Wollen Sie die Noten aus den alten Abschnitten ziehen? (j/N)");
                        x = Console.ReadKey();

                        Global.AufConsoleSchreiben(x.Key.ToString());

                    } while (x.Key.ToString().ToLower() != "j" && x.Key.ToString().ToLower() != "n" && x.Key.ToString() != "Enter");

                    Leistungen zuHolendeLeistungen = new Leistungen();

                    // Für den Fall, dass die Leistungen nicht geholt werden sollen, ...

                    if (x.Key.ToString().ToLower() != "j")
                    {                           
                        foreach (var zuholendeLeistung in leistungen)
                        {
                            // ... aber zuvor schonmal geholt wurden, also schon eine Note gesetzt ist, wird alles genullt.

                            var atlantisZielLeistungen = (from a in this
                                                          where (a.Konferenzdatum >= DateTime.Now || a.Konferenzdatum.Year == 1) // eine Leistung des aktuellen Abschnitts
                                                          where a.SchlüsselExtern == zuholendeLeistung.SchlüsselExtern
                                                          where a.Schuljahr == aktSj[0] + "/" + aktSj[1]
                                                          where a.HzJz == Global.HzJz
                                                          where a.Fach == zuholendeLeistung.Fach
                                                          where a.Klasse == zuholendeLeistung.Klasse
                                                          where (a.Gesamtnote != null || a.Tendenz != null || a.Gesamtpunkte != null)
                                                          orderby a.LeistungId descending // Die höchste ID ist die ID des aktuellen Abschnitts
                                                          select a).ToList();
                            if (atlantisZielLeistungen.Count > 0)
                            {
                                zuholendeLeistung.Gesamtnote = "";
                                zuholendeLeistung.Gesamtpunkte = "";
                                zuholendeLeistung.Tendenz = "";
                                zuHolendeLeistungen.Add(zuholendeLeistung);
                            }
                        }
                        leistungen = new Leistungen();
                        leistungen.AddRange(zuHolendeLeistungen);
                    }
                }

                return leistungen;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public Leistungen(string connetionstringAtlantis, List<string> aktSj, string user, List<string> interessierendeKlassen, Leistungen alleWebuntisLeistungen)
        {
            Global.Defizitleistungen = new Leistungen();

            // Es werden alle Atlantisleistungen für alle SuS gezogen.
            // Weil es denkbar ist, dass ein Schüler eine Leistung in einer anderen Klasse erbracht hat, werden 
            // für alle SuS alle Leistungen in diesem Bildungsgang gezogen.

            var abfrage = "";

            foreach (var iK in interessierendeKlassen)
            {
                var schuelersId = (from a in alleWebuntisLeistungen where a.Klasse == iK select a.SchlüsselExtern).Distinct().ToList();

                var klassenStammOhneJahrUndOhneZähler = Regex.Match(iK, @"^[^0-9]*").Value;

                foreach (var schuelerId in schuelersId)
                {
                    abfrage += "(DBA.schueler.pu_id = " + schuelerId + @" AND  DBA.klasse.klasse like '" + klassenStammOhneJahrUndOhneZähler + @"%') OR ";
                }
            }

            try
            {
                abfrage = abfrage.Substring(0, abfrage.Length - 4);
            }
            catch (Exception)
            {
                Global.AufConsoleSchreiben("Kann es sein, dass für die Auswahl keine Datensätze vorliegen?");
            }

            var leistungenUnsortiert = new Leistungen();

            try
            {
                using (OdbcConnection connection = new OdbcConnection(connetionstringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.noten_einzel.noe_id AS LeistungId,
DBA.noten_einzel.fa_id,
DBA.noten_einzel.kurztext AS Fach,
DBA.noten_einzel.zeugnistext AS Zeugnistext,
DBA.noten_einzel.s_note AS Note,
DBA.noten_einzel.punkte AS Punkte,
DBA.noten_einzel.punkte_12_1 AS Punkte_12_1,
DBA.noten_einzel.punkte_12_2 AS Punkte_12_2,
DBA.noten_einzel.punkte_13_1 AS Punkte_13_1,
DBA.noten_einzel.punkte_13_2 AS Punkte_13_2,
DBA.noten_einzel.s_tendenz AS Tendenz,
DBA.noten_einzel.s_einheit AS Einheit,
DBA.noten_einzel.ls_id_1 AS LehrkraftAtlantisId,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.dat_geburt,
DBA.schueler.pu_id AS SchlüsselExtern,
DBA.schue_sj.s_religions_unterricht AS Religion,
DBA.schue_sj.dat_austritt AS ausgetreten,
DBA.schue_sj.vorgang_akt_satz_jn AS SchuelerAktivInDieserKlasse,
DBA.schue_sj.vorgang_schuljahr AS Schuljahr,
(substr(schue_sj.s_berufs_nr,4,5)) AS Fachklasse,
DBA.klasse.s_klasse_art AS Anlage,
DBA.klasse.jahrgang AS Jahrgang,
DBA.schue_sj.s_gliederungsplan_kl AS Gliederung,
DBA.noten_kopf.s_typ_nok AS HzJz,
DBA.noten_kopf.nok_id AS NOK_ID,
s_art_fach,
DBA.noten_kopf.s_art_nok AS Zeugnisart,
DBA.noten_kopf.bemerkung_block_1 AS Bemerkung1,
DBA.noten_kopf.bemerkung_block_2 AS Bemerkung2,
DBA.noten_kopf.bemerkung_block_3 AS Bemerkung3,
DBA.noten_kopf.dat_notenkonferenz AS Konferenzdatum,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE schue_sj.s_typ_vorgang = 'A' AND (s_typ_nok = 'JZ' OR s_typ_nok = 'HZ' OR s_typ_nok = 'GO') AND
(  
  " + abfrage + @"
)
ORDER BY DBA.klasse.s_klasse_art DESC, DBA.noten_kopf.dat_notenkonferenz DESC, DBA.klasse.klasse ASC, DBA.noten_kopf.nok_id, DBA.noten_einzel.position_1; ", connection);


                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    string bereich = "";

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        if (theRow["s_art_fach"].ToString() == "U")
                        {
                            bereich = theRow["Zeugnistext"].ToString();
                        }
                        else
                        {
                            DateTime austrittsdatum = theRow["ausgetreten"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["ausgetreten"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                            Leistung leistung = new Leistung();

                            try
                            {
                                // Wenn der Schüler nicht in diesem Schuljahr ausgetreten ist ...

                                if (!(austrittsdatum > new DateTime(DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1, 8, 1) && austrittsdatum < DateTime.Now))
                                {
                                    leistung.LeistungId = Convert.ToInt32(theRow["LeistungId"]);
                                    leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"]);

                                    leistung.ReligionAbgewählt = theRow["Religion"].ToString() == "N";
                                    leistung.Schuljahr = theRow["Schuljahr"].ToString();
                                    leistung.Gliederung = theRow["Gliederung"].ToString();
                                    leistung.HatBemerkung = (theRow["Bemerkung1"].ToString() + theRow["Bemerkung2"].ToString() + theRow["Bemerkung3"].ToString()).Contains("Fehlzeiten") ? true : false;
                                    leistung.Jahrgang = Convert.ToInt32(theRow["Jahrgang"].ToString().Substring(3, 1));
                                    leistung.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                                    leistung.Nachname = theRow["Nachname"].ToString();
                                    leistung.Vorname = theRow["Vorname"].ToString();

                                    if ((theRow["LehrkraftAtlantisId"]).ToString() != "")
                                    {
                                        leistung.LehrkraftAtlantisId = Convert.ToInt32(theRow["LehrkraftAtlantisId"]);
                                    }
                                    leistung.Bereich = bereich;
                                    leistung.Geburtsdatum = theRow["dat_geburt"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["dat_geburt"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                    leistung.Volljährig = leistung.Geburtsdatum.AddYears(18) > DateTime.Now ? false : true;
                                    leistung.Klasse = theRow["Klasse"].ToString();                                    
                                    leistung.Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();                                    
                                    leistung.Gesamtnote = theRow["Note"].ToString() == "" ? null : theRow["Note"].ToString() == "Attest" ? "A" : theRow["Note"].ToString();
                                    leistung.Gesamtpunkte_12_1 = theRow["Punkte_12_1"].ToString() == "" ? null : (theRow["Punkte_12_1"].ToString()).Split(',')[0];
                                    leistung.Gesamtpunkte_12_2 = theRow["Punkte_12_2"].ToString() == "" ? null : (theRow["Punkte_12_2"].ToString()).Split(',')[0];
                                    leistung.Gesamtpunkte_13_1 = theRow["Punkte_13_1"].ToString() == "" ? null : (theRow["Punkte_13_1"].ToString()).Split(',')[0];
                                    leistung.Gesamtpunkte_13_2 = theRow["Punkte_13_2"].ToString() == "" ? null : (theRow["Punkte_13_2"].ToString()).Split(',')[0];
                                    leistung.Gesamtpunkte = theRow["Punkte"].ToString() == "" ? null : (theRow["Punkte"].ToString()).Split(',')[0];

                                    // In der Gym ab Jahrgang 2 wird die Gesamtnote an die Kursabschnittnote angepasst.
                                    // Das kann notwendig sein, wenn manuell eine Kursabschnittnote gesetzt wurde.

                                    //if (leistung.EinheitNP == "P" && leistung.Jahrgang == 2 && leistung.HzJz == "GO" && leistung.Datum.Month == 1)
                                    //{
                                    //    if (leistung.Gesamtpunkte != leistung.Gesamtpunkte_12_1)
                                    //    {
                                    //        leistung.Gesamtpunkte = leistung.Gesamtpunkte_12_1;                                                
                                    //    }
                                    //}
                                    //if (leistung.EinheitNP == "P" && leistung.Jahrgang == 2 && leistung.HzJz == "GO" && leistung.Datum.Month > 2)
                                    //{
                                    //    if (leistung.Gesamtpunkte != leistung.Gesamtpunkte_12_2)
                                    //    {
                                    //        leistung.Gesamtpunkte = leistung.Gesamtpunkte_12_2;
                                    //    }
                                    //}
                                    //if (leistung.EinheitNP == "P" && leistung.Jahrgang == 3 && leistung.HzJz == "GO" && leistung.Datum.Month == 1)
                                    //{
                                    //    if (leistung.Gesamtpunkte != leistung.Gesamtpunkte_13_1)
                                    //    {
                                    //        leistung.Gesamtpunkte = leistung.Gesamtpunkte_13_1;
                                    //    }
                                    //}
                                    //if (leistung.EinheitNP == "P" && leistung.Jahrgang == 3 && leistung.HzJz == "GO" && leistung.Datum.Month > 2)
                                    //{
                                    //    if (leistung.Gesamtpunkte != leistung.Gesamtpunkte_13_2)
                                    //    {
                                    //        leistung.Gesamtpunkte = leistung.Gesamtpunkte_13_2;
                                    //    }
                                    //}
                                    leistung.Tendenz = theRow["Tendenz"].ToString() == "" ? null : theRow["Tendenz"].ToString();
                                    leistung.EinheitNP = theRow["Einheit"].ToString() == "" ? "N" : theRow["Einheit"].ToString();
                                    leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString());
                                    leistung.HzJz = theRow["HzJz"].ToString();
                                    leistung.Anlage = theRow["Anlage"].ToString();
                                    leistung.Zeugnisart = theRow["Zeugnisart"].ToString();
                                    leistung.Zeugnistext = theRow["Zeugnistext"].ToString();
                                    leistung.Konferenzdatum = theRow["Konferenzdatum"].ToString().Length < 3 ? new DateTime() : (DateTime.ParseExact(theRow["Konferenzdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)).AddHours(15);
                                    leistung.SchuelerAktivInDieserKlasse = theRow["SchuelerAktivInDieserKlasse"].ToString() == "J";
                                    leistung.Abschlussklasse = leistung.IstAbschlussklasse();
                                    leistung.Beschreibung = "";

                                    // Noten, deren Zeugnisdatum 3 Monate in der Vergangenheit liegt, sind geholte Noten.

                                    leistung.GeholteNote = false;

                                    if (leistung.Konferenzdatum.AddDays(90) < DateTime.Now && leistung
                                        .Konferenzdatum.Year > 1)
                                    {
                                        leistung.GeholteNote = true;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Global.AufConsoleSchreiben("Fehler beim Einlesen der Atlantis-Leistungsdatensätze: ENTER" + ex);
                                Console.ReadKey();
                            }

                            // Wenn für diesen Schüler eine Webuntis-Leistung existiert, wird die Atlantis-Leistung nicht geladen.

                            if (alleWebuntisLeistungen.AtlantisLeistungZiehen(leistung))
                            {
                                // Wenn für einen Schüler bereits eine neuere Leistung vorliegt, wird die ältere nicht mehr gezogen,

                                if (leistung.Konferenzdatum.Year > 1 && leistung.Konferenzdatum < DateTime.Now)
                                {
                                    if (!(from t in (from tt in this
                                                     where (tt.Konferenzdatum.Year > 1 && tt.Konferenzdatum < DateTime.Now) // es werden nur vergangene Leistungen
                                                     select tt).ToList()                                                    // miteinander verglichen.
                                          where t.Konferenzdatum > leistung.Konferenzdatum
                                          where t.Fach == leistung.Fach
                                          where t.SchlüsselExtern == leistung.SchlüsselExtern
                                          where t.Gesamtnote != null && t.Gesamtnote != ""
                                          select t).Any() || (leistung.Konferenzdatum.Year == 1 || leistung.Konferenzdatum > DateTime.Now))
                                    {
                                        leistungenUnsortiert.Add(leistung);
                                    }
                                }
                                else  // Leistungen des aktuellen Abschnitts werden immer gezogen.
                                {
                                    leistungenUnsortiert.Add(leistung);
                                }
                            }

                            if (leistung.Gesamtnote == "5" || leistung.Gesamtnote == "6")
                            {
                                Global.Defizitleistungen.Add(leistung);
                            }
                        }
                    }
                    connection.Close();

                    this.AddRange(leistungenUnsortiert.OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            var alleFächerDesAktuellenAbschnitts = (from l in this where (l.Konferenzdatum > DateTime.Now.AddDays(-20) || l.Konferenzdatum.Year == 1) select l.Fach).ToList();

            Global.AufConsoleSchreiben(("Leistungsdaten der SuS der Klasse " + interessierendeKlassen[0] + " aus Atlantis (" + Global.HzJz + ") ").PadRight(Global.PadRight, '.') + this.Count.ToString().PadLeft(4));

            var xx = (from l in this where l.GeholteNote select l).Count();

            Global.AufConsoleSchreiben((" ... davon Leistungsdaten aus abgeschlossenen Fächern alter Abschnitte ").PadRight(Global.PadRight,'.') + xx.ToString().PadLeft(4));            
        }

        private bool AtlantisLeistungZiehen(Leistung leistung)
        {
            // Wenn die Atlantisleistung keinem vergangenen Abschnitt zugeordnet werden kann oder in der Zukunft liegt, wird sie gezogen.

            if (leistung.Konferenzdatum.Year == 1 || leistung.Konferenzdatum>= DateTime.Now)
            {
                return true;
            }
            else
            {
                // Vergangene Leistungen ohne Note werden nicht gezogen

                if (leistung.Konferenzdatum.AddDays(90) < DateTime.Now && leistung.Gesamtnote == null)
                {
                    return false;
                }
            }

            // Wenn eine Webuntisleistung für diese Klasse und das selbe Fach existiert, wird sie 
            // nicht gezogen.

            foreach (var webuntisleistung in this)
            {
                if (webuntisleistung.Klasse == leistung.Klasse)
                {
                    if (webuntisleistung.Fach == leistung.Fach)
                    {
                        if (webuntisleistung.Gesamtnote != null && webuntisleistung.Gesamtnote != "")
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        internal void SprachenZuordnen(Leistungen atlantisLeistungen)
        {
            Console.Write(("Sprachen zuordnen ").PadRight(71, '.'));
            int i = 0;

            // Bei Sprachen-Fächern ...

            foreach (var we in (from t in this
                                where (
                                    t.Fach.Split(' ')[0] == "EK" ||
                                    t.Fach.Split(' ')[0] == "E" ||
                                    t.Fach.Split(' ')[0] == "N" ||
                                    t.Fach.Split(' ')[0] == "NL" ||
                                    t.Fach.Split(' ')[0] == "LA" ||
                                    t.Fach.Split(' ')[0] == "S" ||
                                    t.Fach.EndsWith("B1") ||
                                    t.Fach.EndsWith("B2") ||
                                    t.Fach.EndsWith("A1") ||
                                    t.Fach.EndsWith("A2")
                                )
                                select t).ToList())
            {
                var atlantisFächerDiesesSchuelers = (from a in atlantisLeistungen where a.Klasse == we.Klasse where a.SchlüsselExtern == we.SchlüsselExtern select a).ToList();
                int j = i;

                foreach (var a in atlantisFächerDiesesSchuelers)
                {
                    // Bei einer 1-zu-1-Zuordnung ist alles tutti.

                    if (a.Fach == we.Fach)
                    {
                        i++;
                        break;
                    }

                    // Bei Kursen (erkennbar am Leerzeichen) muss der Teil vor dem Leerzeichen übereinstimmen.

                    if (a.Fach.Contains(" ") && we.Fach.Contains(" ") && a.Fach.Split(' ')[0] == we.Fach.Split(' ')[0])
                    {
                        we.Beschreibung += we.Fach + "->" + a.Fach + ",";
                        i++;
                        break;
                    }

                    // Evtl. hängt die Niveaustufe in Untis am Namen 

                    if (we.Fach.EndsWith("A1") || we.Fach.EndsWith("A2") || we.Fach.EndsWith("B1") || we.Fach.EndsWith("B2"))
                    {

                    }
                }

                // Wenn keine Zuordnung vorgenommen werden konnte

                if (i == j)
                {
                    AusgabeSchreiben("Achtung: In der " + we.Klasse + " kann ein Fach in Webuntis keinem Atlantisfach im Notenblatt der Klasse zugeordnet werden:", new List<string>() { we.Fach });
                }

                // ... wenn keine 1:1-Zuordnung möglich ist ...

                //if (!(from a in atlantisLeistungen where a.Fach == we.Fach where a.Klasse == we.Klasse select a).Any())
                //{
                //    // wird versucht die Niveaustufe abzuschneiden

                //    if (!(from a in atlantisLeistungen where a.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "").Replace("KA2", "") == we.Fach where a.Klasse == we.Klasse select a).Any())
                //    {
                //        // wird versucht den Kurs abzuschneiden

                //        if ((we.Fach.Replace("  ", " ").Split(' ')).Count() > 1)
                //        {
                //            foreach (var a in atlantisLeistungen)
                //            {
                //                if (a.Klasse == we.Klasse)
                //                {
                //                    if (a.Fach.Split(' ')[0] == we.Fach.Split(' ')[0])
                //                    {
                //                        if ((a.Fach.Replace("  ", " ").Split(' ')).Count() > 1)
                //                        {
                //                            // Ziffer am Ende entfernen z.B. bei Kursen G1, G2 usw.

                //                            string wFachOhneZifferAmEnde = Regex.Replace(we.Fach, @"\d+$", "");
                //                            string aFachOhneZifferAmEnde = Regex.Replace(we.Fach, @"\d+$", "");

                //                            // alle Leerzeichen entfernen

                //                            string wFachOhneLeerzeichen = Regex.Replace(wFachOhneZifferAmEnde, @"\s+", "");
                //                            string aFachOhneLeerzeichen = Regex.Replace(aFachOhneZifferAmEnde, @"\s+", "");

                //                            if (wFachOhneLeerzeichen == aFachOhneLeerzeichen && we.Fach.Contains(" "))
                //                            {
                //                                we.Beschreibung += we.Fach + "->" + a.Fach + ",";
                //                                we.Fach = a.Fach;
                //                                i++;
                //                            }
                //                            else
                //                            {
                //                                AusgabeSchreiben("Achtung: In der " + we.Klasse + " kann ein Fach in Webuntis keinem Atlantisfach im Notenblatt der Klasse zugeordnet werden:", new List<string>() { we.Fach });
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }
                //    else
                //    {
                //        var f = (from a in atlantisLeistungen where a.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "").Replace("KA2", "") == we.Fach where a.Klasse == we.Klasse select a.Fach).FirstOrDefault();
                //        if (we.Fach != f)
                //        {
                //            we.Beschreibung += "|" + we.Fach + "->" + f;
                //            we.Fach = f;
                //        }
                //    }
                //}
                //else
                //{
                //    var f = (from a in atlantisLeistungen where a.Fach == we.Fach where a.Klasse == we.Klasse select a.Fach).FirstOrDefault();
                //    if (we.Fach != f)
                //    {
                //        we.Beschreibung += "|" + we.Fach + "->" + f;
                //        we.Fach = f;
                //    }
                //}
            }
            Global.AufConsoleSchreiben((" " + i).PadLeft(30, '.'));
        }

        internal void BindestrichfächerZuordnen(Leistungen atlantisLeistungen)
        {   
            int i = 0;

            // Bei Webuntis-Bindestrich-Fächern (z.B. VEF-GUP) 

            foreach (var we in (from t in this where t.Fach.Contains("-") select t).ToList())
            {
                // ... wenn keine 1:1-Zuordnung möglich ist ...

                if (!(from a in atlantisLeistungen where a.Fach == we.Fach select a).Any())
                {
                    // ... wird versucht das Fach aufzutrennen und dann zuzuordnen.

                    foreach (var fach in we.Fach.Split('-'))
                    {
                        var atlantisLeistung = (from a in atlantisLeistungen
                                                where a.Klasse == we.Klasse
                                                where a.Fach == fach
                                                where a.SchlüsselExtern == we.SchlüsselExtern
                                                select a).FirstOrDefault();

                        if (atlantisLeistung != null)
                        {
                            var x = (from wee in this where wee.Fach == we.Fach where wee.SchlüsselExtern == atlantisLeistung.SchlüsselExtern select wee).FirstOrDefault();
                            x.Beschreibung += we.Fach + "->" + fach + ",";
                            x.Fach = fach;
                            i++;
                        }
                    }
                }
            }
            Global.AufConsoleSchreiben("Bindestrichfächer zuordnen ".PadRight(110, '.') + i.ToString().PadLeft(4));
        }

        internal void ReligionsabwählerBehandeln(Leistungen atlantisLeistungen)
        {   
            int i = 0;

            // Für die verschiednenen Klassen ...

            foreach (var klasse in (from k in this select k.Klasse).Distinct().ToList())
            {
                // ... wenn Religion in der Klasse unterrichtet wird ...

                if (
                    (from t in this
                     where (t.Fach == "KR" || t.Fach == "ER" || t.Fach == "REL" || t.Fach.StartsWith("KR ") || t.Fach.StartsWith("ER "))
                     where t.Klasse == klasse
                     select t).Any())
                {
                    // ... wird bei allen Atlantis-Reli-Abwählern ...

                    foreach (var aLeistung in (from a in atlantisLeistungen
                                               where a.Klasse == klasse
                                               where (a.Fach == "KR" || a.Fach == "ER" || a.Fach == "REL" || a.Fach.StartsWith("KR ") || a.Fach.StartsWith("ER "))
                                               where a.ReligionAbgewählt
                                               select a).ToList())
                    {
                        var wf = ((from w in this
                                   where w.Klasse == klasse
                                   where w.SchlüsselExtern == aLeistung.SchlüsselExtern
                                   where (w.Fach == "KR" || w.Fach == "ER" || w.Fach == "REL" || w.Fach.StartsWith("KR ") || w.Fach.StartsWith("ER "))
                                   select w).FirstOrDefault());
                        if (wf == null)
                        {
                            // ... ein neuer Webuntis-Datensatz hinzugefügt und ein '-' gesetzt.

                            Leistung leistung = new Leistung();
                            leistung.Datum = new DateTime();
                            leistung.Name = aLeistung.Name;
                            leistung.Klasse = aLeistung.Klasse;
                            leistung.Fach = aLeistung.Fach;
                            leistung.Gesamtnote = "-";
                            leistung.Gesamtpunkte = "99";
                            leistung.Bemerkung = "";
                            leistung.Lehrkraft = "";
                            leistung.Beschreibung = "abgewählt,";
                            leistung.SchlüsselExtern = aLeistung.SchlüsselExtern;
                            this.Add(leistung);
                            i++;
                        }
                        else
                        {
                            wf.Gesamtnote = "-";
                        }
                    }
                }
            }
            Global.AufConsoleSchreiben("Religionsabwähler mit '-' versehen ".PadRight(110, '.') + i.ToString().PadLeft(4));
        }

        internal void ReligionZuordnen()
        {
            Console.Write(("Religion zuordnen").PadRight(71, '.'));
            int i = 0;
            foreach (var leistung in this)
            {
                if (leistung.Fach == "KR" || leistung.Fach == "ER" || leistung.Fach.StartsWith("KR ") || leistung.Fach.StartsWith("ER "))
                {
                    leistung.Beschreibung += leistung.Fach + "->REL,";
                    leistung.Fach = "REL";
                    i++;
                }
            }
            Global.AufConsoleSchreiben((" " + i.ToString()).PadLeft(30, '.'));
        }

        internal void WeitereFächerZuordnen(Leistungen atlantisLeistungen)
        {
            Console.Write(("Weitere Fächer zuordnen").PadRight(71, '.'));
            int i = 0;

            // Für die verschiednenen Klassen ...

            foreach (var klasse in (from k in this select k.Klasse).Distinct())
            {
                // ... und deren verschiedene Fächer aus Webuntis, in denen auch Noten gegeben wurden (also kein FU) ... 

                foreach (var webuntisFach in (from t in this
                                              where t.Klasse == klasse
                                              where t.Gesamtnote != null
                                              where t.Fach != "KR" && t.Fach != "ER" && !t.Fach.StartsWith("KR ") && !t.Fach.StartsWith("ER ") && !t.Fach.Contains("-")
                                              select t.Fach).Distinct())
                {
                    // ... wenn es in Atlantis kein entsprechendes Fach gibt ...

                    if (!(from a in atlantisLeistungen where a.Fach == webuntisFach select a).Any())
                    {
                        // ... wird geprüft, ob es ein Fach in Atlantis gibt, dass den Anfangsbuchstaben teilt und gleichzeitig keine Entsprechung in Webuntis hat.

                        var afa = (from a in atlantisLeistungen where a.Klasse == klasse where a.Fach.Substring(0, 1) == webuntisFach.Substring(0, 1) select a.Fach).Distinct().ToList();

                        var af = (from a in atlantisLeistungen where a.Klasse == klasse where a.Fach.Substring(0, 1) == webuntisFach.Substring(0, 1) select a).FirstOrDefault();

                        // ... und dass seinerseits in Webuntis nicht vorkommt ...

                        if (afa.Count == 1 && !(from w in this where w.Fach == af.Fach select w).Any())
                        {
                            AusgabeSchreiben("Klasse: " + klasse + ". Das Webuntisfach " + webuntisFach + " wird dem Atlantisfach " + af.Fach + " zugeordnet. Wenn das nicht gewünscht ist, dann hier mit Suchen & Ersetzen der gewünschte Name gesetzt werden.", new List<string>());

                            foreach (var leistung in this)
                            {
                                if (leistung.Fach == webuntisFach && leistung.Klasse == klasse)
                                {
                                    leistung.Fach = af.Fach;
                                    leistung.Beschreibung += leistung.Fach + "->" + af.Fach + ",";
                                    i++;
                                }
                            }
                        }
                        else
                        {
                            AusgabeSchreiben("Achtung: In der " + klasse + " kann ein Fach in Webuntis keinem Atlantisfach im Notenblatt der Klasse zugeordnet werden:", new List<string>() { webuntisFach });
                        }
                    }
                }
            }
            Global.AufConsoleSchreiben((" " + i.ToString()).PadLeft(30, '.'));
        }

        public Leistungen()
        {
        }

        public Leistungen(IEnumerable<Leistung> collection) : base(collection)
        {
            foreach (var item in (from a in collection where a.Gesamtnote == "5" || a.Gesamtnote == "6" select a).ToList())
            {
                this.Add(item);
            }
        }

        internal void Add(Leistungen atlantisLeistungen)
        {
            int outputIndex = Global.SqlZeilen.Count();
            
            int i = 0;

            foreach (var w in (from t in this where t.Beschreibung.Contains("NEU") select t).ToList())
            {
                //UpdateLeistung(w.Klasse.PadRight(6) + "|" + (w.Name.Substring(0, Math.Min(w.Name.Length, 6))).PadRight(6) + "|" + w.Gesamtnote + " " + "|" + w.Fach.PadRight(5) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET s_note='" + w.Gesamtnote + "', ls_id_1=" + w.LehrkraftAtlantisId + " WHERE noe_id=" + w.ZielLeistungId + ";", w.Datum);
                i++;
            }

            Global.AufConsoleSchreiben(("Neu anzulegende Leistungen in Atlantis ").PadRight(Global.PadRight, '.') + i.ToString().PadLeft(4));
            Global.PrintMessage(outputIndex, ("Neu anzulegende Leistungen in Atlantis: ").PadRight(65, '.') + (" " + i.ToString()).PadLeft(30, '.'));
        }

        public string Gesamtpunkte2Gesamtnote(string gesamtpunkte)
        {
            if (gesamtpunkte == "0")
            {
                return "6";
            }
            if (gesamtpunkte == "1")
            {
                return "5";
            }
            if (gesamtpunkte == "2")
            {
                return "5";
            }
            if (gesamtpunkte == "3")
            {
                return "5";
            }
            if (gesamtpunkte == "4")
            {
                return "4";
            }
            if (gesamtpunkte == "5")
            {
                return "4";
            }
            if (gesamtpunkte == "6")
            {
                return "4";
            }
            if (gesamtpunkte == "7")
            {
                return "3";
            }
            if (gesamtpunkte == "8")
            {
                return "3";
            }
            if (gesamtpunkte == "9")
            {
                return "3";
            }
            if (gesamtpunkte == "10")
            {
                return "2";
            }
            if (gesamtpunkte == "11")
            {
                return "2";
            }
            if (gesamtpunkte == "12")
            {
                return "2";
            }
            if (gesamtpunkte == "13")
            {
                return "1";
            }
            if (gesamtpunkte == "14")
            {
                return "1";
            }
            if (gesamtpunkte == "15")
            {
                return "1";
            }
            if (gesamtpunkte == "84")
            {
                return "A";
            }
            if (gesamtpunkte == "99")
            {
                return "-";
            }
            return null;
        }

        internal string Update(Leistungen atlantisLeistungen)
        {
            if (atlantisLeistungen is null)
            {
                throw new ArgumentNullException(nameof(atlantisLeistungen));
            }

            int outputIndex = Global.SqlZeilen.Count();

            int i = 0;

            Global.AufConsoleSchreiben((" "));
            Global.AufConsoleSchreiben(("Leistungen in Atlantis updaten: "));

            foreach (var w in (from t in this.OrderBy(x=>x.Name) where t.Beschreibung != null where t.Query!= null select t).ToList())
            {
                UpdateLeistung(w.Beschreibung, w.Query);
                i++;
            }

            if (i==0)
            {
                return "A C H T U N G: Es wurde keine einzige Leistung angelegt oder verändert. Das kann daran liegen, dass das Notenblatt in Atlantis nicht richtig angelegt ist. Die Tabelle zum manuellen eintragen der Noten muss in der Notensammelerfassung sichtbar sein.";
            }
            return "";
        }

        internal void Delete(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen, List<string> aktSj)
        {   
            int outputIndex = Global.SqlZeilen.Count();                        
            int i = 0;

            try
            {
                foreach (var a in this)
                {
                    if (a.SchuelerAktivInDieserKlasse)
                    {
                        // Wenn es zu einem Atlantis-Datensatz keine Entsprechung in Webuntis gibt, ... 
                        
                        var webuntisLeistung = (from w in webuntisLeistungen
                                                where w.Zielfach == a.Fach
                                                where w.Klasse == a.Klasse
                                                where w.SchlüsselExtern == a.SchlüsselExtern
                                                where w.Gesamtpunkte != null
                                                select w).FirstOrDefault();

                        if (webuntisLeistung == null)
                        {
                            if (a.Gesamtnote != null)
                            {
                                // Wenn die Zeugniskonferenz mehr als drei Tage hinter uns liegt, wird nicht mehr gelöscht.

                                if (DateTime.Now <= a.Konferenzdatum)
                                {
                                    // Geholte Noten aus Vorjahren werden nicht gelöscht.

                                    if (!a.GeholteNote)
                                    {
                                        //UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 6)) + "|" + a.Fach.Substring(0, Math.Min(a.Fach.Length, 3)).PadRight(3) + "|" + (a.Gesamtnote == null ? "NULL" : a.Gesamtnote + (a.Tendenz == null ? " " : a.Tendenz)) + ">NULL" + a.Beschreibung, "UPDATE noten_einzel SET      s_note=NULL WHERE noe_id=" + a.LeistungId + ";", new DateTime(1, 1, 1));

                                        if (a.Gesamtpunkte != null)
                                        {
                                            //UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 6)) + "|" + a.Fach.Substring(0, Math.Min(a.Fach.Length, 3)).PadRight(3) + "|" + ((a.Gesamtpunkte == null ? "NULL" : a.Gesamtpunkte.Split(',')[0])).PadLeft(2) + ">NULL" + a.Beschreibung, "UPDATE noten_einzel SET      punkte=NULL WHERE noe_id=" + a.LeistungId + ";", new DateTime(1, 1, 1));
                                        }

                                        if (a.Tendenz != null)
                                        {
                                            //UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 6)) + "|" + a.Fach.Substring(0, Math.Min(a.Fach.Length, 3)).PadRight(3) + "|" + (a.Tendenz == null ? "NULL" : " " + a.Tendenz) + ">NULL" + a.Beschreibung, "UPDATE noten_einzel SET s_tendenz=NULL WHERE noe_id=" + a.LeistungId + ";", new DateTime(1, 1, 1));
                                        }
                                        i++;
                                    }
                                }
                            }                            
                        }
                    }
                }                
                Global.AufConsoleSchreiben(("Zu löschende Leistungen in Atlantis ").PadRight(71, '.') + i.ToString().PadLeft(4));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
                
        internal string GetIntessierendeKlassen(List<string> aktSj)
        {
            try
            {                
                do
                {
                    
                    Console.Write("Bitte eine Klasse wählen: ");

                    var x = Console.ReadLine();

                    foreach (var item in this)
                    {
                        if (item.Klasse != "" && item.Klasse == x.ToUpper())
                        {
                            return x.ToUpper();
                        }
                    }
                    
                    
                } while (true);                
            }
            catch (Exception ex)
            {
                Global.AufConsoleSchreiben("Bei der Auswahl der interessierenden Klasse ist es zum Fehler gekommen. \n " + ex);
            }
            return null;
        }

        public void AusgabeSchreiben(string text, List<string> klassen)
        {
            try
            {                
                int z = 0;

                do
                {
                    var zeile = "";

                    try
                    {
                        while ((zeile + text.Split(' ')[z] + ", ").Length <= 96)
                        {
                            zeile += text.Split(' ')[z] + " ";
                            z++;
                        }
                    }
                    catch (Exception)
                    {
                        z++;
                        zeile.TrimEnd(',');
                    }

                    zeile = zeile.TrimEnd(' ');

                    string o = "/* " + zeile.TrimEnd(' ');
                    Global.SqlZeilen.Add((o.Substring(0, Math.Min(101, o.Length))).PadRight(181) + "*/");

                } while (z < text.Split(' ').Count());

                z = 0;

                do
                {
                    var zeile = " ";

                    try
                    {
                        if (klassen[z].Length >= 95)
                        {
                            klassen[z] = klassen[z].Substring(0, Math.Min(klassen[z].Length, 150));
                            zeile += klassen[z];
                            throw new Exception();
                        }

                        while ((zeile + klassen[z] + ", ").Length <= 97)
                        {
                            zeile += klassen[z] + ", ";
                            z++;
                        }
                    }
                    catch (Exception)
                    {
                        z++;
                        zeile.TrimEnd(',');
                    }

                    zeile = zeile.TrimEnd(' ');
                    int s = zeile.Length;
                    string o = "/* " + zeile.TrimEnd(',');
                    Global.SqlZeilen.Add((o.Substring(0, Math.Min(101, o.Length))).PadRight(181) + "*/");

                } while (z < klassen.Count);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        private void UpdateLeistung(string message, string updateQuery)
        {
            try
            {
                string o = updateQuery.PadRight(103, ' ') + (updateQuery.Contains("Keine Änderung")?"   ": "/* ") + message;
                Global.SqlZeilen.Add((o.Substring(0, Math.Min(178, o.Length))).PadRight(180) + " */");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void ErzeugeSqlDatei(List<string> files)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(files[2], true, Encoding.Default))
                {
                    foreach (var o in Global.SqlZeilen)
                    {
                        writer.WriteLine(o);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }       
    }
}