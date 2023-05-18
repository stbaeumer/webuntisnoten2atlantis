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
        public Leistungen(string targetMarksPerLesson, Lehrers lehrers)
        {
            var leistungen = new Leistungen();
            var leistungenOhneWidersprüchlicheNoten = new Leistungen();

            using (StreamReader reader = new StreamReader(targetMarksPerLesson))
            {
                var überschrift = reader.ReadLine();
                int i = 1;

                while (true)
                {
                    i++;
                    Leistung leistung = new Leistung();

                    string line = reader.ReadLine();

                    try
                    {
                        if (line != null)
                        {
                            var x = line.Split('\t');

                            if (x.Length == 10)
                            {
                                leistung = new Leistung();
                                leistung.MarksPerLessonZeile = i;
                                leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                leistung.Name = x[1];
                                leistung.Klasse = x[2];
                                leistung.Fach = x[3];
                                leistung.Prüfungsart = x[4];
                                leistung.Note = Gesamtpunkte2Gesamtnote(x[5]); // Für die Blauen Briefe
                                leistung.Punkte = x[5].Length > 0 ? x[5].Substring(0, 1) : ""; // Für die blauen Briefe
                                leistung.Gesamtpunkte = x[9].Split('.')[0] == "" ? "" : x[9].Split('.')[0];
                                leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);
                                leistung.Tendenz = Gesamtpunkte2Tendenz(leistung.Gesamtpunkte);
                                leistung.Bemerkung = x[6];
                                leistung.Lehrkraft = x[7];
                                leistung.LehrkraftAtlantisId = (from l in lehrers where l.Kuerzel == leistung.Lehrkraft select l.AtlantisId).FirstOrDefault();
                                leistung.SchlüsselExtern = Convert.ToInt32(x[8]);

                                if (leistung.Fach != null && leistung.Fach != "")
                                {
                                    leistung.GetFachAliases();
                                    this.Add(leistung);
                                }
                            }

                            // Wenn in den Bemerkungen eine zusätzlicher Umbruch eingebaut wurde:

                            if (x.Length == 7)
                            {
                                leistung = new Leistung();
                                leistung.MarksPerLessonZeile = i;
                                leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                leistung.Name = x[1];
                                leistung.Klasse = x[2];
                                leistung.Fach = x[3];
                                leistung.Prüfungsart = x[4];
                                leistung.Note = x[5];
                                leistung.Bemerkung = x[6];
                            }

                            if (x.Length == 4)
                            {
                                leistung.Lehrkraft = x[1];
                                leistung.SchlüsselExtern = Convert.ToInt32(x[2]);
                                leistung.Gesamtpunkte = x[3].Split('.')[0];
                                leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);

                                if (leistung.Fach != null && leistung.Fach != "")
                                {
                                    leistung.GetFachAliases();
                                    this.Add(leistung);
                                }
                            }

                            if (x.Length < 4)
                            {
                                Global.WriteLine("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
                                Console.ReadKey();
                                //OpenFiles(new List<string>() { targetMarksPerLesson });
                                throw new Exception("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
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
            }               

            Global.WriteLine(("Alle Webuntis-Leistungen (mit & ohne Gesamtnote; ohne Dopplungen) ").PadRight(Global.PadRight - 2, '.') + this.Count.ToString().PadLeft(6));
        }

        internal void Hinzufügen(Leistungen webuntisLeistungen)
        {
            this.Add(webuntisLeistungen);
        }

        internal Leistungen GetIntessierendeWebuntisLeistungen(List<int> interessierendeSchülerDieserKlasse, string interessierendeKlasse)
        {
            // Wenn ein Schüler im laufenden Schuljahr herabgestuft wurde, kann es sein, dass die MarksPerLesson
            // fälschlicherweise Leistungen aus der anderen Klasse in die neue Klasse übernimmt. Es sieht dann so aus,
            // als würde das Fach von dem Kollegen der höheren Klasse auch in dieser Klasse unterrichtet werden.

            var zurückgestufteSuS = GetAndereZurückgestufteSuS(interessierendeSchülerDieserKlasse, interessierendeKlasse);
            var verworfeneLeistungen = new List<Leistung>();

            var leistungen = new Leistungen();

            foreach (var wl in (from t in this where interessierendeSchülerDieserKlasse.Contains(t.SchlüsselExtern) select t).ToList())
            {
                // Die Leistung eines zurückgestuften Schülers wird verworfen, ...

                if (zurückgestufteSuS.Contains(wl.SchlüsselExtern))
                {
                    // ... wenn kein anderer nicht zurückgestufter Schüler in diesem Fach bisher eine Leistung bekommen hat.

                    if ((from t in this
                         where t.Klasse == interessierendeKlasse
                         where interessierendeSchülerDieserKlasse.Contains(t.SchlüsselExtern)
                         where !zurückgestufteSuS.Contains(t.SchlüsselExtern)
                         where t.Fach == wl.Fach
                         where t.Gesamtnote != ""
                         select t
                         ).Any())
                    {
                        // Leistungen aus der Parallelklasse werden berücksichtigt.

                        if (wl.Klasse.Substring(0, wl.Klasse.Length - 1) == interessierendeKlasse.Substring(0, interessierendeKlasse.Length - 1))
                        {
                            leistungen.Add(wl);
                        }
                    }
                    else
                    {
                        verworfeneLeistungen.Add(wl);
                    }
                }
                else
                {
                    // Leistungen aus der Parallelklasse werden berücksichtigt.

                    if (wl.Klasse.Substring(0, wl.Klasse.Length - 1) == interessierendeKlasse.Substring(0, interessierendeKlasse.Length - 1))
                    {
                        leistungen.Add(wl);
                    }
                }
            }

            if (verworfeneLeistungen.Any())
            {
                Global.WriteLine("Es scheint Schüler zu geben, die im laufenden SJ in die Klasse " + interessierendeKlasse + " gewechselt sind. In der MarksPerLesson kann");
                Global.WriteLine(" es dann fälschlicherweise dazu kommen, dass Leistungen aus einer alten Klasse in die neue Klasse übernommen werden.");
                Global.WriteLine(" Folgende Leistungen aus Webuntis werden also verworfen, weil niemand sonst bisher eine Gesamtote bekommen hat:");

                foreach (var z in zurückgestufteSuS)
                {
                    var xx = (from f in verworfeneLeistungen where f.SchlüsselExtern == z select f).ToList();

                    if (xx.Any())
                    {
                        var vorherigeKlasse = (from w in this where w.SchlüsselExtern == z where w.Klasse != interessierendeKlasse select w.Klasse).FirstOrDefault();
                        var verworfenesFachString = Global.List2String((from f in verworfeneLeistungen where f.SchlüsselExtern == z select f.Fach + "(Zeile:" + f.MarksPerLessonZeile + ")").ToList(), ",");
                        Global.WriteLine("  " + xx[0].Name + ", vorher:" + vorherigeKlasse.PadRight(7) + ": " + verworfenesFachString);
                    }
                }
                Global.WriteLine("  ");
            }
            return leistungen;
        }

        private List<int> GetAndereZurückgestufteSuS(List<int> interessierendeSuS, string interessierendeKlasse)
        {
            List<int> andere = new List<int>();

            // Für jeden interessierenden S wird geprüft, ...

            foreach (var i in interessierendeSuS)
            {
                // ... ob er mehr als einer Klasse zugeordnet werden kann ...

                if ((from w in this where w.SchlüsselExtern == i select w.Klasse).Distinct().Count() > 1)
                {
                    // ... und ob es nicht die Parallelklassse ist, da Leistungen aus der Parallelklasse nicht verworfen werden ...

                    var dieAndereKlasse = (from w in this where w.SchlüsselExtern == i where w.Klasse != interessierendeKlasse select w.Klasse).FirstOrDefault();

                    if (dieAndereKlasse.Substring(0, dieAndereKlasse.Length - 1) != interessierendeKlasse.Substring(0, interessierendeKlasse.Length - 1))
                    {
                        andere.Add(i);
                    }
                }
            }
            return andere;
        }

        internal Leistungen GetBlaueBriefeLeistungen()
        {
            var leistungen = new Leistungen();

            foreach (var leistung in this)
            {
                if (leistung.Prüfungsart.StartsWith("M"))
                {
                    leistung.Gesamtnote = leistung.Note;
                    leistung.Gesamtpunkte = leistung.Punkte;
                    leistungen.Add(leistung);
                }
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

            var leistungenÄhnlicheFächerMitUnterschiedlichenNoten = (from w in this
                                                           where w.Klasse == wLeistung.Klasse
                                                           where w.SchlüsselExtern == wLeistung.SchlüsselExtern
                                                           where ähnlicheFächer.Contains(w.Fach)
                                                           select w).ToList();

            Global.WriteLine("Widersprechende Gesamtnoten im Fach " + Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value + "* beim Schüler " + wLeistung.SchlüsselExtern + ":");

            for (int i = 0; i < leistungenÄhnlicheFächerMitUnterschiedlichenNoten.Count; i++)
            {
                Global.WriteLine((i + 1) + ". " + leistungenÄhnlicheFächerMitUnterschiedlichenNoten[i].SchlüsselExtern.ToString().PadRight(7) + leistungenÄhnlicheFächerMitUnterschiedlichenNoten[i].Name.ToString().PadRight(25) + leistungenÄhnlicheFächerMitUnterschiedlichenNoten[i].Lehrkraft.PadRight(4) + leistungenÄhnlicheFächerMitUnterschiedlichenNoten[i].Fach.PadRight(6) + "Note: " + leistungenÄhnlicheFächerMitUnterschiedlichenNoten[i].Gesamtnote + leistungenÄhnlicheFächerMitUnterschiedlichenNoten[i].Tendenz + " (Punkte: " + leistungenÄhnlicheFächerMitUnterschiedlichenNoten[i].Gesamtpunkte +")");
            }

            bool wiederholen = true;

            int index = 0;

            do
            {
                Console.WriteLine("Bitte eine Zahl zwischen 1 und " + leistungenÄhnlicheFächerMitUnterschiedlichenNoten.Count + " eingeben, um generell zwischen den genannten LuL in der " + wLeistung.Klasse + " im Fach " + Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value + "* zu wählen:");

                ConsoleKeyInfo eingabe;

                eingabe = Console.ReadKey();
                 
                if (char.IsDigit(eingabe.KeyChar))
                {
                    index = int.Parse(eingabe.KeyChar.ToString()); // use Parse if it's a Digit

                    if (index > 0 && index <= leistungenÄhnlicheFächerMitUnterschiedlichenNoten.Count)
                    {
                        wiederholen = false;
                    }
                }
            } while (wiederholen);

            Global.WriteLine("Sie haben " + index + " gewählt.");

            return wLeistung.Klasse + "|" + Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value + "|" + leistungenÄhnlicheFächerMitUnterschiedlichenNoten[index-1].Lehrkraft;
        }

        private bool EsGibtWidersprechendeNoten(Leistung wLeistung)
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
                                                           where w.Gesamtnote != null
                                                           where w.Gesamtnote != ""
                                                           where w.Gesamtpunkte != ""
                                                           select w.Gesamtpunkte).Distinct().ToList();

            if (ähnlicheFächerMitUnterschiedlichenNoten.Count > 1)
            {
                return true;
            }
            return false;
        }

        private bool EsGibtDoppelteFächerZuDieser(Leistung wLeistung)
        {
            var ähnlicheFächer = (from w in this
                                  where w.Klasse == wLeistung.Klasse
                                  where w.SchlüsselExtern == wLeistung.SchlüsselExtern
                                  where Regex.Match(w.Fach, @"^[^0-9]*").Value == Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value
                                  select w.Fach).Distinct().ToList();

            if (ähnlicheFächer.Count > 1)
            {
                var ähnlicher = (from w in this                                 
                                 where w.Klasse == wLeistung.Klasse
                                 where w.SchlüsselExtern == wLeistung.SchlüsselExtern
                                 where Regex.Match(w.Fach, @"^[^0-9]*").Value == Regex.Match(wLeistung.Fach, @"^[^0-9]*").Value
                                 where w.Gesamtpunkte != ""
                                 select w).ToList();

                return true;
            }
            return false;
        }

        //internal void ErzeugeSerienbriefquelleFehlendeTeilleistungen(Leistungen webuntisLeistungen)
        //{
        //    List<Leistung> fehlendeTeilleistungen = new List<Leistung>();

        //    foreach (var atlantisLeistung in (from t in this where t.Fach != "" where !(t.Zeugnistext.StartsWith("Abschluss")) select t).ToList())
        //    {
        //        // Wenn es für eine Atlantisleistung keine Entsprechung in Untis gibt, wird der Datensatz hinzugefügt.

        //        if (!(from w in webuntisLeistungen where w.Klasse == atlantisLeistung.Klasse where w.Fach == atlantisLeistung.Fach select w).Any())
        //        {
        //            fehlendeTeilleistungen.Add(atlantisLeistung);
        //        }
        //    }
        //}

        //internal void Gym12NotenInDasGostNotenblattKopieren(List<string> interessierendeKlassen, List<string> aktSj)
        //{
        //    foreach (var item in this)
        //    {
        //        // Wenn es sich um die Leistung einer Gym 12 oder 13 handelt ...

        //        if (item.Klasse.StartsWith("G") && (item.Klasse.Contains((Convert.ToInt32(aktSj[1]) - 2).ToString()) || item.Klasse.Contains((Convert.ToInt32(aktSj[1]) - 3).ToString())))
        //        {
        //            // ... wird die Leistung in das Notenblatt der GO geschrieben.

        //            throw new NotImplementedException();

        //        }
        //    }
        //}

        //internal void ErzeugeSerienbriefquelleNichtversetzer()
        //{
        //    var l = new Leistungen();

        //    foreach (var leistung in this)
        //    {
        //        if (leistung.SchlüsselExtern == 153203)
        //        {
        //            l.Add(leistung);
        //        }
        //    }
        //}

        //internal void AtlantisLeistungenZuordnenUndQueryBauen(Leistungen atlantisLeistungen, string aktuellesSchuljahr, string interessierendeKlasse, string hzJz)
        //{
        //    var kombinationAusKlasseUndFachWirdNurEinmalAufgerufen = new List<string>();

        //    Zuordnungen gespeicherteZuordnungen = new Zuordnungen();

        //    // Es wird versucht jedes Webuntis-Fach einem Atlantis-Fach zuzuordnen

        //    foreach (var webuntisLeistung in this.OrderBy(x => x.Klasse).ThenBy(x => x.Fach))
        //    {                
        //        // Zielleistungen dieses Schülers im aktuellen Abschnitt werden ermittelt.

        //        var atlantisZielLeistungen = (from a in atlantisLeistungen
        //                                      where !a.IstGeholteNote // eine Leistung des aktuellen Abschnitts
        //                                      where a.SchlüsselExtern == webuntisLeistung.SchlüsselExtern
        //                                      where a.Schuljahr == aktuellesSchuljahr
        //                                      where a.HzJz == hzJz
        //                                      orderby a.LeistungId descending // Die höchste ID ist die ID des aktuellen Abschnitts
        //                                      select a).ToList();

        //        // Wenn es den Schüler überhaupt nicht in Atlantis gibt, dann wird die Verarbeitung abgebrochen

        //        if (atlantisZielLeistungen.Count > 0 && !kombinationAusKlasseUndFachWirdNurEinmalAufgerufen.Contains(webuntisLeistung.Klasse + webuntisLeistung.Fach))
        //        {
        //            // Im Idealfall stimmen die Fachnamen überein.

        //            if (webuntisLeistung.Query == null)
        //            {
        //                var aL = (from a in atlantisZielLeistungen.OrderBy(x => x.Bereich)
        //                          where a.Fach == webuntisLeistung.Fach // Prüfung auf Fachnamensgleichheit
        //                          select a).ToList();

        //                if (aL.Count > 0)
        //                {
        //                    // Wenn der Schüler Reli-Abwähler ist, dürfte es eigentlich diese Webuntis-Leistung nicht geben.

        //                    if (aL[0].ReligionAbgewählt && aL[0].FachAliases.Contains("REL"))
        //                    {
        //                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "Abwähler!");
        //                    }
        //                    else
        //                    {
        //                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "");
        //                    }
        //                }
        //            }
        //            // Suffix entfernen

        //            if (webuntisLeistung.Query == null)
        //            {
        //                if (char.IsDigit(webuntisLeistung.Fach.Last()))
        //                {
        //                    var aL = (from a in atlantisZielLeistungen
        //                              where a.Fach == Regex.Match(webuntisLeistung.Fach, @"^[^0-9]*").Value
        //                              select a).ToList();

        //                    webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->" + Regex.Match(webuntisLeistung.Fach, @"^[^0-9]*").Value + "|");
        //                }
        //            }

        //            // Religion wird immer zu REL

        //            if (webuntisLeistung.Query == null)
        //            {
        //                if (webuntisLeistung.Fach == "KR" || webuntisLeistung.Fach == "ER" || webuntisLeistung.Fach.StartsWith("KR ") || webuntisLeistung.Fach.StartsWith("ER "))
        //                {
        //                    var aL = (from a in atlantisZielLeistungen
        //                              where a.Fach == "REL"
        //                              select a).ToList();

        //                    // Wenn der Schüler Reli-Abwähler ist, dürfte es eigentlich diese Webuntis-Leistung nicht geben.

        //                    if (aL[0].ReligionAbgewählt)
        //                    {
        //                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->REL|Abwähler!");
        //                    }
        //                    else
        //                    {
        //                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->REL|");
        //                    }                            
        //                }
        //            }

        //            // EK -> E

        //            if (webuntisLeistung.Query == null)
        //            {
        //                if (webuntisLeistung.Fach.StartsWith("EK") && webuntisLeistung.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "E")
        //                {
        //                    var aL = (from a in atlantisZielLeistungen
        //                              where a.Fach == "E"
        //                              select a).ToList();

        //                    webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->E|");
        //                }
        //                if (webuntisLeistung.Fach.StartsWith("EB1"))
        //                {
        //                    var aL = (from a in atlantisZielLeistungen
        //                              where a.Fach.StartsWith("E")
        //                              where a.Fach.Contains("B1")
        //                              select a).ToList();

        //                    webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->EB1|");
        //                }
        //            }

        //            // N -> NKA1

        //            if (webuntisLeistung.Query == null)
        //            {
        //                if (webuntisLeistung.Fach.StartsWith("N"))
        //                {
        //                    if (webuntisLeistung.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "NK")
        //                    {
        //                        var aL = (from a in atlantisZielLeistungen
        //                                  where a.Fach == "N"
        //                                  select a).ToList();

        //                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->N|");
        //                    }
        //                }
        //            }

        //            // NB1 -> N

        //            if (webuntisLeistung.Query == null)
        //            {
        //                if (webuntisLeistung.Fach.StartsWith("NB1"))
        //                {
        //                    if (webuntisLeistung.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "N")
        //                    {
        //                        var aL = (from a in atlantisZielLeistungen
        //                                  where a.Fach == "N"
        //                                  select a).ToList();

        //                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->N|");
        //                    }
        //                }
        //            }

        //            // EB2 -> N

        //            if (webuntisLeistung.Query == null)
        //            {
        //                if (webuntisLeistung.Fach.StartsWith("EB1") || webuntisLeistung.Fach.StartsWith("EB2"))
        //                {
        //                    if (webuntisLeistung.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "E")
        //                    {
        //                        var aL = (from a in atlantisZielLeistungen
        //                                  where a.Fach == "E"
        //                                  select a).ToList();

        //                        webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->E|");
        //                    }
        //                }
        //            }

        //            // Kurse mit Leerzeichen

        //            if (webuntisLeistung.Query == null)
        //            {
        //                var aL = (from a in atlantisZielLeistungen
        //                          where (webuntisLeistung.Fach.Contains(" ") && a.Fach.Contains(" ")) // Wenn beide Leerzeichen enthalten
        //                          where (webuntisLeistung.Fach.Split(' ')[0] == a.Fach.Split(' ')[0]) // Wenn der erste Teil identisch ist
        //                          where (webuntisLeistung.Fach.Substring(webuntisLeistung.Fach.LastIndexOf(' ') + 1).Substring(0, 1) == a.Fach.Substring(a.Fach.LastIndexOf(' ') + 1).Substring(0, 1)) // Wenn der zweite Teil identisch anfängt
        //                          select a).ToList();

        //                webuntisLeistung.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->" + (aL.Count > 0 ? aL[0].Fach : "") + "|");

        //                // Kurse mit Leerzeichen

        //                // Wenn in Atlantis die Anzahl der Zielfächer größer als 1 ist, dann wird eine weitere Leistung hinzugefügt 

        //                if (aL.Count > 1)
        //                {
        //                    Leistung neueLeistung = new Leistung();

        //                    neueLeistung.Abschlussklasse = webuntisLeistung.Abschlussklasse;
        //                    neueLeistung.Anlage = webuntisLeistung.Anlage;
        //                    neueLeistung.Bemerkung = webuntisLeistung.Bemerkung;
        //                    neueLeistung.Beschreibung = webuntisLeistung.Fach + "->" + aL[1] + ",";
        //                    neueLeistung.Datum = webuntisLeistung.Datum;
        //                    neueLeistung.EinheitNP = webuntisLeistung.EinheitNP;
        //                    neueLeistung.Fach = aL[1].Fach;
        //                    neueLeistung.Zielfach = aL[1].Fach;
        //                    neueLeistung.IstGeholteNote = webuntisLeistung.IstGeholteNote;
        //                    neueLeistung.Gesamtnote = webuntisLeistung.Gesamtnote;
        //                    neueLeistung.Gesamtpunkte = webuntisLeistung.Gesamtpunkte;
        //                    neueLeistung.Gliederung = webuntisLeistung.Gliederung;
        //                    neueLeistung.HatBemerkung = webuntisLeistung.HatBemerkung;
        //                    neueLeistung.HzJz = webuntisLeistung.HzJz;
        //                    neueLeistung.Jahrgang = webuntisLeistung.Jahrgang;
        //                    neueLeistung.Klasse = webuntisLeistung.Klasse;
        //                    neueLeistung.Konferenzdatum = webuntisLeistung.Konferenzdatum;
        //                    neueLeistung.Lehrkraft = webuntisLeistung.Lehrkraft;
        //                    neueLeistung.LeistungId = webuntisLeistung.LeistungId;
        //                    neueLeistung.Name = webuntisLeistung.Name;
        //                    neueLeistung.ReligionAbgewählt = webuntisLeistung.ReligionAbgewählt;
        //                    neueLeistung.SchlüsselExtern = webuntisLeistung.SchlüsselExtern;
        //                    neueLeistung.SchuelerAktivInDieserKlasse = webuntisLeistung.SchuelerAktivInDieserKlasse;
        //                    neueLeistung.Schuljahr = webuntisLeistung.Schuljahr;
        //                    neueLeistung.Tendenz = webuntisLeistung.Tendenz;
        //                    neueLeistung.Zeugnisart = webuntisLeistung.Zeugnisart;
        //                    neueLeistung.Zeugnistext = webuntisLeistung.Zeugnistext;
        //                    this.Add(neueLeistung);
        //                }
        //            }

        //            // Wenn keine automatische Zuordnung vorgenommen wurde:

        //            if (webuntisLeistung.Query == null)
        //            {
        //                Global.WriteLine("Das Webuntis-Fach **" + webuntisLeistung.Fach + "** (" +  webuntisLeistung.Klasse + "|" + webuntisLeistung.Lehrkraft + ") kann keinem Atlantis-Fach zugeordnet werden. Bitte manuell zuordnen:");

        //                var alleZulässigenAtlantisZielfächer = gespeicherteZuordnungen.AlleZulässigenAtlantisZielFächerAuflisten(webuntisLeistung, atlantisLeistungen, aktuellesSchuljahr, interessierendeKlasse);

        //                //Global.AufConsoleSchreiben("Wollen Sie die Zuordnung für **" + webuntisLeistung.Fach + "** manuell vornehmen?");

        //                bool wiederholen = true;

        //                int index = 0;

        //                string eingabe;

        //                do
        //                {
        //                    Console.Write(" Bitte Zahl (0 = keine Zuordnung) 1 bis " + alleZulässigenAtlantisZielfächer.Count + " eingeben: ");

        //                    try
        //                    {
        //                        eingabe = Console.ReadLine();
        //                        index = int.Parse(eingabe);

        //                        if (index >= 0 && index <= alleZulässigenAtlantisZielfächer.Count)
        //                        {
        //                            wiederholen = false;
        //                        }
        //                    }
        //                    catch (Exception)
        //                    {
        //                    }
        //                } while (wiederholen);

        //                if (index == 0)
        //                {
        //                    Global.WriteLine("Das Fach " + webuntisLeistung.Fach + " wird nicht zugeordnet und verworfen.");
        //                }
        //                else
        //                {
        //                    Global.WriteLine("Ihre Auswahl: " + webuntisLeistung.Fach + " wird " + alleZulässigenAtlantisZielfächer[index - 1] + " zugeordnet.");

        //                    DieZuordnungFürAlleSchülerVornehmen(webuntisLeistung, atlantisLeistungen, alleZulässigenAtlantisZielfächer, index, aktuellesSchuljahr);

        //                }
        //                kombinationAusKlasseUndFachWirdNurEinmalAufgerufen.Add(webuntisLeistung.Klasse + webuntisLeistung.Fach);
        //            }
        //        }
        //    }
                                    
        //    Global.WriteLine("Schülerleistungen aus Webuntis Atlantisleistungen erfolgreich zugeordnet ".PadRight(Global.PadRight, '.') + this.Count.ToString().PadLeft(4));
        //}

        //private void DieZuordnungFürAlleSchülerVornehmen(Leistung webuntisLeistung, Leistungen atlantisLeistungen, List<string> alleZulässigenAtlantisZielfächer, int index, string aktuellesSchuljahr)
        //{
        //    // Für diese Schüler wird die Zuordnung vorgenommen:

        //    var interessierendeSchüler = (from w in this
        //                                  where w.Klasse == webuntisLeistung.Klasse
        //                                  where w.Fach == w.Fach
        //                                  select w.SchlüsselExtern).Distinct().ToList();

        //    var ww = (from w in this
        //              where w.Fach == webuntisLeistung.Fach
        //              where interessierendeSchüler.Contains(w.SchlüsselExtern)
        //              select w).ToList();

        //    foreach (var w in ww)
        //    {
        //        var aL = (from a in atlantisLeistungen
        //                  where (a.Konferenzdatum >= DateTime.Now || a.Konferenzdatum.Year == 1) // eine Leistung des aktuellen Abschnitts
        //                  where a.Schuljahr == aktuellesSchuljahr
        //                  where a.Fach == alleZulässigenAtlantisZielfächer[index - 1]
        //                  where a.SchlüsselExtern == w.SchlüsselExtern
        //                  orderby a.LeistungId descending // Die höchste ID ist die ID des aktuellen Abschnitts
        //                  select a).ToList();

        //        if (aL.Count > 0)
        //        {
        //            if (index > 0)
        //            {
        //                w.AtlantisLeistungZuordnenUndQueryBauen(aL, webuntisLeistung.Fach + "->" + (aL.Count > 0 ? aL[0].Fach : "") + "|manuell zugeordnet|");
        //            }
        //        }
        //        else
        //        {
        //            if (index > 0)
        //            {
        //                w.Query = "/* " + w.Name + ": Das Fach " + webuntisLeistung.Fach + " wurde keinem Atlantisfach zugeordnet. Das kann passieren, wenn die Belegung in Atlantis von der in Webuntis abweicht.";
        //                w.Beschreibung = "";
        //            }
        //            else
        //            {
        //                w.Query = null;
        //                w.Beschreibung = null;
        //            }
        //        }
        //    }
        //}

        //internal void FehlendeSchülerInAtlantis(Leistungen atlantisLeistungen)
        //{
        //    List<string> keinSchülerInAtlantis = new List<string>();

        //    foreach (var we in this)
        //    {
        //        var atlantisFächerDiesesSchuelers = (from a in atlantisLeistungen
        //                                             where a.Klasse == we.Klasse
        //                                             where a.SchlüsselExtern == we.SchlüsselExtern
        //                                             select a).ToList();

        //        if (atlantisFächerDiesesSchuelers.Count == 0 && !keinSchülerInAtlantis.Contains(we.Name + "|" + we.Klasse))
        //        {
        //            keinSchülerInAtlantis.Add(we.Name + "|" + we.Klasse);
        //        }
        //    }
        //    if (keinSchülerInAtlantis.Count > 0)
        //    {
        //        AusgabeSchreiben("Achtung: Es gibt Leistungsdatensätze in Webuntis, die keinem Schüler in Atlantis zugeordnet werden können:", keinSchülerInAtlantis);
        //    }
        //}

        private string Gesamtpunkte2Tendenz(string gesamtpunkte)
        {
            string tendenz = "";

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
                return "";
            }
            return tendenz;
        }

        internal void FehlendeZeugnisbemerkungBeiStrich(Leistungen webuntisLeistungen, string interessierendeKlasse)
        {
            List<string> fehlendeBemerkungen = new List<string>();

            try
            {

                foreach (var a in (from t in this
                                   where t.Klasse == interessierendeKlasse
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
                                fehlendeBemerkungen.Add(w.Name + "(" + w.Klasse + "," + w.Fach + "," + w.Lehrkraft + ",Zeile: " + w.MarksPerLessonZeile + ")");
                            }
                        }
                    }
                }
                if (fehlendeBemerkungen.Count > 0)
                {
                    AusgabeSchreiben("Es gibt " + fehlendeBemerkungen.Count + " Schüler*innen in der " + interessierendeKlasse + ", die einen '-' in einem Nicht-Reli-Fach als Noten bekommen, ohne dass eine entsprechende Zeugnisbemerkung vorliegt:", fehlendeBemerkungen);
                }
                Global.WriteLine(("Fehlende Zeugnisbemerkung bei Strich ").PadRight(Global.PadRight, '.') + fehlendeBemerkungen.Count.ToString().PadLeft(4));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal Leistungen NotenVergangenerAbschnitteZiehen(List<Leistung> webuntisLeistungen, Leistungen geholteLeistungen, string klasse, List<string> aktSj, string hzJz)
        {   
            try
            {
                Leistungen leistungen = new Leistungen();

                List<string> leistungenDieserListe = new List<string>();

                Global.WriteLine("");

                if (geholteLeistungen.Count > 0)
                {
                    ConsoleKeyInfo x;

                    do
                    {                        
                        Global.WriteLine("Wollen Sie die vergangenen Leistungen aus den alten Abschnitten in Klasse " + klasse + " holen? (j/N)");
                        x = Console.ReadKey();
                    } while (x.Key.ToString().ToLower() != "j" && x.Key.ToString().ToLower() != "n" && x.Key.ToString() != "Enter");

                    // Wenn nur die Gesamtnote belegt ist, müssen Punkte und Tendenz hinzugefügt werden

                    geholteLeistungen.SetzeGesamtpunkte();

                    if (x.Key.ToString().ToLower() == "j")
                    {
                        Global.WriteLine(" Ihre Auswahl: J");
                        Global.WriteLine("");
                        return geholteLeistungen;
                    }
                    else
                    {
                        Global.WriteLine(" Ihre Auswahl: N");

                        // Beim Jahreszeugnis werden die Halbjahresnoten auf jeden Fall gezogen.

                        if (hzJz == "JZ")
                        {   
                            var l = new Leistungen();
                            l.AddRange((from g in geholteLeistungen 
                                        where g.Gesamtnote != null 
                                        where g.Gesamtnote != null      
                                        where g.Klasse.Substring(0,g.Klasse.Length - 1) == klasse.Substring(0,g.Klasse.Length - 1) 
                                        where g.Konferenzdatum.Year == DateTime.Now.Year select g).ToList());

                            string a = "";

                            foreach (var item in l)
                            {
                                if (!a.Contains(item.Fach + ","))
                                {
                                    a += item.Fach+",";
                                }                                
                            }

                            Global.WriteLine("   Da es sich um Jahreszeugnisse handelt, werden die Noten in " + a.TrimEnd(',') + " dennoch aus dem Halbjahr gezogen.");
                            Global.WriteLine("");
                            return l; 
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return new Leistungen();
        }

        private void SetzeGesamtpunkte()
        {
            foreach (var item in this)
            {
                if (item.Gesamtpunkte != null)
                {
                    item.Gesamtnote = Gesamtpunkte2Gesamtnote(item.Gesamtpunkte);
                    item.Tendenz = Gesamtpunkte2Tendenz(item.Gesamtpunkte);
                }
                if (item.Gesamtnote != null)
                {
                    item.Gesamtpunkte = Gesamtnote2Gesamtpunkte(item.Gesamtnote);
                    item.Tendenz = "";
                }
            }
        }

        private string Gesamtnote2Gesamtpunkte(string gesamtnote)
        {
            if (gesamtnote == "6")
            {
                return "0";
            }
            if (gesamtnote == "5")
            {
                return "2";
            }
            if (gesamtnote == "4")
            {
                return "5";
            }
            if (gesamtnote == "3")
            {
                return "8";
            }
            if (gesamtnote == "2")
            {
                return "11";
            }
            if (gesamtnote == "1")
            {
                return "14";
            }            
            if (gesamtnote == "A")
            {
                return "84";
            }
            if (gesamtnote == "-")
            {
                return "99";
            }
            return "";
        }

        public Leistungen(string connetionstringAtlantis, List<string> aktSj, string user, string interessierendeKlasse, Schülers schülers, string hzJz)
        {
            Global.Defizitleistungen = new Leistungen();

            Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum = new Leistungen();
            
            var leistungenUnsortiert = new Leistungen();

            // Es werden alle Atlantisleistungen für alle SuS gezogen.
            // Weil es denkbar ist, dass ein Schüler eine Leistung in einer anderen Klasse erbracht hat, werden 
            // für alle SuS alle Leistungen in diesem Bildungsgang gezogen.

            var abfrage = "";

                var klassenStammOhneJahrUndOhneZähler = Regex.Match(interessierendeKlasse, @"^[^0-9]*").Value;

                foreach (var schuelerId in (from s in schülers select s.SchlüsselExtern))
                {
                    abfrage += "(DBA.schueler.pu_id = " + schuelerId + @" AND  DBA.klasse.klasse like '" + klassenStammOhneJahrUndOhneZähler + @"%') OR ";
                }
            try 
            {
                abfrage = abfrage.Substring(0, abfrage.Length - 4);
            }
            catch (Exception ex)
            {
                Global.WriteLine("Kann es sein, dass für die Auswahl keine Datensätze vorliegen?\n" + ex.Message);
            }
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
DBA.noten_einzel.position_1 AS Reihenfolge,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.dat_geburt,
DBA.schueler.pu_id AS SchlüsselExtern,
DBA.schue_sj.s_religions_unterricht AS Religion,
DBA.schue_sj.dat_austritt AS ausgetreten,
DBA.schue_sj.dat_rel_abmeld AS DatumReligionAbmeldung,
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
                                    try
                                    {
                                        leistung.Reihenfolge = Convert.ToInt32(theRow["Reihenfolge"].ToString());
                                    }
                                    catch (Exception)
                                    {
                                        throw;
                                    }
                                    
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
                                    leistung.Tendenz = theRow["Tendenz"].ToString() == "" ? null : theRow["Tendenz"].ToString();
                                    leistung.EinheitNP = theRow["Einheit"].ToString() == "" ? "N" : theRow["Einheit"].ToString();
                                    leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString());
                                    leistung.HzJz = theRow["HzJz"].ToString();
                                    leistung.Anlage = theRow["Anlage"].ToString();
                                    leistung.Zeugnisart = theRow["Zeugnisart"].ToString();
                                    leistung.Zeugnistext = theRow["Zeugnistext"].ToString();
                                    leistung.Konferenzdatum = theRow["Konferenzdatum"].ToString().Length < 3 ? new DateTime() : (DateTime.ParseExact(theRow["Konferenzdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)).AddHours(15);
                                    leistung.DatumReligionAbmeldung = theRow["DatumReligionAbmeldung"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["DatumReligionAbmeldung"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                    leistung.SchuelerAktivInDieserKlasse = theRow["SchuelerAktivInDieserKlasse"].ToString() == "J";
                                    leistung.Abschlussklasse = leistung.IstAbschlussklasse();
                                    leistung.Beschreibung = "";                                    
                                    leistung.ReligionAbgewählt = leistung.HatReligionAbgewählt(this);
                                    leistung.GetFachAliases();
                                    //leistung.GeholteNote(this, alleWebuntisLeistungen);

                                    leistungenUnsortiert.Add(leistung);

                                    if (leistung.leistungDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum(aktSj))
                                    {
                                        Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum.Add(leistung);
                                    }
                         
                                    if (leistung.Gesamtnote == "5" || leistung.Gesamtnote == "6")
                                    {
                                        Global.Defizitleistungen.Add(leistung);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Global.WriteLine("Fehler beim Einlesen der Atlantis-Leistungsdatensätze: ENTER" + ex);
                                Console.ReadKey();
                            }
                        }
                    }
                    connection.Close();

                    // Sollen die Leistungen berücksichtigt werden, deren Konferenzdatum zwar in der Vergangenheit liegt, allerdings max seit 14 Tage 

                    leistungenUnsortiert.AddRange(LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum());
                    
                    this.AddRange(leistungenUnsortiert.OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name));

                    Global.SetzeReihenfolgeDerFächer(this);

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            var alleFächerDesAktuellenAbschnitts = (from l in this where (l.Konferenzdatum > DateTime.Now.AddDays(-20) || l.Konferenzdatum.Year == 1) select l.Fach).ToList();

            Global.WriteLine(("Leistungsdaten der SuS der Klasse " + interessierendeKlasse + " aus Atlantis (" + hzJz + ") ").PadRight(Global.PadRight, '.') + this.Count.ToString().PadLeft(4));
        }

        private Leistungen LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum()
        {
            if (Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum.Count > 0)
            {
                Console.WriteLine(" ");
                Global.WriteLine(("Es gibt Atlantisleistungen mit dem Konferenzdatum " + Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum.FirstOrDefault().Konferenzdatum.ToShortDateString()).PadRight(Global.PadRight,'.') + Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum.Count.ToString().PadLeft(4));
                Global.WriteLine(" Aus Sicherheitsgründen werden Atlantisleistungen mit zuückliegendem Konferenzdatum nicht ungefragt überschrieben. ");
                ConsoleKeyInfo x;

                do
                {
                    Global.Write("  Sollen die Leistungen mit Konferenzdatum " + Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum.FirstOrDefault().Konferenzdatum.ToShortDateString() + " in Atlantis überschrieben werden? (j/N)");
                    x = Console.ReadKey();
                    Console.WriteLine();

                    Global.WriteLine("   Ihre Auswahl: " + x.Key.ToString());
                    Console.WriteLine(" ");
                } while (x.Key.ToString().ToLower() != "j" && x.Key.ToString().ToLower() != "n" && x.Key.ToString() != "Enter");

                if (x.Key.ToString().ToLower() == "j")
                {
                    return Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum;
                }
            }
            return new Leistungen();
        }

        
        public Leistungen()
        {
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

            Global.WriteLine(("Neu anzulegende Leistungen in Atlantis ").PadRight(Global.PadRight, '.') + i.ToString().PadLeft(4));
            Global.PrintMessage(outputIndex, ("Neu anzulegende Leistungen in Atlantis: ").PadRight(65, '.') + (" " + i.ToString()).PadLeft(30, '.'));
        }

        public string Gesamtpunkte2Gesamtnote(string gesamtpunkte)
        {
            if (gesamtpunkte == "0")
            {
                return "6";
            }
            if (gesamtpunkte == "0.0")
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
            if (gesamtpunkte == "2.0")
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
            return "";
        }

        internal string Update(Leistungen atlantisLeistungen, bool debug)
        {
            if (atlantisLeistungen is null)
            {
                throw new ArgumentNullException(nameof(atlantisLeistungen));
            }

            int outputIndex = Global.SqlZeilen.Count();

            int i = 0;

            Global.WriteLine((" "));
            Global.WriteLine(("Leistungen in Atlantis updaten: "));
            Global.WriteLine(("=============================== "));
            Global.WriteLine((" "));


            //foreach (var w in (from t in this.OrderBy(x=>x.Klasse).ThenBy(x => x.Fach).ThenBy(x=>x.Name) where t.Beschreibung != null where t.Query!= null select t).ToList())
            foreach (var w in (from t in this.OrderBy(x => x.Klasse).ThenBy(x => x.Name).ThenBy(x => x.Fach) where t.Beschreibung != null where t.Query != null select t).ToList())
            {
                UpdateLeistung(w.Beschreibung, w.Query);
                i++;
            }

            if (i==0)
            {
                return "A C H T U N G: Es wurde keine einzige Leistung angelegt oder verändert. Das kann daran liegen, dass das Notenblatt in Atlantis nicht richtig angelegt ist. Die Tabelle zum manuellen eintragen der Noten muss in der Notensammelerfassung sichtbar sein. Außerdem muss ein Zeugnisdatum angelegt sein, das nicht in der Vergangenheit liegen darf.";
            }
            return "";
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