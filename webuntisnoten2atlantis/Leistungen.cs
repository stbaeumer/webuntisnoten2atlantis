// Published under the terms of GPLv3 Stefan Bäumer 2023.

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
                                leistung.Lehrkraft = x[7].Length > 3 && x[7].ToLower().EndsWith("sa") ? x[7].ToUpper().Substring(0, x[7].Length - 2) : x[7].ToUpper();
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

            Console.WriteLine(("Alle Webuntis-Leistungen (mit & ohne Gesamtnote; ohne Dopplungen) ").PadRight(Global.PadRight - 2, '.') + this.Count.ToString().PadLeft(6));
        }

        internal List<string> GetMöglicheKlassen()
        {
            var x = (from m in this
                     where m.Klasse != ""
                     where m.Gesamtnote != null
                     where m.Gesamtnote != "" // nur Klassen, für die schon Noten gegeben wurden                     
                     select m.Klasse).Distinct().ToList();

            Console.WriteLine(("Webuntis-Klassen mit eingetragenen Gesamtnoten").PadRight(Global.PadRight - 2, '.') + x.Count.ToString().PadLeft(6));

            Console.WriteLine(Global.List2String(x, ","));

            return x;
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
                                    leistung.NokId = Convert.ToInt32(theRow["NOK_ID"]);
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

            Console.WriteLine(("Leistungsdaten der SuS der Klasse " + interessierendeKlasse + " aus Atlantis (" + hzJz + ") ").PadRight(Global.PadRight, '.') + this.Count.ToString().PadLeft(4));
        }

        /// <summary>
        /// Wenn kein Notenblatt angelegt ist, wird die Verarbeitung abgebrochen.
        /// </summary>
        /// <param name="hzJz"></param>
        /// <param name="aktSj"></param>
        /// <param name="interessierendeKlasse"></param>
        /// <exception cref="Exception"></exception>
        internal void IstNotenblattAngelegt(string hzJz , List<string> aktSj, string interessierendeKlasse)
        {
            string meldung = "";
            var angelegt = (from atlantisLeistung in this
                            where atlantisLeistung.Konferenzdatum.Date >= DateTime.Now.Date || atlantisLeistung.Konferenzdatum.Year == 1
                            where hzJz == atlantisLeistung.HzJz
                            where aktSj[0] + "/" + aktSj[1] == atlantisLeistung.Schuljahr
                            select atlantisLeistung).ToList();

            if (angelegt.Count > 0)
            {
                Console.WriteLine("");
                meldung += "Das Notenblatt für die Klasse " + interessierendeKlasse + " wurde angelegt.";
                
                if ((from atlantisLeistung in angelegt where atlantisLeistung.Konferenzdatum.Year == 1 select atlantisLeistung).Any())
                {
                    meldung += " Allerdings ist kein Konferenzdatum gesetzt. Vorsichtshalber sollte ein Konferenzdatum gesetzt werden.";
                }
                else
                {
                    meldung+="Konferenzdatum: " + angelegt[0].Konferenzdatum + ".";
                }
                Console.WriteLine(meldung);    
                Console.WriteLine("");
            }
            else
            {                
                throw new Exception("Es ist kein Notenblatt angelegt. Die Verarbeitung endet hier. Legen Sie zuerst ein Notenblatt an:\n" +
                    "1. Klassenverwaltung - > Klasse wählen " + interessierendeKlasse + "\n" +
                    "2. Reiter *Zeugnisdaten* klicken\n" +
                    "3. Unter alle Zeugnisdaten Halbjahreszeugnis oder Jahresendzeugnisse wählen.\n" +
                    "4. Konferenzdatum und Ausgabedatum wählen.\n" +
                    "5. *Notenblatt und Zeugnissatz () aller SuS mit diesen Zeugnisdaten aktualisieren\n" +
                    "6. Webuntisnoten2Atlantis erneut starten.");
            }
        }

        internal void AddLeistungen(List<Leistung> leistungs)
        {
            foreach (var leistung in leistungs)
            {
                this.Add(new Leistung(leistung.Fach,leistung.Gesamtnote, leistung.Konferenzdatum));
            }
        }

        private Leistungen LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum()
        {
            if (Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum.Count > 0)
            {
                Console.WriteLine(" ");
                Console.WriteLine(("Es gibt Atlantisleistungen mit dem Konferenzdatum " + Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum.FirstOrDefault().Konferenzdatum.ToShortDateString()).PadRight(Global.PadRight,'.') + Global.LeistungenDesAktuellenAbschnittsMitZurückliegendemKonferenzdatum.Count.ToString().PadLeft(4));
                Console.WriteLine(" Aus Sicherheitsgründen werden Atlantisleistungen mit zuückliegendem Konferenzdatum nicht ungefragt überschrieben. ");
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

        internal List<Leistung> InfragekommendeLeistungenHinzufügen(int schlüsselExtern, string hzJz, List<string> aktSj)
        {
            var infragekommendeLeistungenA = new Leistungen();

            foreach (var iA in (from al in this.OrderByDescending(x => x.Konferenzdatum)
                                where al.SchlüsselExtern == schlüsselExtern
                                where al.HzJz == hzJz
                                where al.Schuljahr == aktSj[0] + "/" + aktSj[1]
                                where al.Konferenzdatum.Date > DateTime.Now.Date || al.Konferenzdatum.Year == 1
                                select al).ToList())
            {
                // Nur Fächer, die nicht bereits erfolgreich zugeordnet wurden, werden als Infragekommende Fächer
                // für eine spätere manuelle Auswahl erfasst.

                if (!(from a in Global.AlleVerschiedenenUnterrichteInDieserKlasseAktuellUnsortiert
                      where a.FachnameAtlantis == iA.Fach
                      select a).Any())
                {
                    infragekommendeLeistungenA.Add(iA);
                }
            }
            return infragekommendeLeistungenA;
        }
    }
}