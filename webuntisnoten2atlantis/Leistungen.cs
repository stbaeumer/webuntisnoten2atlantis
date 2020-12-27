// Published under the terms of GPLv3 Stefan Bäumer 2020.

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;

namespace webuntisnoten2atlantis
{
    public class Leistungen : List<Leistung>
    {
        public Leistungen(string datei)
        {
            var leistungen = new Leistungen();

            using (StreamReader reader = new StreamReader(datei))
            {
                string überschrift = reader.ReadLine();

                Console.Write("Leistungsdaten aus Webuntis ".PadRight(71, '.'));

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
                            i++;

                            if (x.Length == 10)
                            {
                                leistung = new Leistung();
                                leistung.Datum = DateTime.ParseExact(x[0], "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                leistung.Name = x[1];
                                leistung.Klasse = x[2];
                                leistung.Fach = x[3];
                                leistung.Gesamtpunkte = x[9].Split('.')[0];
                                leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);
                                leistung.Tendenz = Gesamtpunkte2Tendenz(leistung.Gesamtpunkte);
                                leistung.Bemerkung = x[6];
                                leistung.Benutzer = x[7];
                                leistung.SchlüsselExtern = Convert.ToInt32(x[8]);
                                leistungen.Add(leistung);
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
                                Console.WriteLine("\n\n  [!] Achtung: In den Zeilen " + (i - 1) + "-" + i + " hat vermutlich die Lehrkraft eine Bemerkung mit einem Zeilen-");
                                Console.Write("      umbruch eingebaut. Es wird nun versucht trotzdem korrekt zu importieren ... ");
                            }

                            if (x.Length == 4)
                            {
                                leistung.Benutzer = x[1];
                                leistung.SchlüsselExtern = Convert.ToInt32(x[2]);
                                leistung.Gesamtpunkte = x[3].Split('.')[0];
                                leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);
                                Console.WriteLine("ok\n");
                                leistungen.Add(leistung);
                            }

                            if (x.Length < 4)
                            {
                                Console.WriteLine("\n\n[!] MarksPerLesson.csv: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.csv korrigiert werden.");
                                Console.ReadKey();
                                DateiÖffnen(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MarksPerLesson.csv");
                                throw new Exception("\n\n[!] MarksPerLesson.csv: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.csv korrigiert werden.");
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
                Console.WriteLine((" " + leistungen.Count.ToString()).PadLeft(30, '.'));
            }
            this.AddRange((from l in leistungen select l).OrderBy(x => x.Anlage).ThenBy(x => x.Klasse));
        }

        private string Gesamtpunkte2Tendenz(string gesamtpunkte)
        {
            string tendenz = "";

            if (gesamtpunkte == "0")
            {
                tendenz = null;
            }
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

        internal Leistungen HoleAlteNoten(Leistungen webuntisLeistungen, List<string> interessierendeKlassen, List<string> aktSj)
        {
            try
            {
                List<string> holeFächer = new List<string>();

                Leistungen leistungen = new Leistungen();

                string abschlussklassen = "";

                foreach (var klasse in interessierendeKlassen)
                {
                    // Für Klassen der Anlage A, die in diesem Schuljahr in Jahrgang 3 oder 4 sind ... 

                    if ((from a in this where klasse == a.Klasse where a.Anlage.StartsWith("A") where a.Abschlussklasse where a.Schuljahr == aktSj[0] + "/" + aktSj[1] select a).Any())
                    {
                        // ... werden die verschiedenen Schüler gesucht, die in diesem Schuljahr die Klasse besuchen

                        var schuelers = (from w in webuntisLeistungen where w.Klasse == klasse select w.SchlüsselExtern).Distinct().ToList();

                        // für jeden Schüler werden seine Noten der vergangenen Jahre gesucht

                        foreach (var schueler in schuelers)
                        {
                            var gliederungDesSchuelers = (from a in this where a.SchlüsselExtern == schueler where a.Schuljahr == aktSj[0] + "/" + aktSj[1] where a.Klasse == klasse select a.Gliederung).FirstOrDefault();

                            var vergangeneLeistungenDiesesSchuelers = (from a in this
                                                                       where a.SchlüsselExtern == schueler
                                                                       where a.Schuljahr != aktSj[0] + "/" + aktSj[1]
                                                                       where a.Klasse != null
                                                                       where a.Klasse.Substring(0, 1) == klasse.Substring(0, 1)
                                                                       where a.Gliederung == gliederungDesSchuelers
                                                                       where a.Gesamtnote != null
                                                                       where a.Gesamtnote != ""
                                                                       select a).OrderByDescending(x => x.Jahrgang).ToList();

                            var diesjähigeLeistungenDiesesSchuelers = (from a in this
                                                                       where a.SchlüsselExtern == schueler
                                                                       where a.Schuljahr == aktSj[0] + "/" + aktSj[1]
                                                                       where a.Klasse != null
                                                                       where a.Klasse.Substring(0, 1) == klasse.Substring(0, 1)
                                                                       where a.Gliederung == gliederungDesSchuelers
                                                                       select a).OrderByDescending(x => x.Jahrgang).ToList();

                            var alleVerschiedenenFächerDiesesSchuelers = (from a in this
                                                                          where a.Klasse != null
                                                                          where a.Klasse.Substring(0, 1) == klasse.Substring(0, 1)
                                                                          where a.Gliederung == gliederungDesSchuelers
                                                                          where a.Gesamtnote != null
                                                                          where a.Gesamtnote != ""
                                                                          where a.SchlüsselExtern == schueler
                                                                          select a.Fach).Distinct().ToList();

                            // ... für jede Leistung der vergangenen Jahre wird geprüft, ob im aktuellen Jahr eine Note in dem Fach existiert ...

                            foreach (var fach in alleVerschiedenenFächerDiesesSchuelers)
                            {
                                // ... wenn bisher keine Webuntisleistung in diesem Fach existiert, ...

                                if (!(from w in webuntisLeistungen where w.SchlüsselExtern == schueler where w.Fach == fach select w.Fach).Any() && vergangeneLeistungenDiesesSchuelers.Count > 0)
                                {
                                    // ... wird die jüngste Atlantisleistung mit Note in diesem Fach geholt ...

                                    Leistung vLeistung = (from v in vergangeneLeistungenDiesesSchuelers where v.Fach == fach select v).FirstOrDefault();

                                    if (!(from h in holeFächer where h.StartsWith(value: vLeistung.Fach + "(" + vLeistung.Klasse) select h).Any())
                                    {
                                        holeFächer.Add(vLeistung.Fach + "(" + vLeistung.Klasse + "|" + vLeistung.Schuljahr + ")");
                                    }

                                    Leistung aLeistung = (from v in diesjähigeLeistungenDiesesSchuelers where v.Fach == fach select v).FirstOrDefault();

                                    if (!abschlussklassen.Contains(aLeistung.Klasse))
                                    {
                                        abschlussklassen += aLeistung.Klasse + ",";
                                    }
                                    
                                    // Eine neue Webuntis-Leistung wird aus den geholten Angaben generiert.

                                    Leistung leistung = new Leistung();
                                    leistung.Abschlussklasse = true;
                                    leistung.Anlage = "A01";
                                    leistung.Beschreibung = "(" + vLeistung.Klasse + "|" + vLeistung.Schuljahr.Substring(2, 5) + ")";
                                    leistung.Fach = fach;
                                    leistung.GeholteNote = true;
                                    leistung.Gesamtnote = vLeistung.Gesamtnote;
                                    leistung.Gliederung = vLeistung.Gliederung;
                                    leistung.HzJz = vLeistung.HzJz;
                                    leistung.Jahrgang = aLeistung.Jahrgang;
                                    leistung.Klasse = aLeistung.Klasse;
                                    leistung.LeistungId = aLeistung.LeistungId;
                                    leistung.Name = aLeistung.Name;
                                    leistung.ReligionAbgewählt = aLeistung.ReligionAbgewählt;
                                    leistung.SchlüsselExtern = aLeistung.SchlüsselExtern;
                                    leistung.SchuelerAktivInDieserKlasse = aLeistung.SchuelerAktivInDieserKlasse;
                                    leistung.Schuljahr = aLeistung.Schuljahr;
                                    leistung.Zeugnistext = aLeistung.Zeugnistext;
                                        
                                    leistungen.Add(leistung);
                                }
                            }
                        }
                    }
                }

                if (abschlussklassen.Length > 0)
                {
                    int auswahl;

                    do
                    {
                        Console.WriteLine("   In Ihrer Klassenauswahl sind Abschlussklassen, für die alte Noten geholt werden können. ");
                        Console.WriteLine("   Sollen die alten Noten geholt werden?");
                        Console.WriteLine("");
                        Console.WriteLine("    1. Ja, für: " + abschlussklassen.TrimEnd(','));
                        Console.WriteLine("    2. Nein, keine alten Noten holen und bereits geholte wieder löschen.");
                        Console.WriteLine("");
                        Console.Write("     Ihre Auswahl " + (Properties.Settings.Default.AbschlussklassenAuswahl > 0 && Properties.Settings.Default.AbschlussklassenAuswahl < 3 ? "[ " + Properties.Settings.Default.AbschlussklassenAuswahl + " ] : " : ": "));

                        var key = Console.ReadKey();


                        if (int.TryParse(key.KeyChar.ToString(), out auswahl))
                        {
                            if (Convert.ToInt32(key.KeyChar.ToString()) > 0 && Convert.ToInt32(key.KeyChar.ToString()) < 3)
                            {
                                Properties.Settings.Default.AbschlussklassenAuswahl = Convert.ToInt32(key.KeyChar.ToString());
                                Properties.Settings.Default.Save();
                                auswahl = Convert.ToInt32(key.KeyChar.ToString());
                            }
                        }
                        if (key.Key == ConsoleKey.Enter)
                        {
                            auswahl = Properties.Settings.Default.AbschlussklassenAuswahl;
                        }

                        Console.WriteLine("  Sie haben '" + auswahl + "' gewählt ...");

                        if (auswahl < 1 || auswahl > 2)
                        {
                            Console.WriteLine("");
                            Console.Write("       ... Ungültige Auswahl! ");
                            Console.WriteLine("");
                        }
                        if (auswahl == 1)
                        {
                            AusgabeSchreiben("Noten aus Fächern aus vorherigen Schuljahren wurden geholt: ", holeFächer);
                        }
                        if (auswahl == 2)
                        {
                            AusgabeSchreiben("Es hätten Fächer aus vorherigen Schuljahren geholt werden können, sollten aber nicht.", new List<string>());
                            leistungen = new Leistungen();
                        }

                    } while (auswahl < 1 || auswahl > 2);
                }

                return leistungen;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }

        public Leistungen(string connetionstringAtlantis, List<string> aktSj)
        {
            Global.Output.Add("/* ****************************************************************************** */");
            Global.Output.Add("/* Diese Datei enthält alle Noten und Fehlzeiten aus Webuntis.Sie können alle No- */");
            Global.Output.Add("/* ten und Fehlzeiten aus Webuntis nach Atlantis importieren, indem Sie diese Da- */");
            Global.Output.Add("/* tei in Atlantis unter Funktionen>SQL-Anweisung hochladen.                      */");
            Global.Output.Add("/* Published under the terms of GPLv3. Hoping for the best!                       */");
            Global.Output.Add("/* " + (System.Security.Principal.WindowsIdentity.GetCurrent().Name + " " + DateTime.Now.ToString()).PadRight(78) + " */");
            Global.Output.Add("/* ****************************************************************************** */");
            Global.Output.Add("");

            try
            {
                var typ = (DateTime.Now.Month > 2 && DateTime.Now.Month <= 9) ? "JZ" : "HZ";

                Console.Write(("Leistungsdaten aus Atlantis (" + typ + ")").PadRight(71, '.'));

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
DBA.noten_einzel.s_tendenz AS Tendenz,
DBA.noten_einzel.s_einheit AS Einheit,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
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
DBA.noten_kopf.dat_notenkonferenz AS Konferenzdatum,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE s_art_fach <> 'U' AND schue_sj.s_typ_vorgang = 'A' AND (s_typ_nok = 'JZ' OR s_typ_nok = 'HZ') AND
(  
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 0) + "/" + (Convert.ToInt32(aktSj[1]) - 0) + @"')
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 1) + "/" + (Convert.ToInt32(aktSj[1]) - 1) + @"' AND (klasse.jahrgang = 'A011' OR klasse.jahrgang = 'A012' OR klasse.jahrgang = 'A013')  AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 2) + "/" + (Convert.ToInt32(aktSj[1]) - 2) + @"' AND (klasse.jahrgang = 'A011' OR klasse.jahrgang = 'A012') AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 3) + "/" + (Convert.ToInt32(aktSj[1]) - 3) + @"' AND (klasse.jahrgang = 'A011') AND s_note IS NOT NULL)
)
ORDER BY DBA.klasse.s_klasse_art ASC , DBA.klasse.klasse ASC; ", connection);


                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                    foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                    {
                        // Leistungen des aktuellen Abschnitts sind abhängig vom aktuellen Monat.
                        // Leistungen vergangener Jahre sind immer "JZ"

                        if (typ == theRow["HzJz"].ToString() || (theRow["Schuljahr"].ToString() != (Convert.ToInt32(aktSj[0]) - 0) + "/" + (Convert.ToInt32(aktSj[1]) - 0) && theRow["HzJz"].ToString() == "JZ"))
                        {
                            DateTime austrittsdatum = theRow["ausgetreten"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["ausgetreten"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                            
                            Leistung leistung = new Leistung();

                            try
                            {
                                // Wenn der Schüler nicht in diesem Schuljahr ausgetreten ist ...

                                if (!(austrittsdatum > new DateTime(DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1, 8, 1) && austrittsdatum < DateTime.Now))
                                {
                                    leistung.LeistungId = Convert.ToInt32(theRow["LeistungId"]);                         
                                    leistung.ReligionAbgewählt = theRow["Religion"].ToString() == "N";
                                    leistung.Schuljahr = theRow["Schuljahr"].ToString();
                                    leistung.Gliederung = theRow["Gliederung"].ToString();
                                    leistung.Jahrgang = Convert.ToInt32(theRow["Jahrgang"].ToString().Substring(3,1));
                                    leistung.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                                    leistung.Klasse = theRow["Klasse"].ToString();
                                    leistung.Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();
                                    leistung.Gesamtnote = theRow["Note"].ToString() == "Attest" ? "A" : theRow["Note"].ToString();                                    
                                    leistung.Gesamtpunkte = theRow["Punkte"].ToString();
                                    leistung.Tendenz = theRow["Tendenz"].ToString() == "" ? null : theRow["Tendenz"].ToString();
                                    leistung.EinheitNP = theRow["Einheit"].ToString() == "" ? "N" : theRow["Einheit"].ToString();
                                    leistung.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString());
                                    leistung.HzJz = theRow["HzJz"].ToString();
                                    leistung.Anlage = theRow["Anlage"].ToString();
                                    leistung.Zeugnistext = theRow["Zeugnistext"].ToString();
                                    leistung.Konferenzdatum = theRow["Konferenzdatum"].ToString().Length < 3 ? new DateTime() : (DateTime.ParseExact(theRow["Konferenzdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)).AddHours(15);
                                    leistung.SchuelerAktivInDieserKlasse = theRow["SchuelerAktivInDieserKlasse"].ToString() == "J";
                                    leistung.Abschlussklasse = leistung.IstAbschlussklasse();
                                    leistung.Beschreibung = "";
                                    leistung.GeholteNote = false;                                    
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Fehler beim Einlesen der Atlantis-Leistungsdatensätze: ENTER" + ex);
                                Console.ReadKey();
                            }
                                this.Add(leistung);                                                  
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

        internal void SprachenZuordnen(Leistungen atlantisLeistungen)
        {
            // Bei Sprachen-Fächern ...

            foreach (var we in (from t in this where (t.Fach.Split(' ')[0] == "EK" || t.Fach.Split(' ')[0] == "E" || t.Fach.Split(' ')[0] == "NL" || t.Fach.Split(' ')[0] == "LA" || t.Fach.Split(' ')[0] == "S" ) select t).ToList())
            {
                // ... wenn keine 1:1-Zuordnung möglich ist ...

                if (!(from a in atlantisLeistungen where a.Fach == we.Fach where a.Klasse == we.Klasse select a).Any())
                {
                    // wird versucht die Niveaustufe abzuschneiden

                    if (!(from a in atlantisLeistungen where a.Fach.Replace("A1","").Replace("A2", "").Replace("B1", "").Replace("B2", "").Replace("KA2", "") == we.Fach where a.Klasse == we.Klasse select a).Any())
                    {
                        // wird versucht den Kurs abzuschneiden

                        if ((we.Fach.Replace("  ", " ").Split(' ')).Count() > 1)
                        {                        
                            foreach (var a in atlantisLeistungen)
                            {
                                if (a.Klasse == we.Klasse)
                                {
                                    if (a.Fach.Split(' ')[0] == we.Fach.Split(' ')[0])
                                    {
                                        if ((a.Fach.Replace("  ", " ").Split(' ')).Count() > 1)
                                        {
                                            if (a.Fach.Replace("  ", " ").Split(' ')[1].Substring(0, 1) == we.Fach.Replace("  ", " ").Split(' ')[1].Substring(0, 1))
                                            {
                                                we.Fach = a.Fach;
                                            }
                                            else
                                            {
                                                Console.WriteLine("Das Fach " + we.Fach + " in der Klasse " + we.Klasse + " kann keinem Atlantis-Fach zugeordnet werden. ENTER");
                                                Console.ReadKey();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        we.Fach = (from a in atlantisLeistungen where a.Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "").Replace("KA2", "") == we.Fach where a.Klasse == we.Klasse select a.Fach).FirstOrDefault();
                    }
                }
                else
                {
                    we.Fach = (from a in atlantisLeistungen where a.Fach == we.Fach where a.Klasse == we.Klasse select a.Fach).FirstOrDefault();
                }
            }
        }
        
        internal void BindestrichfächerZuordnen(Leistungen atlantisLeistungen)
        {
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
                            x.Fach = fach;
                        }
                    }
                }
            }
        }

        internal void Religionsabwähler(Leistungen atlantisLeistungen)
        {
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
                            leistung.Benutzer = "";
                            leistung.SchlüsselExtern = aLeistung.SchlüsselExtern;
                            this.Add(leistung);
                        }
                        else
                        {
                            wf.Gesamtnote = "-";
                        }                        
                    }
                }
            }            
        }
        
        internal void ReligionKorrigieren()
        {
            foreach (var leistung in this)
            {
                if (leistung.Fach == "KR" || leistung.Fach == "ER" || leistung.Fach.StartsWith("KR ") || leistung.Fach.StartsWith("ER "))
                {
                    leistung.Fach = "REL";
                }
            }             
        }

        internal void WeitereFächerZuordnen(Leistungen atlantisLeistungen)
        {
            // Für die verschiednenen Klassen ...

            foreach (var klasse in (from k in this select k.Klasse).Distinct())
            {
                // ... und deren verschiedene Fächer aus Webuntis, in denen auch Noten gegeben wurden (also kein FU) ... 

                foreach (var webuntisFach in (from t in this where t.Klasse == klasse
                                              where t.Gesamtnote != ""
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
                            AusgabeSchreiben("Klasse: " + klasse + ". Das Webuntisfach " + webuntisFach + " wird dem Atlantisfach " + af.Fach + " zugeordnet. Wenn das nicht gewünscht ist, dann hier mit Suchen & Ersetzen der gewünschte Name gesetzt werden.",new List<string>());
                            

                            foreach (var leistung in this)
                            {
                                if (leistung.Fach == webuntisFach && leistung.Klasse == klasse)
                                {
                                    leistung.Fach = af.Fach;
                                }
                            }
                        }
                        else
                        {
                            AusgabeSchreiben("Achtung: In der " + klasse + " kann ein Fach in Webuntis keinem Atlantisfach zugeordnet werden:", new List<string>() { webuntisFach });                            
                        }
                    }
                }
            }
        }
        
        public Leistungen()
        {
        }

        internal void Add(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {
            int i = 0;

            Global.PrintMessage("Neu anzulegende Leistungen in Atlantis:");

            try
            {
                foreach (var klasse in interessierendeKlassen)
                {
                    foreach (var a in (from t in this where t.Klasse == klasse where t.SchuelerAktivInDieserKlasse select t).OrderBy(x=>x.Name).ToList())
                    {
                        var w = (from webuntisLeistung in webuntisLeistungen
                                 where webuntisLeistung.Fach == a.Fach
                                 where webuntisLeistung.Klasse == a.Klasse
                                 where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                 where a.Gesamtnote == ""
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null && w.Gesamtnote != "")
                        {
                            // Ein '-' in Religion wird nur dann gesetzt, wenn bereits andere Schüler eine Gesamtnote bekommen haben

                            if (!(w.Fach == "REL" && w.Gesamtnote == "-" && (from we in webuntisLeistungen 
                                                                             where we.Klasse == w.Klasse                                                                              
                                                                             where we.Fach == "REL" 
                                                                             where (we.Gesamtnote != "" && we.Gesamtnote != "-")
                                                                             select we).Count() == 0))
                            {
                                UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "|" + w.Gesamtnote + (w.Tendenz == null ? " " : w.Tendenz) + "|" + w.Fach + w.Beschreibung, "UPDATE noten_einzel SET    s_note='" + w.Gesamtnote + "' WHERE noe_id=" + a.LeistungId + ";");

                                if (a.EinheitNP == "P")
                                {
                                    // Wenn die Tendenzen abweichen ...

                                    if (w.Tendenz != a.Tendenz)
                                    {
                                        UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "| " + w.Tendenz + "|" + a.Fach + w.Beschreibung, "UPDATE noten_einzel SET s_tendenz=" + (w.Tendenz == null ? "NU" : "'" + w.Tendenz + "'") + " WHERE noe_id=" + a.LeistungId + ";"); ;
                                    }
                                        
                                    UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "|" + w.Gesamtpunkte.PadLeft(2) + "|" + w.Fach + w.Beschreibung, "UPDATE noten_einzel SET    punkte=" + w.Gesamtpunkte.PadLeft(2) + "  WHERE noe_id=" + a.LeistungId + ";");
                                }
                                i++;
                            }
                        }                       
                    }
                }
                
                if (i == 0)
                {
                    UpdateLeistung("                               ***keine***", "");   
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string Gesamtpunkte2Gesamtnote(string gesamtpunkte)
        {
            string gesamtnote = "";

            if (gesamtpunkte == "0")
            {
                gesamtnote = "6";
            }
            if (gesamtpunkte == "1")
            {
                gesamtnote = "5";
            }
            if (gesamtpunkte == "2")
            {
                gesamtnote = "5";
            }
            if (gesamtpunkte == "3")
            {
                gesamtnote = "5";
            }
            if (gesamtpunkte == "4")
            {
                gesamtnote = "4";
            }
            if (gesamtpunkte == "5")
            {
                gesamtnote = "4";
            }
            if (gesamtpunkte == "6")
            {
                gesamtnote = "4";
            }
            if (gesamtpunkte == "7")
            {
                gesamtnote = "3";
            }
            if (gesamtpunkte == "8")
            {
                gesamtnote = "3";
            }
            if (gesamtpunkte == "9")
            {
                gesamtnote = "3";
            }
            if (gesamtpunkte == "10")
            {
                gesamtnote = "2";
            }
            if (gesamtpunkte == "11")
            {
                gesamtnote = "2";
            }
            if (gesamtpunkte == "12")
            {
                gesamtnote = "2";
            }
            if (gesamtpunkte == "13")
            {
                gesamtnote = "1";
            }
            if (gesamtpunkte == "14")
            {
                gesamtnote = "1";
            }
            if (gesamtpunkte == "15")
            {
                gesamtnote = "1";
            }
            if (gesamtpunkte == "84")
            {
                gesamtnote = "A";
            }
            if (gesamtpunkte == "99")
            {
                gesamtnote = "-";
            }
            return gesamtnote;
        }

        internal void Update(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {
            int i = 0;

            Global.PrintMessage("Zu ändernde Noten in Atlantis:");

            try
            {
                foreach (var a in this)
                {
                    if (a.SchuelerAktivInDieserKlasse)
                    {
                        var w = (from webuntisLeistung in webuntisLeistungen
                                 where webuntisLeistung.Fach == a.Fach
                                 where webuntisLeistung.Klasse == a.Klasse
                                 where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                 where a.Gesamtnote != ""
                                 where a.Gesamtnote != webuntisLeistung.Gesamtnote
                                 where webuntisLeistung.Gesamtnote != ""
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null)
                        {
                            // Ein '-' in Religion wird deleted, wenn kein anderer Schüler eine Gesamtnote bekommen haben

                            if (!(w.Fach == "REL" && w.Gesamtnote == "-" && (from we in webuntisLeistungen
                                                                             where we.Klasse == w.Klasse
                                                                             where we.Fach == "REL"
                                                                             where (we.Gesamtnote != "" && we.Gesamtnote != "-")
                                                                             select we).Count() == 0))
                            {
                                UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "|" + a.Gesamtnote + (w.Tendenz == null ? " " : w.Tendenz) + ">" + w.Gesamtnote + (w.Tendenz == null ? " " : w.Tendenz) + "|" + a.Fach  + w.Beschreibung, "UPDATE noten_einzel SET    s_note='" + w.Gesamtnote + "'  WHERE noe_id=" + a.LeistungId + ";");

                                if (a.EinheitNP == "P")
                                {
                                    UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "|" + (a.Gesamtpunkte == null ? "" : a.Gesamtpunkte.Split(',')[0]).PadLeft(2) + ">" + (w.Gesamtpunkte == null ? " " : w.Gesamtpunkte.Split(',')[0]).PadLeft(2) + "|" + a.Fach + w.Beschreibung, "UPDATE noten_einzel SET    punkte=" + (w.Gesamtpunkte).PadLeft(2) + "   WHERE noe_id=" + a.LeistungId + ";");
                                    
                                    if (a.Tendenz != w.Tendenz)
                                    {
                                        UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "|" + (a.Tendenz == null ? "NULL" : a.Tendenz + " ") + ">" + (w.Tendenz == null ? "NU" : w.Tendenz + " ") + "|" + a.Fach + w.Beschreibung, "UPDATE noten_einzel SET s_tendenz=" + (w.Tendenz == null ? "NULL" : "'" + w.Tendenz + "'") + " WHERE noe_id=" + a.LeistungId + ";");
                                    }                                    
                                }

                                i++;
                            }   
                        }
                    }                    
                }
                if (i == 0)
                {
                    UpdateLeistung("                               ***keine***", "");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }

        internal void Delete(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen, List<string> aktSj)
        {
            Global.PrintMessage("Zu löschende Noten in Atlantis:");

            int i = 0;

            try
            {
                foreach (var a in this)
                {
                    if (a.SchuelerAktivInDieserKlasse)
                    {
                        // Wenn es zu einem Atlantis-Datensatz keine Entsprechung in Webuntis gibt, ... 
                        
                        var webuntisLeistung = (from w in webuntisLeistungen
                                                where w.Fach == a.Fach
                                                where w.Klasse == a.Klasse
                                                where w.SchlüsselExtern == a.SchlüsselExtern
                                                where w.Gesamtpunkte != ""
                                                select w).FirstOrDefault();

                        if (webuntisLeistung == null)
                        {
                            if (a.Gesamtnote != "")
                            {
                                // Geholte Noten aus Vorjahren werden nicht gelöscht.

                                if (!a.GeholteNote)
                                {
                                    UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "|" + a.Fach.Substring(0,Math.Min(a.Fach.Length, 2)).PadRight(2) + "|" + (a.Gesamtnote == null ? "NU" : a.Gesamtnote+ (a.Tendenz == null ? " " : a.Tendenz)) + ">NU" + a.Beschreibung, "UPDATE noten_einzel SET    s_note=NULL WHERE noe_id=" + a.LeistungId + ";");
                                    UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "|" + a.Fach.Substring(0,Math.Min(a.Fach.Length, 2)).PadRight(2) + "|" + ((a.Gesamtpunkte == null ? "NU" : a.Gesamtpunkte.Split(',')[0])).PadLeft(2) + ">NU" + a.Beschreibung, "UPDATE noten_einzel SET    punkte=NULL WHERE noe_id=" + a.LeistungId + ";");

                                    if (a.Tendenz != null)
                                    {
                                        UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 3)) + "|" + a.Fach.Substring(0,Math.Min(a.Fach.Length, 2)).PadRight(2) + "|" + (a.Tendenz == null ? "NU" : " " + a.Tendenz) + ">NU" + a.Beschreibung, "UPDATE noten_einzel SET s_tendenz=NULL WHERE noe_id=" + a.LeistungId + ";");
                                    }                                    

                                    i++;
                                }
                            }                            
                        }
                    }
                }
                if (i == 0)
                {
                    UpdateLeistung("                               ***keine***", "");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal List<string> GetIntessierendeKlassen(Leistungen alleWebuntisLeistungen, List<string> aktSj)
        {
            int auswahl;
            List<string> optionen = new List<string>() {
            @"Verarbeitungshinweise zeigen",
                "Alle Klassen         (in denen auch Gesamtnoten eingetragen sind",
                "Alle Vollzeitklassen (in denen auch Gesamtnoten eingetragen sind)",
                "Alle Teilzeitklassen (in denen auch Gesamtnoten eingetragen sind)",
                "Bestimmte Klassen    (in denen auch Gesamtnoten eingetragen sind)"
            };

            do
            {
                Console.WriteLine("");
                Console.WriteLine(" Auswahl:");
                Console.WriteLine("");
                for (int i = 0; i < optionen.Count; i++)
                {
                    Console.WriteLine(optionen.IndexOf(optionen[i]) + ". " + optionen[i]);
                }
                
                Console.WriteLine("");
                Console.Write(" Bitte wählen Sie " + (Properties.Settings.Default.Auswahl > 0 && Properties.Settings.Default.Auswahl < 5 ? "[ " + Properties.Settings.Default.Auswahl + " ] : " : ": "));
                
                var key = Console.ReadKey();

                if (int.TryParse(key.KeyChar.ToString(), out auswahl))
                {
                    if (Convert.ToInt32(key.KeyChar.ToString()) > 0 && Convert.ToInt32(key.KeyChar.ToString()) < 5)
                    {
                        Properties.Settings.Default.Auswahl = Convert.ToInt32(key.KeyChar.ToString());
                        Properties.Settings.Default.Save();
                        auswahl = Convert.ToInt32(key.KeyChar.ToString());                        
                    }
                }
                if (key.Key == ConsoleKey.Enter)
                {
                    auswahl = Properties.Settings.Default.Auswahl;

                    // Cursor an das Ende der Zeile setzen.

                    Console.SetCursorPosition(26, Console.CursorTop);
                }
                if (auswahl < 0 || auswahl > 4)
                {                    
                    Console.WriteLine("");
                    Console.Write(" ... Ungültige Auswahl! ");
                    Console.WriteLine("");
                }
                else
                {
                    Console.WriteLine("  Sie haben Nr. " + auswahl + " gewählt ...");
                }
                if (auswahl == 0)
                {
                    Console.WriteLine("");
                    Console.WriteLine(" *********************************************************************************************");
                    Console.WriteLine(" Verarbeitungshinweise");
                    Console.WriteLine(" =====================");
                    Console.WriteLine(" * Es wird automatisch (in Abhängigkeit vom aktuellen Datum) zwischen HZ und JZ entschieden.");
                    Console.WriteLine("   Zwischen Februar und September werden Jahreszeugnisse (JZ) verarbeitet.");
                    Console.WriteLine("   Ansonsten Halbjahreszeugnisse (HZ).");
                    
                    Console.WriteLine(" * Wenn einem Untis-Unterricht zwei Berufe zugeordnet sind und der Unterricht in Atlantis je");
                    Console.WriteLine("   nach Beruf zu unterschiedlichen Fächern führt, dann sollte der Untis-Unterricht als Binde- ");
                    Console.WriteLine("   strichfach angelegt werden. Beispiel: Untisfach ABC-DEF wird in Atlantis zu ABC oder DEF.");
                    
                    Console.WriteLine(" * ER und KR oder auch ER G1 usw. werden in Atlantis dem Fach REL zugeordnet");
                    Console.WriteLine(" * Wenn ein Schüler Religion abgewählt hat, dann wird seine Note auf '-' gesetzt");
                    Console.WriteLine(" * Noten in der Gym werden als Note und als Punkte gesetzt.");
                    Console.WriteLine(" * Wenn bei Sprachen keine 1:1-Zuordnung möglich ist, dann wird versucht die Niveautufe");
                    Console.WriteLine("   abzuschneiden.");
                    Console.WriteLine(" *********************************************************************************************");
                    Console.WriteLine("");
                }
            } while (auswahl < 1 || auswahl > 4);

            List<string> alleVerschiedenenUntisKlassenMitNoten = (from k in alleWebuntisLeistungen where k.Gesamtnote != "" where k.Klasse != null select k.Klasse).Distinct().ToList();
            List<string> alleVerschiedenenAtlantisKlassen = (from k in this select k.Klasse).Distinct().ToList();
            List<string> interessierendeKlassen = new List<string>();
                        
            foreach (var klasse in alleVerschiedenenUntisKlassenMitNoten)
            {
                if (auswahl == 1)
                {
                    interessierendeKlassen.Add(klasse);
                }
                if (auswahl == 2)
                {
                    if ((from a in this where a.Klasse == klasse where !a.Anlage.StartsWith("A") select a).Any())
                    {
                        interessierendeKlassen.Add(klasse);
                    }
                }
                if (auswahl == 3)
                {
                    if ((from a in this where a.Klasse == klasse where a.Anlage.StartsWith("A") select a).Any())
                    {
                        interessierendeKlassen.Add(klasse);
                    }
                }
            }

            if (auswahl == 4)
            {
                interessierendeKlassen.AddRange(KlassenAuswählen(alleVerschiedenenUntisKlassenMitNoten));
            }

            interessierendeKlassen = ZeugnisdatumPrüfen(interessierendeKlassen);

            AusgabeSchreiben("Ausgewertete Klassen: " + optionen[auswahl] + ":", interessierendeKlassen);

            return interessierendeKlassen;
        }

        private List<string> ZeugnisdatumPrüfen(List<string> interessierendeKlassen)
        {
            List<string> interessierendeKlassenGefiltert = new List<string>();

            List<string> klassenMitZurückliegenderZeugniskonferenz = new List<string>();

            foreach (var klasse in interessierendeKlassen)
            {
                foreach (var a in (from t in this where t.Klasse == klasse select t).ToList())
                {
                    // Wenn die Notenkonferenz -je nach Halbjahr ...

                    if (a != null && a.HzJz == "HZ")
                    {
                        // ... in der Vergangenheit liegen
                        if (a.Konferenzdatum > new DateTime((DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1), 08, 01) && a.Konferenzdatum < DateTime.Now)
                        {
                            // ... wird die Leistung nicht berücksichtigt

                            if (!(from k in klassenMitZurückliegenderZeugniskonferenz where k.StartsWith(klasse) select k).Any())
                            {
                                klassenMitZurückliegenderZeugniskonferenz.Add(klasse + "(" + a.Konferenzdatum.ToShortDateString() + ")");
                            }
                        }
                        else
                        {
                            if (!(from k in interessierendeKlassenGefiltert where k == klasse select k).Any())
                            {
                                interessierendeKlassenGefiltert.Add(klasse);
                            }                            
                        }
                    }

                    if (a != null && a.HzJz == "JZ")
                    {
                        // ... in der Vergangenheit liegen
                        if (a.Konferenzdatum > new DateTime((DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 : DateTime.Now.Year), 02, 01) && a.Konferenzdatum < DateTime.Now)
                        {
                            // ... wird die Leistung nicht berücksichtigt
                            if (!(from k in klassenMitZurückliegenderZeugniskonferenz where k.StartsWith(klasse) select k).Any())
                            {
                                klassenMitZurückliegenderZeugniskonferenz.Add(klasse+"("+a.Konferenzdatum.ToShortDateString()+")");
                            }
                        }
                        else
                        {
                            interessierendeKlassenGefiltert.Add(klasse);
                        }
                    }
                }
            }
                        
            if (klassenMitZurückliegenderZeugniskonferenz.Count > 0)
            {
                AusgabeSchreiben("Folgende Klassen passen zu Ihrem Suchmuster, werden aber wegen der zurückliegenden Notenkonferenz nicht angefasst:", klassenMitZurückliegenderZeugniskonferenz);
            }
            return interessierendeKlassenGefiltert;
        }

        private List<string> KlassenAuswählen(List<string> alleVerschiedenenUntisKlassenMitNoten)
        {
            List<string> klassenliste = new List<string>();
            string klassenlisteString = "";
            string eingabestring = "";
            string eingaben = "";

            do
            {             
                Console.WriteLine("  Geben Sie einen oder mehrere Klassennamen oder auch nur Teile des Namens ein.");
                Console.WriteLine("");
                Console.Write("  Wählen Sie" + (Properties.Settings.Default.Klassenwahl == "" ? " : " : " [ " + Properties.Settings.Default.Klassenwahl + " ] : "));

                eingaben = Console.ReadLine().ToUpper();

                Console.SetCursorPosition(0, Console.CursorTop - 1);
                                
                if (eingaben == "")
                {
                    eingaben = Properties.Settings.Default.Klassenwahl;
                }

                Console.WriteLine("   Sie haben '" + eingaben + "' gewählt ...");

                foreach (var eingabe in eingaben.ToUpper().Split(','))
                {
                    foreach (var klasse in alleVerschiedenenUntisKlassenMitNoten)
                    {
                        if (klasse.StartsWith(eingabe) || (eingabe.Length > 1 && klasse.Contains(eingabe)))
                        {                            
                            klassenliste.Add(klasse);
                            klassenlisteString += klasse + ",";
                            if (!eingabestring.Contains("," +eingabe) && !eingabestring.StartsWith(eingabe))
                            {
                                eingabestring += eingabe + ",";
                            }                            
                        }
                    }                    
                }
                
                if (klassenliste.Count == 0)
                {
                    Console.WriteLine("");
                    Console.WriteLine("      [!] Ihr Suchkriterium '" + eingaben + "' passt auf keine einzige Klasse.");
                    Console.WriteLine("          Beachten Sie, dass nur Klassen gewählt werden können, für die Gesamtnoten existieren.");
                    Console.WriteLine("          Gültige Eingaben können sein: '" + alleVerschiedenenUntisKlassenMitNoten[0] + "' " + (alleVerschiedenenUntisKlassenMitNoten.Count > 1 ? "oder '" + alleVerschiedenenUntisKlassenMitNoten[0] + "," + alleVerschiedenenUntisKlassenMitNoten[1] : "") +"' oder '" + alleVerschiedenenUntisKlassenMitNoten[0].Substring(0,1) + ",G'");
                    Console.WriteLine("          Wählbare Klassen sind:");
                    WählbareKlassen(alleVerschiedenenUntisKlassenMitNoten);
                }
            } while (klassenliste.Count == 0);

            Properties.Settings.Default.Klassenwahl = eingabestring.TrimEnd(',');
            Properties.Settings.Default.Save();

            return klassenliste;
        }

        private void WählbareKlassen(List<string> alleVerschiedenenUntisKlassenMitNoten)
        {
            int z = 0;

            do
            {
                var zeile = "";

                try
                {
                    while ((zeile + alleVerschiedenenUntisKlassenMitNoten[z] + ", ").Length <= 78)
                    {
                        zeile += alleVerschiedenenUntisKlassenMitNoten[z] + ", ";
                        z++;
                    }
                }
                catch (Exception)
                {
                    z++;
                    zeile.TrimEnd(',');
                }

                zeile = "          " + zeile.TrimEnd(' ');

            } while (z < alleVerschiedenenUntisKlassenMitNoten.Count);
        }

        private void AusgabeSchreiben(string text, List<string> klassen)
        {
            Global.Output.Add("");
            int z = 0;

            do
            {
                var zeile = "";
                
                try
                {
                    while ((zeile + text.Split(' ')[z] + ", ").Length <= 78)
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
                Global.Output.Add((o.Substring(0, Math.Min(82, o.Length))).PadRight(82) + "*/");

            } while (z < text.Split(' ').Count());



            z = 0;
            
            do
            {
                var zeile = " ";                

                try
                {
                    while ((zeile + klassen[z] + ", ").Length <= 78)
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

                string o = "/* " + zeile.TrimEnd(',');
                Global.Output.Add((o.Substring(0, Math.Min(82, o.Length))).PadRight(82) + "*/");

            } while (z < klassen.Count);
        }

        private void UpdateLeistung(string message, string updateQuery)
        {
            try
            {
                string o = updateQuery + "/*" + message;
                Global.Output.Add((o.Substring(0, Math.Min(81, o.Length))).PadRight(81) + " */");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal void ErzeugeSqlDatei(string outputSql)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(outputSql, true, Encoding.Default))
                {
                    foreach (var o in Global.Output)
                    {
                        writer.WriteLine(o);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            DateiÖffnen(outputSql);
        }

        private void DateiÖffnen(string pfad)
        {
            try
            {
                System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Notepad++\Notepad++.exe", pfad);
            }
            catch (Exception)
            {
                System.Diagnostics.Process.Start("Notepad.exe", pfad);
            }
        }
    }
}