// Published under the terms of GPLv3 Stefan Bäumer 2021.

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
        public Leistungen(string targetMarksPerLesson)
        {
            var leistungen = new Leistungen();

            using (StreamReader reader = new StreamReader(targetMarksPerLesson))
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
                                leistung.Gesamtpunkte = x[9].Split('.')[0] == "" ? null : x[9].Split('.')[0];
                                leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);
                                leistung.Tendenz = Gesamtpunkte2Tendenz(leistung.Gesamtpunkte);
                                leistung.Bemerkung = x[6];
                                leistung.Lehrkraft = x[7];
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
                                Console.WriteLine("[!] Achtung: In den Zeilen " + (i - 1) + "-" + i + " hat vermutlich die Lehrkraft eine Bemerkung mit einem Zeilen-");
                                Console.Write("      umbruch eingebaut. Es wird nun versucht trotzdem korrekt zu importieren ... ");
                            }

                            if (x.Length == 4)
                            {
                                leistung.Lehrkraft = x[1];
                                leistung.SchlüsselExtern = Convert.ToInt32(x[2]);
                                leistung.Gesamtpunkte = x[3].Split('.')[0];
                                leistung.Gesamtnote = Gesamtpunkte2Gesamtnote(leistung.Gesamtpunkte);
                                Console.WriteLine("hat geklappt.\n");
                                leistungen.Add(leistung);
                            }

                            if (x.Length < 4)
                            {
                                Console.WriteLine("\n\n[!] MarksPerLesson.CSV: In der Zeile " + i + " stimmt die Anzahl der Spalten nicht. Das kann passieren, wenn z. B. die Lehrkraft bei einer Bemerkung einen Umbruch eingibt. Mit Suchen & Ersetzen kann die Datei MarksPerLesson.CSV korrigiert werden.");
                                Console.ReadKey();
                                OpenFiles(new List<string>() { targetMarksPerLesson });
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
                Console.WriteLine((" " + leistungen.Count.ToString()).PadLeft(30, '.'));
                Global.PrintMessage(Global.Output.Count(), "Webuntisleistungen: ".PadRight(45, '.') + (" " + leistungen.Count.ToString()).PadLeft(45, '.'));
            }
            this.AddRange((from l in leistungen select l).OrderBy(x => x.Anlage).ThenBy(x => x.Klasse));
        }

        internal void WidersprechendeGesamtnotenKorrigieren(List<string> interessierendeKlassen)
        {
            List<string> liste = new List<string>();

            foreach (var ik in interessierendeKlassen)
            {
                var verschiedeneFächerDerKlasse = (from t in this where t.Klasse == ik select t.Fach).Distinct();

                // Prüfe, ob es verschiende Noten für das selbe Fach gibt.

                foreach (var fach in verschiedeneFächerDerKlasse)
                {
                    var verschiedeneSchülerMitDiesemFachInDieserKlasse = (from t in this where t.Klasse == ik where t.Fach == fach select t.Name).Distinct();

                    foreach (var v in verschiedeneSchülerMitDiesemFachInDieserKlasse)
                    {   
                        var anzahl = (from t in this 
                                      where t.Klasse == ik 
                                      where t.Fach == fach 
                                      where t.Name == v 
                                      where t.Gesamtnote != null
                                      select t.Gesamtnote).Distinct().Count();

                        if (anzahl > 1)
                        {
                            var xxx = (from t in this
                                       where t.Klasse == ik
                                       where t.Fach == fach
                                       where t.Name == v
                                       where t.Gesamtnote != null                                       
                                       select t).OrderByDescending(y=>y.Datum).ToList();

                            for (int i = 0; i < xxx.Count; i++)
                            {
                                if (!(from l in liste where l == xxx[i].Klasse + "|" + xxx[i].Fach + "|" + xxx[i].Name select l).Any())
                                {
                                    liste.Add(xxx[i].Klasse + "|" + xxx[i].Fach + "|" + xxx[i].Name);                                    
                                }
                                xxx[i].Gesamtnote = xxx[0].Gesamtnote;
                                xxx[i].Gesamtpunkte = xxx[0].Gesamtpunkte;
                                xxx[i].Tendenz = xxx[0].Tendenz;
                            }                            
                        }
                    }
                }                
            }
            if (liste.Count() > 0)
            {
                AusgabeSchreiben("Achtung: Es gibt Fälle, in denen derselbe Schüler im selben Fach unterschiedliche Gesamtnoten bekommen hat. Evtl. ist in Untis das Fach 2x angelegt worden. Evtl. teilen sich mehrere Lehrkräfte ein Fach und haben unterschiedliche Noten eingetragen Es wird die zuletzt eingetragene Note gewählt:", liste);
            }            
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
            string aaaa = "";
        }

        internal Zuordnungen FächerZuordnen(Leistungen atlantisLeistungen)
        {
            Console.Write(("Fächer zuordnen ").PadRight(71, '.'));
            int i = 0;

            Zuordnungen zuordnungen = new Zuordnungen();

            for (int y = 0; y < this.Count; y++)            
            {   
                // Fächer, die sich nicht zuordnen lassen, werden später zur manuellen Zuordnung vorgeschlagen
                
                var zugeordnet = false;

                // Wenn es keine Zuordnung in Atlantis gibt, ...

                if (!(from a in atlantisLeistungen where a.Klasse == this[y].Klasse where a.Fach == this[y].Fach select a).Any())
                {
                    // Förderunterricht bekommt keine Note

                    if (this[y].Fach.EndsWith(" FU"))
                    {
                        zugeordnet = true;
                    }

                    if (!zugeordnet && (this[y].Fach == "KR" || this[y].Fach == "ER" || this[y].Fach.StartsWith("KR ") || this[y].Fach.StartsWith("ER ")))
                    {
                        this[y].Beschreibung += this[y].Fach + "->REL,";
                        this[y].Fach = "REL";
                        zugeordnet = true;
                        i++;
                    }

                    // Kurse mit Leerzeichen
                                        
                    if (!zugeordnet && (from a in atlantisLeistungen
                                        where a.Klasse == this[y].Klasse
                                        where this[y].Fach.Contains(" ") && a.Fach.Contains(" ") // Leerzeichen
                                        where this[y].Fach.Split(' ')[0] == a.Fach.Split(' ')[0] // Erster Teil identisch                                        
                                        where this[y].Fach.Substring(this[y].Fach.LastIndexOf(' ') + 1).Substring(0, 1) == a.Fach.Substring(a.Fach.LastIndexOf(' ') + 1).Substring(0, 1) // Zweiter Teil identisch
                                        select a).Any())
                    {

                        // Wenn in Atlantis die Anzahl der Zielfächer größer als 1 ist, dann wird die 

                        var zielfaecher = (from a in atlantisLeistungen
                                        where a.Klasse == this[y].Klasse
                                        where this[y].Fach.Contains(" ") && a.Fach.Contains(" ") && this[y].Fach.Split(' ')[0] == a.Fach.Split(' ')[0]
                                        where this[y].Fach.Substring(this[y].Fach.LastIndexOf(' ') + 1).Substring(0, 1) == a.Fach.Substring(a.Fach.LastIndexOf(' ') + 1).Substring(0, 1) // Zweiter Teil identisch
                                        select a.Fach).Distinct().ToList();

                        this[y].Beschreibung += this[y].Fach + "->" + zielfaecher[0] + ",";
                        this[y].Fach = zielfaecher[0];
                        zugeordnet = true;
                        i++;

                        if (zielfaecher.Count > 1)
                        {
                            Leistung wel = new Leistung();

                            wel.Abschlussklasse = this[y].Abschlussklasse;
                            wel.Anlage = this[y].Anlage;
                            wel.Bemerkung = this[y].Bemerkung;
                            wel.Beschreibung = this[y].Fach + "->" + zielfaecher[1] + ",";
                            wel.Datum = this[y].Datum;
                            wel.EinheitNP = this[y].EinheitNP;
                            wel.Fach = zielfaecher[1];
                            wel.GeholteNote = this[y].GeholteNote;
                            wel.Gesamtnote = this[y].Gesamtnote;
                            wel.Gesamtpunkte = this[y].Gesamtpunkte;
                            wel.Gliederung = this[y].Gliederung;
                            wel.HatBemerkung = this[y].HatBemerkung;
                            wel.HzJz = this[y].HzJz;
                            wel.Jahrgang = this[y].Jahrgang;
                            wel.Klasse = this[y].Klasse;
                            wel.Konferenzdatum = this[y].Konferenzdatum;
                            wel.Lehrkraft = this[y].Lehrkraft;
                            wel.LeistungId = this[y].LeistungId;
                            wel.Name = this[y].Name;
                            wel.ReligionAbgewählt = this[y].ReligionAbgewählt;
                            wel.SchlüsselExtern = this[y].SchlüsselExtern;
                            wel.SchuelerAktivInDieserKlasse = this[y].SchuelerAktivInDieserKlasse;
                            wel.Schuljahr = this[y].Schuljahr;
                            wel.Tendenz = this[y].Tendenz;
                            wel.Zeugnisart = this[y].Zeugnisart;
                            wel.Zeugnistext = this[y].Zeugnistext;
                            this.Add(wel);
                        }
                    }

                    // EK -> E

                    if (!zugeordnet && this[y].Fach.StartsWith("EK") && this[y].Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "E")
                    {
                        this[y].Beschreibung += this[y].Fach + "->" + this[y].Fach + ",";
                        zugeordnet = true;
                        i++;
                    }

                    // N -> NKA1

                    if (!zugeordnet && this[y].Fach.StartsWith("N") && this[y].Fach.Replace("A1", "").Replace("A2", "").Replace("B1", "").Replace("B2", "") == "NK")
                    {
                        this[y].Beschreibung += this[y].Fach + "->" + this[y].Fach + ",";
                        zugeordnet = true;
                        i++;
                    }

                    if (!zugeordnet && !(from al in atlantisLeistungen where al.Klasse == this[y].Klasse where al.Fach == this[y].Fach select al).Any())
                    {
                        // Dann wird die Zuordnung hinzugefügt, falls noch nicht geschehen

                        if (!(from z in zuordnungen where z.Quellklasse == this[y].Klasse where z.Quellfach == this[y].Fach select z).Any())
                        {
                            if (this[y].Fach != "")
                            {
                                zuordnungen.Add(new Zuordnung(this[y].Klasse, this[y].Fach));
                            }                            
                        }
                    }
                }









                //// IF -> WI

                //if (this[j].Fach == "IF" && t.Fach == "WI")
                //{
                //    this[j].Beschreibung += this[j].Fach + "->" + t.Fach + ",";
                //    zugeordnet = true;
                //    i++;
                //    break;
                //}

                //// ... ob es es die Klasse-Fach-Kombination in Atlantis nicht gibt.


            }

                //if (!zugeordnet && atlantisFächerDiesesSchuelers.Count > 0 && this[j].Gesamtnote != null)
                //{
                //    if (!(from z in zuordnungen where z.Quellklasse == this[j].Klasse where z.Quellfach == this[j].Fach select z).Any())
                //    {
                //        zuordnungen.Add(new Zuordnung(this[j].Klasse, this[j].Fach));
                //    }
                //}
            

            Console.WriteLine((" " + i.ToString()).PadLeft(30, '.'));

            return zuordnungen;
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

        internal void GetKlassenMitFehlendenZeugnisnoten(List<string> interessierendeKlassen, Leistungen alleWebuntisLeistungen)
        {
            var kl = new List<string>();

            // Alle interessierenden Klassen werden durchlaufen ...

            Console.Write(("Klassen mit fehlenden Zeugnisnoten").PadRight(71, '.'));

            foreach (var iKlasse in interessierendeKlassen)
            {
                // ... wenn in einer interessierenden Klasse mehr als 3 Fächer bereits Noten bekommen haben, wird angenommen, dass
                //     eine Zeugnisdruck ansteht, dann ...

                if ((from k in alleWebuntisLeistungen where k.Klasse == iKlasse where k.Gesamtnote != null select k.Fach).Distinct().Count() > 3)
                {
                    // ... werden diejenigen Fächer ermittelt, in denen alle Noten fehlen ...

                    var kll = (from k in alleWebuntisLeistungen where k.Klasse == iKlasse where k.Gesamtnote == null select new { k.Fach, k.Lehrkraft }).Distinct();
                    var kkkk = iKlasse + "(";
                    foreach (var item in kll)
                    {
                        kkkk += item.Fach + "(" + item.Lehrkraft + "),";
                    }
                    kkkk = kkkk.TrimEnd(',');
                    kkkk = kkkk + ")";
                    if (kll.Count() > 0)
                    {
                        kl.Add(kkkk);
                    }
                }
            }

            AusgabeSchreiben("In folgenden Klassen sind in mehreren Fächern Zeugnisnoten gesetzt. Es fehlen noch: ", kl);
            Console.WriteLine((" " + kl.Count.ToString()).PadLeft(30, '.'));
        }

        internal void FehlendeZeugnisbemerkungBeiStrich(Leistungen webuntisLeistungen, List<string> interessierendeKlassen)
        {
            Console.Write(("Fehlende Zeugnisbemerkung bei Strich").PadRight(71, '.'));
            
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
                    AusgabeSchreiben("Es gibt Schüler*innen " + fehlendeBemerkungen.Count + " in Vollzeitklassen, die einen '-' in einem Nicht-Reli-Fach als Noten bekommen, ohne dass eine entsprechende Zeugnisbemerkung vorliegt:", fehlendeBemerkungen);
                }
                Console.WriteLine((" " + fehlendeBemerkungen.Count.ToString()).PadLeft(30, '.'));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        internal Leistungen HoleAlteNoten(Leistungen webuntisLeistungen, List<string> interessierendeKlassen, List<string> aktSj)
        {
            Console.Write(("Leistungsdaten aus alten Jahren holen").PadRight(71, '.'));
            try
            {
                List<string> holeFächer = new List<string>();

                Leistungen leistungen = new Leistungen();

                string abschlussklassen = "";

                foreach (var klasse in interessierendeKlassen)
                {                    
                    // Für Klassen der Anlage A, die in diesem Schuljahr in Jahrgang 3 oder 4 sind oder deren Schüler ein Abschlusszeugnis bekommen ... 

                    if ((from a in this where klasse == a.Klasse 
                         where a.Anlage.StartsWith(Properties.Settings.Default.Klassenart) 
                         where (a.Abschlussklasse || a.Zeugnisart == "A01AS") 
                         where a.Schuljahr == aktSj[0] + "/" + aktSj[1] 
                         select a).Any())
                    {
                        // ... werden die verschiedenen Schüler gesucht, die in diesem Schuljahr die Klasse besuchen.

                        var schuelers = (from w in webuntisLeistungen 
                                         where w.Klasse == klasse 
                                         where (from a in this
                                                where a.SchlüsselExtern == w.SchlüsselExtern
                                                where klasse == a.Klasse
                                                where a.Anlage.StartsWith(Properties.Settings.Default.Klassenart)
                                                where (a.Abschlussklasse || a.Zeugnisart == "A01AS")
                                                where a.Schuljahr == aktSj[0] + "/" + aktSj[1]
                                                select a).Any()
                                         select w.SchlüsselExtern).Distinct().ToList();

                        // Für jeden Schüler werden seine Noten der vergangenen Jahre gesucht ...

                        foreach (var schueler in schuelers)
                        {
                            if (schueler == 151703)
                            {
                                string f = "";
                            }
                            var gliederungDesSchuelers = (from a in this where a.SchlüsselExtern == schueler where a.Schuljahr == aktSj[0] + "/" + aktSj[1] where a.Klasse == klasse select a.Gliederung).FirstOrDefault();

                            var vergangeneLeistungenDiesesSchuelers = (from a in this
                                                                       where a.SchlüsselExtern == schueler
                                                                       where a.Schuljahr != aktSj[0] + "/" + aktSj[1]
                                                                       where a.Klasse != null
                                                                       where a.Klasse.Substring(0, 1) == klasse.Substring(0, 1)
                                                                       where a.Gliederung == gliederungDesSchuelers
                                                                       where a.Gesamtnote != null
                                                                       select a).OrderByDescending(x => x.Jahrgang).ToList();

                            var diesjähigeLeistungenDiesesSchuelers = (from a in this
                                                                       where a.SchlüsselExtern == schueler
                                                                       where a.Schuljahr == aktSj[0] + "/" + aktSj[1]
                                                                       where a.Klasse != null
                                                                       where a.Klasse.Substring(0, 1) == klasse.Substring(0, 1)
                                                                       where a.Gliederung == gliederungDesSchuelers
                                                                       select a).OrderByDescending(x => x.Jahrgang).ToList();

                            foreach (var item in this)
                            {
                                if (item.SchlüsselExtern == 151703)
                                {
                                    string a = "";
                                }
                            }

                            var alleVerschiedenenFächerDiesesSchuelers = (from a in this
                                                                          where a.Klasse != null
                                                                          where a.Klasse.Substring(0, 1) == klasse.Substring(0, 1)
                                                                          where a.Gliederung == gliederungDesSchuelers
                                                                          where a.Gesamtnote != null
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

                                    if (vLeistung != null)
                                    {
                                        if (!(from h in holeFächer where h.StartsWith(value: vLeistung.Fach + "(" + vLeistung.Klasse) select h).Any())
                                        {
                                            holeFächer.Add(vLeistung.Fach + "(" + vLeistung.Klasse + "|" + vLeistung.Schuljahr + ")");
                                                   
                                            Leistung aLeistung = (from v in diesjähigeLeistungenDiesesSchuelers where v.Fach == fach select v).FirstOrDefault();

                                            if (aLeistung != null)
                                            {
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
                                                leistung.Gesamtpunkte = leistung.Gesamtnote2Gesamtpunkte(leistung.Gesamtnote);
                                                leistung.Tendenz = leistung.Gesamtnote2Tendenz(leistung.Tendenz);
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
                                            else
                                            {
                                                Console.WriteLine("Problem beim Schüler (ID " + schueler +"; Gliederung: " + gliederungDesSchuelers + "; Fach " + fach + " )" );
                                                Console.ReadKey();
                                            }
                                        }
                                    }                                    
                                }
                            }
                        }
                    }
                }

                AusgabeSchreiben("Noten aus Fächern aus vorherigen Schuljahren werden geholt: ", holeFächer);

                Console.WriteLine((" " + holeFächer.Count.ToString()).PadLeft(30, '.'));

                return leistungen;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }

        internal Lehrers LehrerDieserKlasse(Termin konferenz, Lehrers lehrers)
        {
            Lehrers lehrersInDerKlasse = new Lehrers();

            foreach (var w in this)
            {
                if (w.Klasse == konferenz.Klasse)
                {
                    foreach (var l in lehrers)
                    {
                        if (l.Kuerzel == w.Lehrkraft)
                        {
                            if (!(from t in lehrersInDerKlasse where t.Kuerzel == l.Kuerzel select t).Any())
                            {
                                lehrersInDerKlasse.Add(l);
                            }
                        }
                    }
                }                           
            }

            return lehrersInDerKlasse;
        }

        public Leistungen(string connetionstringAtlantis, List<string> aktSj, string user)
        {
            Global.Output.Add("/* ************************************************************************************************* */");
            Global.Output.Add("/* Diese Datei enthält alle Noten und Fehlzeiten aus Webuntis. Sie können die Datei in Atlantis im-  */");
            Global.Output.Add("/* portieren, indem Sie sie in Atlantis unter Funktionen>SQL-Anweisung hochladen.                    */");
            Global.Output.Add("/* Published under the terms of GPLv3. Hoping for the best!                                          */");
            Global.Output.Add("/* " + (user + " " + DateTime.Now.ToString()).PadRight(97) + " */");
            Global.Output.Add("/* ************************************************************************************************* */");
            Global.Output.Add("");

            Global.Defizitleistungen = new Leistungen();

            try
            {
                Console.Write(("Leistungsdaten aus Atlantis (" + Global.HzJz + ")").PadRight(71, '.'));

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
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 0) + "/" + (Convert.ToInt32(aktSj[1]) - 0) + @"')
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 1) + "/" + (Convert.ToInt32(aktSj[1]) - 1) + @"' AND (klasse.jahrgang = 'A011' OR klasse.jahrgang = 'A012' OR klasse.jahrgang = 'A013')  AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 2) + "/" + (Convert.ToInt32(aktSj[1]) - 2) + @"' AND (klasse.jahrgang = 'A011' OR klasse.jahrgang = 'A012') AND s_note IS NOT NULL)
OR
  (vorgang_schuljahr = '" + (Convert.ToInt32(aktSj[0]) - 3) + "/" + (Convert.ToInt32(aktSj[1]) - 3) + @"' AND (klasse.jahrgang = 'A011') AND s_note IS NOT NULL)
)
ORDER BY DBA.klasse.s_klasse_art ASC , DBA.klasse.klasse ASC, DBA.noten_kopf.nok_id, DBA.noten_einzel.position_1; ", connection);


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
                            // Leistungen des aktuellen Abschnitts sind abhängig vom aktuellen Monat.
                            // Leistungen vergangener Jahre sind immer "JZ"

                            if (theRow["Schuljahr"].ToString() == (Convert.ToInt32(aktSj[0]) - 0) + "/" + (Convert.ToInt32(aktSj[1]) - 0) || (theRow["Schuljahr"].ToString() != (Convert.ToInt32(aktSj[0]) - 0) + "/" + (Convert.ToInt32(aktSj[1]) - 0) && theRow["HzJz"].ToString() == "JZ"))
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
                                        leistung.HatBemerkung = (theRow["Bemerkung1"].ToString() + theRow["Bemerkung2"].ToString() + theRow["Bemerkung3"].ToString()).Contains("Fehlzeiten") ? true : false;
                                        leistung.Jahrgang = Convert.ToInt32(theRow["Jahrgang"].ToString().Substring(3, 1));
                                        leistung.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                                        leistung.Nachname = theRow["Nachname"].ToString();
                                        leistung.Vorname = theRow["Vorname"].ToString();
                                        leistung.Bereich = bereich;
                                        leistung.Geburtsdatum = theRow["dat_geburt"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["dat_geburt"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                        leistung.Volljährig = leistung.Geburtsdatum.AddYears(18) > DateTime.Now ? false : true;
                                        leistung.Klasse = theRow["Klasse"].ToString();
                                        leistung.Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();
                                        leistung.Gesamtnote = theRow["Note"].ToString() == "" ? null : theRow["Note"].ToString() == "Attest" ? "A" : theRow["Note"].ToString();                                        
                                        //leistung.Gesamtpunkte_12_1 = theRow["Punkte_12_1"].ToString() == "" ? null : (theRow["Punkte_12_1"].ToString()).Split(',')[0];
                                        //leistung.Gesamtpunkte_12_2 = theRow["Punkte_12_2"].ToString() == "" ? null : (theRow["Punkte_12_2"].ToString()).Split(',')[0];
                                        //leistung.Gesamtpunkte_13_1 = theRow["Punkte_13_1"].ToString() == "" ? null : (theRow["Punkte_13_1"].ToString()).Split(',')[0];
                                        //leistung.Gesamtpunkte_13_2 = theRow["Punkte_13_2"].ToString() == "" ? null : (theRow["Punkte_13_2"].ToString()).Split(',')[0];
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
                                        leistung.GeholteNote = false;                                        
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Fehler beim Einlesen der Atlantis-Leistungsdatensätze: ENTER" + ex);
                                    Console.ReadKey();
                                }
                                this.Add(leistung);

                                if (leistung.Gesamtnote == "5" || leistung.Gesamtnote == "6")
                                {
                                    Global.Defizitleistungen.Add(leistung);
                                }                                
                            }
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
            Global.PrintMessage(Global.Output.Count(), "Atlantisleistungen: ".PadRight(45, '.') + (" " + this.Count.ToString()).PadLeft(45, '.'));
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
            Console.WriteLine((" " + i).PadLeft(30, '.'));
        }

        internal void BindestrichfächerZuordnen(Leistungen atlantisLeistungen)
        {
            Console.Write(("Bindestrichfächer zuordnen").PadRight(71, '.'));
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
            Console.WriteLine((" " + i.ToString()).PadLeft(30, '.'));
        }

        internal void ReligionsabwählerBehandeln(Leistungen atlantisLeistungen)
        {
            Console.Write(("Religionsabwähler mit '-' versehen").PadRight(71, '.'));
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
            Console.WriteLine((" " + i.ToString()).PadLeft(30, '.'));
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
            Console.WriteLine((" " + i.ToString()).PadLeft(30, '.'));
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
                            AusgabeSchreiben("Klasse: " + klasse + ". Das Webuntisfach " + webuntisFach + " wird dem Atlantisfach " + af.Fach + " zugeordnet. Wenn das nicht gewünscht ist, dann hier mit Suchen & Ersetzen der gewünschte Name gesetzt werden.",new List<string>());
                            
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
            Console.WriteLine((" " + i.ToString()).PadLeft(30, '.'));
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

        internal void Add(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen)
        {

            Console.Write(("Neu anzulegende Leistungen in Atlantis ").PadRight(71, '.'));
            int i = 0;
            int outputIndex = Global.Output.Count();
            
            try
            {
                foreach (var klasse in interessierendeKlassen)
                {
                    foreach (var a in (from t in this where t.Klasse == klasse where t.HzJz != "GO" where t.SchuelerAktivInDieserKlasse select t).OrderBy(x => x.Klasse).ThenBy(x => x.Fach).ThenBy(x => x.Name).ToList())
                    {
                        if (IstRelevant(a, webuntisLeistungen))
                        {
                            var w = (from webuntisLeistung in webuntisLeistungen
                                     where webuntisLeistung.Fach == a.Fach
                                     where webuntisLeistung.Klasse == a.Klasse
                                     where webuntisLeistung.SchlüsselExtern == a.SchlüsselExtern
                                     where a.Gesamtnote == null
                                     where webuntisLeistung.Gesamtnote != null // Gibt den Fall, dass nicht hinter allen Teilleistungen eine Gesamtnote steht.
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
                                    if (w.Gesamtnote == "-" && a.Bemerkung == null && a.Fach != "REL" && a.Fach != "KR G1")
                                    {
                                        w.Beschreibung += "Bemerkung fehlt";
                                    }
                                    UpdateLeistung(a.Klasse.PadRight(6) + "|" + (a.Name.Substring(0, Math.Min(a.Name.Length, 6))).PadRight(6) + "|" + w.Gesamtnote + " " + "|" + w.Fach.PadRight(4) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET      s_note='" + w.Gesamtnote + "' WHERE noe_id=" + a.LeistungId + ";");

                                    if (w.Gesamtnote == "5" || w.Gesamtnote == "6")
                                    {
                                        Global.Defizitleistungen.Add(w);
                                    }

                                    if (a.EinheitNP == "P")
                                    {
                                        // Wenn die Tendenzen abweichen ...

                                        if (w.Tendenz != a.Tendenz)
                                        {
                                            UpdateLeistung(a.Klasse.PadRight(6) + "|" + (a.Name.Substring(0, Math.Min(a.Name.Length, 6))).PadRight(6) + "| " + w.Tendenz + "|" + a.Fach + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET   s_tendenz=" + (w.Tendenz == null ? "NU" : "'" + w.Tendenz + "'") + " WHERE noe_id=" + a.LeistungId + ";"); ;
                                        }

                                        UpdateLeistung(a.Klasse.PadRight(6) + "|" + (a.Name.Substring(0, Math.Min(a.Name.Length, 6))).PadRight(6) + "|" + w.Gesamtpunkte.PadLeft(2) + "|" + w.Fach.PadRight(4) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET      punkte=" + w.Gesamtpunkte.PadLeft(2) + "  WHERE noe_id=" + a.LeistungId + ";");

                                        //// Bei Klassen der gym Oberstufe mit Gym-Notenblatt ...

                                        //if (a.HzJz == "GO")
                                        //{
                                        //    if (a.Jahrgang == 2)
                                        //    {
                                        //        if (a.Konferenzdatum.Month == 1)
                                        //        {
                                        //            UpdateLeistung(a.Klasse.PadRight(6) + "|" + (a.Name.Substring(0, Math.Min(a.Name.Length, 6))).PadRight(6) + "|" + w.Gesamtpunkte.PadLeft(2) + "|" + w.Fach.PadRight(4) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET punkte_12_1=" + w.Gesamtpunkte.PadLeft(2) + "  WHERE noe_id=" + a.LeistungId + ";");
                                        //        }
                                        //    }
                                        //    if (a.Jahrgang == 2)
                                        //    {
                                        //        if (a.Konferenzdatum.Month > 2)
                                        //        {
                                        //            UpdateLeistung(a.Klasse.PadRight(6) + "|" + (a.Name.Substring(0, Math.Min(a.Name.Length, 6))).PadRight(6) + "|" + w.Gesamtpunkte.PadLeft(2) + "|" + w.Fach.PadRight(4) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET punkte_12_2=" + w.Gesamtpunkte.PadLeft(2) + "  WHERE noe_id=" + a.LeistungId + ";");
                                        //        }
                                        //    }
                                        //    if (a.Jahrgang == 3)
                                        //    {
                                        //        if (a.Konferenzdatum.Month == 1)
                                        //        {
                                        //            UpdateLeistung(a.Klasse.PadRight(6) + "|" + (a.Name.Substring(0, Math.Min(a.Name.Length, 6))).PadRight(6) + "|" + w.Gesamtpunkte.PadLeft(2) + "|" + w.Fach.PadRight(4) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET punkte_13_1=" + w.Gesamtpunkte.PadLeft(2) + "  WHERE noe_id=" + a.LeistungId + ";");
                                        //        }
                                        //    }
                                        //    if (a.Jahrgang == 3)
                                        //    {
                                        //        if (a.Konferenzdatum.Month > 2)
                                        //        {
                                        //            UpdateLeistung(a.Klasse.PadRight(6) + "|" + (a.Name.Substring(0, Math.Min(a.Name.Length, 6))).PadRight(6) + "|" + w.Gesamtpunkte.PadLeft(2) + "|" + w.Fach.PadRight(4) + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET punkte_13_2=" + w.Gesamtpunkte.PadLeft(2) + "  WHERE noe_id=" + a.LeistungId + ";");
                                        //        }
                                        //    }
                                        //}
                                    }

                                    i++;
                                }
                            }
                        }                                               
                    }
                }

                Global.PrintMessage(outputIndex, ("Neu anzulegende Leistungen in Atlantis: ").PadRight(65, '.') + (" " + i.ToString()).PadLeft(30, '.'));

                Console.WriteLine((" " + i.ToString()).PadLeft(30, '.'));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private bool IstRelevant(Leistung a, List<Leistung> webuntisLeistungen)
        {
            // Bei Jahreszeugnissen in Anlage C und D im Jahrgang 11 werden keine Halbjahresnoten gezogen


            // In Anlage A werden im Jahreszeugnis in Jahrgang 3 Noten aus alten Jahren geholt.
            return true;
        }

        public string Gesamtpunkte2Gesamtnote(string gesamtpunkte)
        {
            string gesamtnote = null;

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
            int outputIndex = Global.Output.Count();
            Console.Write(("Leistungen in Atlantis updaten").PadRight(71, '.'));
            int i = 0;

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
                                 where ((a.Gesamtpunkte != null && a.Gesamtpunkte != webuntisLeistung.Gesamtpunkte) || (a.Gesamtnote != null && a.Gesamtnote != webuntisLeistung.Gesamtnote))
                                 where webuntisLeistung.Gesamtpunkte != null
                                 select webuntisLeistung).FirstOrDefault();

                        if (w != null)
                        {
                            // Ein '-' in Religion wird deleted, wenn kein anderer Schüler eine Gesamtnote bekommen hat.

                            if (!(w.Fach == "REL" && w.Gesamtnote == "-" && (from we in webuntisLeistungen
                                                                             where we.Klasse == w.Klasse
                                                                             where we.Fach == "REL"
                                                                             where (we.Gesamtnote != null && we.Gesamtnote != "-")
                                                                             select we).Count() == 0))
                            {
                                if (a.Gesamtnote != w.Gesamtnote)
                                {
                                    UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 5)) + "|" + a.Gesamtnote + (a.Tendenz ?? " ") + ">" + w.Gesamtnote + (w.Tendenz ?? " ") + "|" + a.Fach + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET      s_note='" + w.Gesamtnote + "'  WHERE noe_id=" + a.LeistungId + ";");
                                }
                                
                                if (a.EinheitNP == "P")
                                {
                                    UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 5)) + "|" + (a.Gesamtpunkte == null ? "" : a.Gesamtpunkte.Split(',')[0]).PadLeft(2) + ">" + (w.Gesamtpunkte == null ? " " : w.Gesamtpunkte.Split(',')[0]).PadLeft(2) + "|" + a.Fach + (w.Beschreibung == null ? "" : w.Beschreibung), "UPDATE noten_einzel SET      punkte=" + (w.Gesamtpunkte).PadLeft(2) + "   WHERE noe_id=" + a.LeistungId + ";");
                                    
                                    if (a.Tendenz != w.Tendenz)
                                    {
                                        UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 5)) + "|" + (a.Tendenz == null ? "NULL" : a.Tendenz + " ") + ">" + (w.Tendenz == null ? "NU" : w.Tendenz + " ") + "|" + a.Fach + (w.Beschreibung == null ? "" : "|" + w.Beschreibung.TrimEnd(',')), "UPDATE noten_einzel SET   s_tendenz=" + (w.Tendenz == null ? "NULL " : "'" + w.Tendenz + "' ") + "WHERE noe_id=" + a.LeistungId + ";");
                                    }
                                }
                                i++;
                            }   
                        }
                    }                    
                }
                Console.WriteLine((" " + i.ToString()).PadLeft(30, '.'));
                Global.PrintMessage(outputIndex, ("Zu ändernde Leistungen in Atlantis: ").PadRight(65, '.') + (" " + i.ToString()).PadLeft(30, '.'));
            }
            catch (Exception ex)
            {
                throw ex;
            }            
        }

        internal void Delete(List<Leistung> webuntisLeistungen, List<string> interessierendeKlassen, List<string> aktSj)
        {
            Console.Write(("Zu löschende Leistungen in Atlantis ").PadRight(71, '.'));
            int outputIndex = Global.Output.Count();                        
            int i = 0;

            try
            {
                foreach (var a in this)
                {
                    if (a.LeistungId == 4368314)
                    {
                        string aa = "";
                    }
                    if (a.SchuelerAktivInDieserKlasse)
                    {
                        // Wenn es zu einem Atlantis-Datensatz keine Entsprechung in Webuntis gibt, ... 
                        
                        var webuntisLeistung = (from w in webuntisLeistungen
                                                where w.Fach == a.Fach
                                                where w.Klasse == a.Klasse
                                                where w.SchlüsselExtern == a.SchlüsselExtern
                                                where w.Gesamtpunkte != null
                                                select w).FirstOrDefault();

                        if (webuntisLeistung == null)
                        {
                            if (a.Gesamtnote != null)
                            {
                                // Geholte Noten aus Vorjahren werden nicht gelöscht.

                                if (!a.GeholteNote)
                                {   
                                    UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 6)) + "|" + a.Fach.Substring(0,Math.Min(a.Fach.Length, 3)).PadRight(3) + "|" + (a.Gesamtnote == null ? "NULL" : a.Gesamtnote+ (a.Tendenz == null ? " " : a.Tendenz)) + ">NULL" + a.Beschreibung, "UPDATE noten_einzel SET      s_note=NULL WHERE noe_id=" + a.LeistungId + ";");
                                    
                                    if (a.Gesamtpunkte != null)
                                    {
                                        UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 6)) + "|" + a.Fach.Substring(0, Math.Min(a.Fach.Length, 3)).PadRight(3) + "|" + ((a.Gesamtpunkte == null ? "NULL" : a.Gesamtpunkte.Split(',')[0])).PadLeft(2) + ">NULL" + a.Beschreibung, "UPDATE noten_einzel SET      punkte=NULL WHERE noe_id=" + a.LeistungId + ";");                                       
                                    }
                                    
                                    if (a.Tendenz != null)
                                    {
                                        UpdateLeistung(a.Klasse + "|" + a.Name.Substring(0, Math.Min(a.Name.Length, 6)) + "|" + a.Fach.Substring(0,Math.Min(a.Fach.Length, 3)).PadRight(3) + "|" + (a.Tendenz == null ? "NULL" : " " + a.Tendenz) + ">NULL" + a.Beschreibung, "UPDATE noten_einzel SET s_tendenz=NULL WHERE noe_id=" + a.LeistungId + ";");
                                    }                                    
                                    i++;
                                }
                            }                            
                        }
                    }
                }
                Global.PrintMessage(outputIndex, ("Zu löschende Leistungen in Atlantis: ").PadRight(65, '.') + (" " + i.ToString()).PadLeft(30, '.'));
                Console.WriteLine((" " + i.ToString()).PadLeft(30, '.'));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //internal List<string> GetIntessierendeKlassen(Leistungen alleWebuntisLeistungen, List<string> aktSj)
        //{
        //    Console.Write(("Interessierende Klassen").PadRight(71, '.'));

        //    var interessierendeKlassen = (from k in alleWebuntisLeistungen where k.Gesamtnote != null where k.Klasse != null select k.Klasse).Distinct().ToList();

        //    Console.WriteLine((" " + interessierendeKlassen.Count.ToString()).PadLeft(30, '.'));

        //    interessierendeKlassen = AddKonferenzdatum(interessierendeKlassen);

        //    AusgabeSchreiben("Folgende Klassen mit Gesamtnoten in Webuntis werden ausgewertet:", AddAbsckusszeugnis(interessierendeKlassen));

        //    //var klassenMitZeugnisnotenInMehrAls3Fächern = (from k in alleWebuntisLeistungen where k.Gesamtnote != null group k by k.Klasse)

        //    //var klassenMitFehlendenNoten = (from k in alleWebuntisLeistungen where k.Gesamtnote == null where k.Klasse != null select k.Klasse).Distinct().ToList();
        //    //AusgabeSchreiben("In folgenden Klassen sind bereits mehrere Zeugnisnoten gesetzt. Es fehlen noch:", klassenMitFehlendenNoten);

        //    return interessierendeKlassen;
        //}

        internal List<string> GetIntessierendeKlassen(Leistungen alleWebuntisLeistungen, List<string> aktSj)
        {
            var infragekommendeKlassen = (from k in alleWebuntisLeistungen where k.Klasse != null select k.Klasse).Distinct().ToList();

            List<string> ausgewählteKlassen = new List<string>();

            if (Properties.Settings.Default.Klassenwahl == "alle" || Properties.Settings.Default.Klassenwahl == "")
            {
                ausgewählteKlassen = infragekommendeKlassen;
            }
            else
            {
                foreach (var klasse in Properties.Settings.Default.Klassenwahl.Split(','))
                {
                    if (!ausgewählteKlassen.Contains(klasse.Trim()))
                    {
                        if (infragekommendeKlassen.Contains(klasse.Trim()))
                        {
                            ausgewählteKlassen.Add(klasse.Trim());
                        }
                    }
                }
            }

            Properties.Settings.Default.Klassenwahl = "";

            string ausgewählteKlassenString = "";

            foreach (var item in ausgewählteKlassen)
            {
                ausgewählteKlassenString += item + ",";
            }

            if (ausgewählteKlassenString == "")
            {
                ausgewählteKlassenString = "alle";
            }

            Console.WriteLine("");
            Console.WriteLine("Möchten Sie die Auswahl der Klassen einschränken? Das können Sie tun:");
            Console.WriteLine("");
            Console.WriteLine(" * kommasepariert Klassennamen eingeben");
            Console.WriteLine(" * kommasepariert Anfangsbuchstaben oder Anfänge von Klassenamen eingeben");
            Console.WriteLine(" * das Wort 'alle' eintippen");
            Console.WriteLine(" * die Vorauswahl mit ENTER bestätigen");
            Console.WriteLine("");
            Console.Write("Vorauswahl: [ " + ausgewählteKlassenString.TrimEnd(',') + "]: ");

            var x = Console.ReadLine();

            Console.Write(("Ausgewählte Klassen ").PadRight(71, '.'));

            if (x == "" && ausgewählteKlassenString == "alle")
            {
                ausgewählteKlassen = infragekommendeKlassen;
            }

            if (x != "")
            {
                ausgewählteKlassen = new List<string>();

                foreach (var klasse in x.Split(','))
                {
                    foreach (var ikl in infragekommendeKlassen)
                    {
                        if (ikl.StartsWith(klasse))
                        {
                            if (!ausgewählteKlassen.Contains(ikl))
                            {
                                ausgewählteKlassen.Add(ikl);
                            }
                        }
                    }
                }
            }

            foreach (var item in ausgewählteKlassen)
            {
                Properties.Settings.Default.Klassenwahl += item + ",";
            }

            Properties.Settings.Default.Klassenwahl = Properties.Settings.Default.Klassenwahl.TrimEnd(',');
            Properties.Settings.Default.Save();

            Console.WriteLine((" " + ausgewählteKlassen.Count.ToString()).PadLeft(30, '.'));

            ausgewählteKlassen = AddKonferenzdatum(ausgewählteKlassen);


            AusgabeSchreiben("Folgende Klassen mit Gesamtnoten in Webuntis werden ausgewertet:", AddAbschlusszeugnis(ausgewählteKlassen));

            return ausgewählteKlassen;
        }

        private List<string> AddAbschlusszeugnis(List<string> interessierendeKlassen)
        {
            List<string> ik = new List<string>();

            foreach (var item in interessierendeKlassen)
            {
                if ((from a in this where a.Abschlussklasse == true where a.Anlage.StartsWith("A") where a.Klasse == item select a).Any())
                {
                    ik.Add(item + " (Abschlusszeugnis)");
                }
                else
                {
                    ik.Add(item);
                }
            }
            return ik;
        }

        private List<string> AddKonferenzdatum(List<string> interessierendeKlassen)
        {
            List<string> interessierendeKlassenGefiltert = new List<string>();

            List<string> klassenMitZurückliegenderZeugniskonferenz = new List<string>();

            foreach (var klasse in interessierendeKlassen)
            {
                foreach (var a in (from t in this where t.Klasse == klasse select t).ToList())
                {
                    // Wenn die Notenkonferenz -je nach Halbjahr- ...

                    if (a != null && a.HzJz == "HZ")
                    {
                        // ... in der Vergangenheit liegen ...

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
                            if (!(from i in interessierendeKlassenGefiltert where i == klasse select i).Any())
                            {
                                interessierendeKlassenGefiltert.Add(klasse);
                            }                            
                        }
                    }
                }
            }
                        
            if (klassenMitZurückliegenderZeugniskonferenz.Count > 0)
            {
                AusgabeSchreiben("Für folgende Klassen sind zwar Noten in Webuntis eingetragen, wegen der zurückliegenden Notenkonferenzen werden sie aber nicht angefasst:", klassenMitZurückliegenderZeugniskonferenz);
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

        public void AusgabeSchreiben(string text, List<string> klassen)
        {
            Global.Output.Add("");
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
                Global.Output.Add((o.Substring(0, Math.Min(101, o.Length))).PadRight(101) + "*/");

            } while (z < text.Split(' ').Count());



            z = 0;
            
            do
            {
                var zeile = " ";                

                try
                {
                    if (klassen[z].Length >= 95)
                    {
                        klassen[z] = klassen[z].Substring(0, Math.Min(klassen[z].Length, 95));
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
                Global.Output.Add((o.Substring(0, Math.Min(101, o.Length))).PadRight(101) + "*/");

            } while (z < klassen.Count);
        }

        private void UpdateLeistung(string message, string updateQuery)
        {
            try
            {
                string o = updateQuery + "/* " + message;
                Global.Output.Add((o.Substring(0, Math.Min(100, o.Length))).PadRight(100) + " */");
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

            OpenFiles(files);
        }

        private void OpenFiles(List<string> files)
        {
            try
            {
                Process notepadPlus = new Process();
                notepadPlus.StartInfo.FileName = "notepad++.exe";

                for (int i = 0; i < files.Count; i++)
                {
                    if (i==0)
                    {
                        notepadPlus.StartInfo.Arguments = @"-multiInst -nosession " + files[i];
                    }
                    else
                    {
                        notepadPlus.StartInfo.Arguments = files[i];
                    }
                }
                notepadPlus.Start();
            }
            catch (Exception)
            {
                System.Diagnostics.Process.Start("Notepad.exe", files[0]);
            }
        }        
    }
}