# WebuntisNoten2Atlantis

Mit *WebuntisNoten2Atlantis* können die Zeugnisnoten, Noten für Blaue Briefe und Fehlzeiten von Webuntis nach Atlantis übertragen werden, ohne dass einzelne Noten oder Fehlzeiten händisch angefasst werden.

## Voraussetzungen

* Klassen- und Fächerbezeichnungen sind (mit Ausnahmen) identisch in Untis und Atlantis.
* Ein Notenblatt wurde in Atlantis angelegt.
* Zeugnisnoten wurden als Gesamtnoten in Webuntis erfasst. Alle Noten in allen Anlagen werden als Noten mit Tendenzen (+ oder -) eingegeben. Die Lehrkraft muss Teilleistungen in Webuntis eingetragen haben, damit die Gesamtnote übertragen wird.
* Der Benutzer hat die Berechtigung Prüfungen und Fehlzeiten aus Webuntis zu exportieren und SQL-Dateien in Atlantis zu importieren.
* Die Schülerinnen und Schüler haben in Webuntis die Atlantis-ID als externen Schlüssel gesetzt bekommen.

## Schritte

1. Prüfungen aus Webuntis exportieren und im Download-Ordner ablegen.
2. Fehlzeiten aus Webuntis exportieren und im Download-Ordner ablegen.
3. Unterrichte, Gruppen und Schülerinnen und Schüler exportieren.
4. *WebuntisNoten2Atlantis* starten.
   Falls die Zuordnung eines Webuntis-Fachs nicht möglich ist, wird das angezeigt. Eine Zuordnung kann dann händisch vorgenommen werden.
5. Erzeugte SQL-Datei in Atlantis importieren.

### Prüfungen aus Webuntis exportieren

1. Mit administrativer Berechtigung in Webuntis anmelden.
2. Den Pfad *Klassenbuch > Berichte* gehen.
3. Alle Klassen wählen oder auch nur eine Klasse wählen oder auch nur einen Schüler in einer Klasse wählen. Das aktuelle Schuljahr oder das Halbjahr wählen.
4. Unter der Rubrik *Noten* das Icon *CSV-Ausgabe* klicken. Der Haken bei *Notennamen ausgeben* darf _nicht_ gesetzt sein. 
5. Die Datei *MarksPerLesson.CSV* im Download-Ordner speichern. Die Datei hat möglicherweise sehr viele Zeilen und es kann dauern, bis Webuntis die Datei bereitstellt. Die Datei hat folgenden Aufbau:

```
Datum	Name	Klasse	Fach	Prüfungsart	Note	Bemerkung	Benutzer	Schlüssel (extern)	Gesamtnote
28.09.2019	Müller Ina	HBW20A	INW	schriftl. Arbeit	12.0		BM	149565   11.0	
```

Die Gesamtnote steht in der letzten Spalte als Punktzahl von ```0``` bis ```15```. Davor steht die Atlantis-ID.

### Fehlzeiten aus Webuntis exportieren

1. Mit administrativer Berechtigung in Webuntis anmelden.
2. Den Pfad *Administration > Export* gehen.
3. Hinter *Gesamtfehlzeiten* das Icon *CSV-Ausgabe* klicken.
5. Die Datei *AbsenceTimesTotal.CSV* i Download-Ordner speichern. Die Datei hat möglicherweise sehr viele Zeilen und folgenden Aufbau:

```
studentId	name		klasse	klasseId	absentMins	absentMinsNotExcused	absentHours	absentHoursNotExcused
		151989	Müller	Ina	HBW20A	2984	360	0	8	1
```

Die Atlantis-ID heißt hier ```studentId```. Diese Schülerin hat 360 Minuten gefehlt. Davon sind 0 Minuten unentschldigt.

### WebuntisNoten2Atlantis bedienen

1. Programm im Visual Studio selbst kompilieren und starten oder *webuntis2Atlantis.exe* herunterladen und starten.
2. Mit dem Starten des Programms werden die Bedingungen der Open Source Lizenz GPLv3 anerkannt.
3. Eine Datei namens *20200930-webuntisnoten2atlantis.SQL* wird erzeugt und öffnet sich im Notepad. 


### Die Datei 20200930-webuntisnoten2atlantis.SQL
Die Noten werden _nicht_ unmittelbar vom Programm *WebuntisNoten2Atlantis* in die Atlantis-Datenbank zurückgeschrieben. Stattdessen wird eine SQL-Datei erzeugt, die die Noten als SQL-Anweisungen an die Datenbank übergibt.

Jede Note entspricht einer SQL-Anweisungen in einer Zeile. Die Anweisungen sind alle vollkommen unabhängig voneinander und sehen wie folgt aus:
```SQL
UPDATE noten_einzel SET s_note='3' WHERE noe_id=3760033;/*HBW20A,INW,3,Müller I*/
```  
Hinter dem Semikolon steht ein kurzer Kommentar, der das inhaltliche Prüfen des Befehls vereinfachen soll. 
Die gezeigte Zeile sagt aus, dass die Schülerin Ina Müller aus der Klasse HBW20A im Fach INW die Note 3 eingetragen bekommt. Der Notendatensatz hat die ID 3760033. Bei Schülerinnen oder Schülern der Anlage D wird zusätzlich die Punktzahl und die Tendenz in einem zweiten SQL-Statement übergeben. WebuntisNoten2Atlantis sorgt dafür, dass die Einträge in der Datenbank widerspruchsfrei bleiben.
Nach der sorgfältigen Prüfung der Datei kann sie in Altlantis (entsprechende Berechtigungen vorausgesetzt) unter *Funktionen>SQL-Anweisung ausführen* in die Datenbank eingelesen werden. 
Evtl. macht es Sinn zunächst alle Zeilen bis auf eine Befehlszeile zu löschen und dann auszuführen. So kann zunächst bei einer einzelnen Note eines einzelnen Schülers geprüft werden, ob alles funktioniert.

### FAQ

#### Was ist mit den Umlauten passiert?
Die SQL-Anweisungen selbst enthalten niemals Umlaute. Insofern ist das unkritisch. Programmseitig ist das Encoding auf *Default* gestellt.

#### Was ist, wenn die Namen der Fächer in Atlantis und Untis nicht übereinstimmen
Grundsätzlich müssen die Fächerkürzel übereinstimmen. *WebuntisNoten2Atlantis* hat aber Routinen eingebaut, die versuchen eine Matching herzustellen. 
Bei Sprachen darf beispielsweise in Untis auf die Angabe der Niveaustufe verzichtet werden. Es wird dann automatisch versucht auf die Sprache in Atlantis zu matchen.
Wenn ein Fach nicht zugeordnet weren kann, wird das gemeldet. Es besteht dann die Möglichkeit einer manuellen Zuordnung.

#### Kann *WebuntisNoten2Atlantis* in Anlage A Noten aus Vorjahreszeugnissen ziehen?
Ja.

#### Unter welcher Lizenz steht das Programm?
GNU General Public License v3.0
This program comes with ABSOLUTELY NO WARRANTY.
This is free software, and you are welcome to redistribute it under certain conditions.

Stefan Bäumer
stefan.baeumer@berufskolleg-borken.de 27.12.2020
