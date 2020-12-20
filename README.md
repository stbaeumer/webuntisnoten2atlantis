# WebuntisNoten2Atlantis

Mit *WebuntisNoten2Atlantis* können die Zeugnisnoten und Fehlzeiten von Webuntis nach Atlantis übertragen werden, ohne dass einzelne Noten oder Fehlzeiten händisch angefasst werden.

## Voraussetzungen

* Klassen- und Fächerbezeichnungen sind (mit Ausnahmen) identisch in Untis und Atlantis.
* Ein Notenblatt wurde in Atlantis angelegt.
* Zeugnisnoten wurden als Gesamtnoten in Webuntis erfasst.
* Der Benutzer hat Berechtigung Prüfungen und Fehlzeiten aus Webuntis zu exportieren und SQL-Dateien in Atlantis zu importieren.
* Die Schülerinnen und Schüler haben in Webuntis die Atlantis-ID als externen Schlüssel gesetzt bekommen.

## Vorgehen

1. Prüfungen aus Webntis exportieren und auf den Desktop legen.
2. Fehlzeiten aus Webuntis exportieren und auf den Desktop legen.
2. *WebuntisNoten2Atlantis* starten.
3. Optinal gewünschte Klassen filtern.
4. Erzeugte SQL-Datei in Atlantis importieren.

### Prüfungen aus Webntis exportieren

1. Mit administrativer Berechtigung in Webuntis anmelden.
2. Den Pfad *Klassenbuch > Berichte* gehen.
3. Alle Klassen wählen. Das aktuelle Schuljahr wählen.
4. Unter der Rubrik *Noten* das Icon *CSV-Ausgabe* klicken. Der Haken bei *Notennamen ausgeben* darf nicht gesetzt sein.
5. Die Datei *MarksPerLesson.csv* auf dem Desktop speichern. Die Datei hat möglicherweise sehr viele Zeilen und folgenden Aufbau:

```
Datum	Name	Klasse	Fach	Prüfungsart	Note	Bemerkung	Benutzer	Schlüssel (extern)	Gesamtnote
28.09.2019	Müller Ina	HHO1	INW	Halbjahreszeugnis	12.0		BM	149565   11.0	
```

Die Gesamtnote steht in der letzten Spalte.  Davor steht die Atlantis-ID.

### Fehlzeiten aus Webntis exportieren

1. Mit administrativer Berechtigung in Webuntis anmelden.
2. Den Pfad *Administration > Export* gehen.
3. Hinter *Gesamtfehlzeiten* das Icon *CSV-Ausgabe* klicken.
5. Die Datei *AbsenceTimesTotal.csv* auf dem Desktop speichern. Die Datei hat möglicherweise sehr viele Zeilen und folgenden Aufbau:

```
studentId	name		klasse	klasseId	absentMins	absentMinsNotExcused	absentHours	absentHoursNotExcused
		151989	Müller	Ina	HHO1	2984	360	0	8	1
```

Der externe Schlüssel heißt hier ```studentId``` und entspricht der Atlantis-ID des Schülers.

### WebuntisNoten2Atlantis bedienen

1. Programm im Visual Studio selbst kompilieren und starten oder *webuntis2Atlantis.exe* herunterladen und starten.
2. Mit dem Starten des Programms werden die Bedingungen der Open Source Lizenz GPLv3 anerkannt.
3. Optional kann eine Filter auf die gewünschten Klassen gesetzt werden. Das macht sicherlich Sinn, wenn es gilt erste vorsichtige Erfahrungen zu sammeln.
4. Eine Datei namens *20190930-webuntisnoten2atlantis.SQL* wird erzeugt und öffnet sich im Notepad. 


### Die Datei webuntisnoten2atlantis_20190930.SQL
Die Noten werden _nicht_ unmittelbar vom Programm *WebuntisNoten2Atlantis* in die Atlantis-Datenbank zurückgeschrieben. Stattdessen wird eine SQL-Datei erzeugt, die die Noten als SQL-Anweisungen an die Datenbank übergibt.

Jede Note entspricht einer SQL-Anweisungen in einer Zeile. Die Anweisungen sind alle vollkommen unabhängig voneinander und sehen wie folgt aus:
```SQL
UPDATE noten_einzel SET s_note='3' WHERE noe_id=3760033;/*HHO1,INW,3,Müller I*/
```  
Hinter dem Semikolon steht ein kurzer Kommentar, der das Prüfen der Datei vereinfachen soll. 
Die gezeigte Zeile sagt aus, dass die Schülerin Ina Müller aus der Klasse HHO1 im Fach INW die Note 3 eingetragen bekommt. Der Notendatensatz hat die ID 3760033. Bei Sschülerinnen oder Schülern der Anlage D wird zusätzlich die Punktzahl in einem zweiten SQL-Statement übergeben.
Nach der sorgfältigen Prüfung der Datei kann sie in Altlantis (entsprechende Berechtigungen vorausgesetzt) unter *Funktionen>SQL-Anweisung ausführen* in die Datenbank eingelesen werden. 
Evtl. macht es Sinn zunächst alle bis auf eine SQL-Anweisung zu löschen und dann auszuführen. So kann zunächst bei einer einzelnen Note eines einzelnen Schülers geprüft werden, ob alles funktioniert.

### FAQ

#### Was ist mit den Umlauten passiert?
Die SQL-Anweisungen selbst enthalten niemals Umlaute. Insofern ist das unkritisch. Programmseitig ist das Encoding auf *Default* gestellt.

#### Was ist, wenn die Namen der Fächer in Atlantis und Untis nicht übereinstimmen
Grundsätzlich müssen die Fächerkürzel überinstimmen. *WebuntisNoten2Atlantis* hat aber Routinen, die versuchen eine Matching herzustellen. 
Bei Sprachen darf beispielsweise in Untis auf die Angabe der Niveaustufe verzichtet werden. Es wird dann automatisch versucht auf die Sprache in Atlantis zu matchen.
Wenn ein Fach nicht zugeordnet weren kann, wird das gemeldet.

#### Unter welcher Lizenz steht das Programm?
GNU General Public License v3.0
This program comes with ABSOLUTELY NO WARRANTY.
This is free software, and you are welcome to redistribute it under certain conditions.

stefan.baeumer@berufskolleg-borken.de 03.10.2019