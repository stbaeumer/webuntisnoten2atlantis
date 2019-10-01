# WebuntisNoten2Atlantis

Mit *WebuntisNoten2Atlantis* können die Zeugnisnoten von Webuntis nach Atlantis übertragen werden, ohne dass einzelne Noten händisch angefasst werden.

## Voraussetzungen

* Klassen- und Fächerbezeichnungen sind identisch in Untis und Atlantis.
* Notenblatt und Zeugnisformular wurden in Atlantis angelegt.
* Eine Prüfungsart wurde in Webuntis angelegt, die denselben Namen trägt wie das Zeugnisformular in Atlantis.
* Der Benutzer hat Berechtigung Prüfungen aus Webuntis zu exportieren und SQL-Dateien in Atlantis zu importieren .               
* Die Schülerinnen und Schüler haben in Webuntis die Atlantis-ID als externen Schlüssel gesetzt bekommen.

## Vorgehen

1. Prüfungen aus Webntis exportieren und auf den Desktop legen.
2. *WebuntisNoten2Atlantis* starten.
3. Gewünschte Klassen filtern oder 10 Sekunden warten, um alle Klassen zu wählen.
4. Erzeugte SQL-Datei in Atlantis importieren.

### Prüfungen aus Webntis exportieren

1. Mit administrativer Berechtigung in Webuntis anmelden.
2. Den Pfad *Klassenbuch > Berichte* gehen.
3. Unter der Rubrik *Noten* die gewünschte Prüfungsart *Halbjahreszeugnis* wählen, bzw. *Jahreszeugnis* usw. Der Name der Prüfungsart muss mit dem Wert des Attributs *Auflösung* in der Atlantis-Schlüsselverwaltung übereinstimmen.
4. Das Icon *CSV-Ausgabe* klicken.
5. Die Datei *MarksPerLesson.csv* auf dem Desktop speichern und nach Abschluss des Übertrags von dort wieder löschen. Die Datei hat möglicherweise sehr viele Zeilen und folgenden Aufbau:

```
Datum	Name	Klasse	Fach	Prüfungsart	Note	Bemerkung	Benutzer	Schlüssel (extern)	Gesamtnote
28.09.2019	Müller Ina	HHO1	INW	Halbjahreszeugnis	2.0		admin	149565	
```
Der externe Schlüssel ist die Atlantis-ID.

### WebuntisNoten2Atlantis bedienen
1. Programm im Visual Studio selbst kompilieren und starten oder *webuntis2Atlantis.exe* herunterladen und starten.
2. Den Anweisungen folgen. Es kann beispielsweise eine Filter auf die gewünschten Klassen gesetzt werden. Das macht sicherlich Sinn, wenn es gilt erste vorsichtige Erfahrungen zu sammeln.
3. Eine Datei namens *webuntisnoten2atlantis_20190930.SQL* wird auf den Desktop gelegt un öffnet sich im Notepad. 


### Die Datei webuntisnoten2atlantis_20190930.SQL
Die Noten werden nicht unmittelbar vom Programm *WebuntisNoten2Atlantis* in die Atlantis-Datenbank zurückgeschrieben. Stattdessen wird eine SQL-Datei erzeugt, die die Noten als SQL-Anweisungen an die Datenbank übergibt.

Jede Note entspricht einer SQL-Anweisungen in einer Zeile. Die Anweisungen sind alle vollkommen unabhängig voneinander und sehen wie folgt aus:
```SQL
UPDATE noten_einzel SET s_note=3 WHERE noe_id=3760033;/*HHO1,INW,3,Müller I*/
```  
Hinter dem Semikolon steht ein kurzer Kommentar, der das Prüfen der Datei vereinfachen soll. 

Nach der sorgfältigen Prüfung der Datei kann sie in Altlantis (entsprechende Berechtigungen vorausgesetzt) unter *Funktionen>SQL-Anweisung ausführen* in die Datenbank eingelesen werden. 
Evtl. macht es Sinn zunächst alle bis auf eine SQL-Anweisung zu löschen und dann auszuführen. 
So kann zunächst bei einer einzelnen Note eines einzelnen Schülers gerüft werden, ob alles funktioniert.

### FAQ

#### Was ist mit den Umlauten passiert?
Die SQL-Anweisungen selbst enthalten niemals Umlaute. Insofern ist das unkritisch. Programmseitig steht ist das Encoding auf Default gestellt.

#### Unter welcher Lizenz steht das Programm?

GNU General Public License



stefan.baeumer@berufskolleg-borken.de 01.10.2019