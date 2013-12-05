Rstyx.VBA.Excel.GeoTools
========================

Eine Sammlung von Werkzeugen für Excel 2003, die insbesondere das Arbeiten mit einfachen Listen effizienter gestalten sollen, im einzelnen:


Funktionen für beliebige Tabellen
---------------------------------
 - Standard-Fußzeile eintragen.
 - Interpolationsformel erstellen.
 - Doppelte Werte in einer Spalte markieren.
 - Strukturierung der Tabelle in Datenbereich und Kopf

Funktionen für bereits strukturierte Tabellen
---------------------------------------------
 - Projektdaten in den Tabellenkopf eintragen.
 - Formatierung des Datenbereiches.
 - Berechnungen oder/und Textmanipulationen über ganze Spalten
 - Import / Export von Daten für z.T. spezielle Anwendungen.
 - Ergebnis ist immer eine neue Excel-Tabelle
 - Als Datenquelle können dienen:
   -- die aktive Tabelle
   -- eine spezielle ASCII-Datei, für die ein Import-Modul realisiert ist.
   -- eine CSV-Datei mit passendem Kopf

Zugriff auf die Funktionen der GeoTools
---------------------------------------
 - Hauptmenü, Punkt "GeoTools"
 - Symbolleisten bzw. Toolboxen
 - Die Befehle stehen i.d.R. nur dann zur Verfügung, wenn sie sinnvoll einsetzbar sind. Ein Befehl, der eine vorhandene Tabelle bearbeitet, ist z.B. nicht verfügbar, wenn keine Tabelle aktiv ist usw. Die Beschriftungen und Tooltips von Schaltern geben den aktiven Status an.
 - Die Import-Funktionen sind abhängig vom Vorhandensein passender Formatvorlagen. Diese Vorlagen können wie jede andere Vorlage als Grundlage zum Erstellen einer neuen Datei dienen (Menü Datei | neu).

Lizenz
-------
 Dieses Programm unterliegt den Bedingungen der [MIT License] (http://opensource.org/licenses/mit-license.html).

Abhängigkeiten
--------------
 - Excel 2003
 - Microsoft Scripting Runtime
 - Windows Script Host Object Model
 - Microsoft Windows Common Controls 6.0 (MSComctl.ocx)
