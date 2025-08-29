========================================================================
       MICROSOFT FOUNDATION CLASS BIBLIOTHEK : mtexcexp
========================================================================


Diese mtexcexp-DLL hat der Klassen-Assistent für Sie erstellt. Die DLL zeigt 
die prinzipielle Anwendung der Microsoft Foundation Classes und dient Ihnen 
als Ausgangspunkt für die Erstellung Ihrer eigenen DLLs.

Diese Datei enthält die Zusammenfassung der Bestandteile aller Dateien, die 
Ihre mtexcexp-DLL bilden.

mtexcexp.dsp
    Diese Datei (Projektdatei) enthält Informationen auf Projektebene und wird zur
    Erstellung eines einzelnen Projekts oder Teilprojekts verwendet. Andere Benutzer können
    die Projektdatei (.dsp) gemeinsam nutzen, sollten aber die Makefiles lokal exportieren.

mtexcexp.h
	Dies ist die Haupt-Header-Datei für die DLL, in der die Klasse 
	CMtexcexpApp deklariert ist.

mtexcexp.cpp
	Hierbei handelt es sich um die Haupt-Quellcodedatei der DLL. Diese enthält 
	die Klasse CMtexcexpApp.


mtexcexp.rc
	Hierbei handelt es sich um eine Auflistung aller Ressourcen von Microsoft Windows, die 
	vom Programm verwendet werden. Sie enthält die Symbole, Bitmaps und Cursors, die im 
	Unterverzeichnis RES abgelegt sind. Diese Datei lässt sich direkt in Microsoft
	Visual C++ bearbeiten.

mtexcexp.clw
    Diese Datei enthält Informationen, die vom Klassen-Assistenten verwendet wird, um bestehende
    Klassen zu bearbeiten oder neue hinzuzufügen.  Der Klassen-Assistent verwendet diese Datei auch,
    um Informationen zu speichern, die zum Erstellen und Bearbeiten von Nachrichtentabellen und
    Dialogdatentabellen benötigt werden und um Prototyp-Member-Funktionen zu erstellen.

res\mtexcexp.rc2
    Diese Datei enthält Ressourcen, die nicht von Microsoft Visual C++ bearbeitet wurden.
	In diese Datei werden alle Ressourcen abgelegt, die vom Ressourcen-Editor nicht bearbeitet 
	werden können.

mtexcexp.def
    Diese Datei enthält Informationen über die DLL, die für das Ausführen unter
	Microsoft Windows erforderlich sind. Es werden Parameter definiert, wie beipielsweise
	der Name und die Beschreibung der DLL. Außerdem werden die von der DLL exportierten
	Funktionen angegeben.

/////////////////////////////////////////////////////////////////////////////
Andere Standarddateien:

StdAfx.h, StdAfx.cpp
    	Mit diesen Dateien werden vorkompilierte Header-Dateien (PCH) mit der Bezeichnung 
	mtexcexp.pch und eine vorkompilierte Typdatei mit der Bezeichnung StdAfx.obj
	erstellt.

Resource.h
    	Dies ist die Standard-Header-Datei, die neue Ressourcen-IDs definiert.
    	Microsoft Visual C++ liest und aktualisiert diese Datei.

/////////////////////////////////////////////////////////////////////////////
Weitere Hinweise:

Der Klassen-Assistent fügt "ZU ERLEDIGEN:" im Quellcode ein, um die Stellen anzuzeigen, 
an denen Sie Erweiterungen oder Anpassungen vornehmen können.

/////////////////////////////////////////////////////////////////////////////
