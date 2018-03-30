' Das Ziel ist zu testen, wie sich DB-Einträge als Type abspeichern lassen.
' D.h.:
' - Auslesen eines Eintrags der DB in einen Type.
' - Speichern von Arrays von DB-Typs
' - Funktionen zum Konstruieren, Auslesen und Weiterverarbeiten des DB-Typs

Type Aufgabe
	AufgID As Integer
	Fachgebiet As Integer
	Author As Integer
	Schwierigkeitsgrad As Integer
	LetzteAenderung As Date
	AufgabenText As String
	LinkZuBild As String
	MC2 As String
	MC3 As String
	MC4 As String
	MC5 As String
	MC6 As String
	Loesung As String
	Auswahl As Boolean
	Position As Integer
	LoesungJN As Boolean
	IDVarianten As String
	Kurzloesung As String
	KurzloesungJN As Boolean
	AlteID As String
	MCjanein As Boolean
	PermanenteID As String
	EnglischJN As Boolean
	Bemerkung As String
	AbstNachAufg As Integer
End Type

Sub setAufgID(a As Aufgabe, AufgID As Integer)
	a.AufgID = AufgID
End Sub

Function getAufgID(a As Aufgabe) As Integer
	' diese Funktion ist sinnlos...
	getAufgID = a.AufgID
End Function

Sub delme
	Dim c as Aufgabe
	
	Print c.AufgID
	
	setAufgID(c,4)
	
	Print c.AufgID
	
	Print getAufgID(c)
	
End Sub


Sub AufgabePrint(a As Aufgabe)
	' 1.3.17, Erasmus Bieri
	' Gibt Inhalt von Aufgabe als String zurück:
	
	s = "AufgID: " & cstr(a.AufgID) & ", Permanente ID: " & aPermanenteID & ", Fachgebiet: " & cstr(a.Fachgebiet) & ", Autor: " & cstr(a.Author) & ", Schwierigkeitsgrad: " & cstr(a.Schwierigkeitsgrad) & ", LetzteAenderung: " & cstr(LetzteAenderung) & chr(13) &_
	"AufgabenText: " & a.AufgabenText & chr(13) &_
	"LinkZuBild: " & a.LinkZuBild & chr(13) &_
	"Lösung: " & a.Loesung
	
	msgbox s
End Sub

Function AufgabeByValueKopieren(ByVal a As Aufgabe) As Aufgabe
	Dim b As Aufgabe
	b.AufgID = a.AufgID
	b.Fachgebiet = a.Fachgebiet
	b.Author = a.Author
	b.Schwierigkeitsgrad = a.Schwierigkeitsgrad
	b.LetzteAenderung = a.LetzteAenderung
	b.AufgabenText = a.AufgabenText
	b.LinkZuBild = a.LinkZuBild
	b.MC2 = a.MC2
	b.MC3 = a.MC3
	b.MC4 = a.MC4
	b.MC5 = a.MC5
	b.MC6 = a.MC6
	b.Loesung = a.Loesung
	b.Auswahl = a.Auswahl
	b.Position = a.Position
	b.LoesungJN = a.LoesungJN
	b.IDVarianten = a.IDVarianten
	b.Kurzloesung = a.Kurzloesung
	b.KurzloesungJN = a.KurzloesungJN
	b.AlteID = a.AlteID
	b.MCjanein = a.MCjanein
	b.PermanenteID = a.PermanenteID
	b.EnglischJN = a.EnglischJN
	b.Bemerkung = a.Bemerkung
	b.AbstNachAufg = a.AbstNachAufg
	
	AufgabeByValueKopieren = b
End Function 

Function afAusgewaehlteAufgabeAusDatenbanklesen()
	Dim a As Aufgabe
	Dim AufgabenArray As Variant 'Wird Aufgabenarray
	AufgabenArray = Array(a)
	
	Dim DatabaseContext, DataSource, Connection, InteractionHandler, Statement, ResultSet As Object	
	'DatenbankKontext erstellen
    DatabaseContext = createUnoService("com.sun.star.sdb.DatabaseContext")
    'Datenbank kontaktieren
    DataSource = DatabaseContext.getByName("aufgdbphys")
    'Testen, ob Passwort benötigt wird
    If Not DataSource.IsPasswordRequired Then
    	Connection = DataSource.GetConnection("","")
    Else	
    	InteractionHandler = createUnoService("com.sun.star.sdb.InteractionHandler")
    	Connection = DataSource.ConnectWithCompletion(InteractionHandler)
    End If
    'Abfrage erstellen
    Statement = Connection.createStatement()
    StatementUpdate = Connection.createStatement()
	ResultSet = Statement.executeQuery("SELECT * FROM `Aufgaben` WHERE ( `Aufgaben`.`Auswahl` = True )")
	
	'Länge Resultset bestimmen:
	iEintraegeResultSet = ResultSet.columns.count
	ReDim AufgabenArray(iEintraegeResultSet - 1) As Variant
	
	iZaehler = 0
	If Not IsNull(ResultSet) Then
   		While ResultSet.next
   			'print "test: " & ResultSet.getString(14)
   			a.AufgID = cint(ResultSet.getString(1))
			a.Fachgebiet = cint(ResultSet.getString(2))
			a.Author = cint(ResultSet.getString(3))
			a.Schwierigkeitsgrad = cint(ResultSet.getString(4))
			a.LetzteAenderung = cDate(ResultSet.getString(5))
			a.AufgabenText = ResultSet.getString(6)
			a.LinkZuBild = ResultSet.getString(7)
			a.MC2 = ResultSet.getString(8)
			a.MC3 = ResultSet.getString(9)
			a.MC4 = ResultSet.getString(10)
			a.MC5 = ResultSet.getString(11)
			a.MC6 = ResultSet.getString(12)
			a.Loesung = ResultSet.getString(13)
			a.Auswahl = ResultSet.getString(14)
			a.Position = cInt(ResultSet.getString(15))
			a.LoesungJN = ResultSet.getString(16)
			a.IDVarianten = ResultSet.getString(17)
			a.Kurzloesung = ResultSet.getString(18)
			a.KurzloesungJN = ResultSet.getString(19)
			a.AlteID = ResultSet.getString(20)
			a.MCjanein = ResultSet.getString(21)
			a.PermanenteID = ResultSet.getString(22)
			a.EnglischJN = ResultSet.getString(23)
			a.Bemerkung = ResultSet.getString(24)
			a.AbstNachAufg = ResultSet.getString(25)
			
			AufgabenArray(iZaehler) = AufgabeByValueKopieren(a) 
			
			iZaehler = iZaehler + 1
   		Wend
   	Else
		MsgBox "Es gibt keine ausgewählten Aufgaben."
   		Exit Function
	End If
	
	afAusgewaehlteAufgabeAusDatenbanklesen() = AufgabenArray()
End Function

Sub AppendToArray(oData(), ByVal x)
'While accumulating data, it is convenient to add data to an array without worrying about the size
'of the array—it may not be efficient, but it is convenient. The macro in Listing 63 increases the
'size of an array by one, and then appends data to the end of the array. This technique is
'invaluable when space and complexity are more important than runtime—appending data to an
'array is similar to pushing data onto a stack.
	Dim iUB As Integer 'The upper bound of the array.
	Dim iLB As Integer 'The lower bound of the array.
	iUB = UBound(oData()) + 1
	iLB = LBound(oData())
	ReDim Preserve oData(iLB To iUB)
	oData(iUB) = x
End Sub

Function AufgabeKonstruktor (a As Aufgabe, Optional AufgID As Integer, Optional	Fachgebiet As Integer, Optional	Author As Integer, Optional Schwierigkeitsgrad As Integer, Optional LetzteAenderung As Date, Optional AufgabenText As String, Optional LinkZuBild As String, Optional MC2 As String, Optional MC3 As String, Optional MC4 As String, Optional MC5 As String, Optional MC6 As String, Optional Loesung As String, Optional Auswahl As Boolean, Optional Position As Integer, Optional LoesungJN As Boolean, Optional IDVarianten As String, Optional Kurzloesung As String, Optional KurzloesungJN As Boolean, Optional AlteID As String, Optional MCjanein As Boolean, Optional PermanenteID As String, Optional EnglischJN As Boolean, Optional Bemerkung As String, Optional AbstNachAufg As Integer)
	'Letzte Änderung: 17.2.17, Erasmus Bieri
	
	' Kontruiert eine Datenstuktur vom Typ Aufgabe
	' Gibt bei korrekten Eingaben 1, sonst 0 zurück

	Dim AufgabeKonstruktor As Boolean
	AufgabeKonstruktor = 1
	' Falls keine optional
	a.Farbe = ""
	a.Leistung = 0
	If Not(IsMissing(sFarbe)) and Left(sFarbe,5) <> "Error" Then
		Print("Setze Farbe auf: <" & sFarbe & ">")
		a.Farbe = sFarbe
    End If
    If Not(IsMissing(iLeistung)) Then
		Print("Setze Leistung auf: <" & iLeistung & ">")
		a.Leistung = iLeistung
    End If
	'Print("UB: " & ubound(sList()))
	'if(ubound
	'for i = 0 to ubound(sList())
	'	
	'next i

End Function
