Attribute VB_Name = "modXRechnung"
'==============================================================================
' modXRechnung
' Word-Add-In: erzeugt aus einer Rechnungsvorlage eine XRechnung (UBL 2.1,
' XRechnung 3.0 / EN 16931). Die Kopfdaten kommen aus Textmarken bzw. manuellen
' Eingaben, die Positionen aus einer Tabelle des Dokuments.
'==============================================================================
Option Explicit

Public Const XR_PROP As String = "XRechnung"
Public Const XR_CFG_VERSION As String = "1"
Public Const XR_SPEC_ID As String = "urn:cen.eu:en16931:2017#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0"
Public Const XR_PROFILE_ID As String = "urn:fdc:peppol.eu:2017:poacc:billing:01:1.0"
Public Const SRC_MANUAL As String = "(manuell)"
Public Const SRC_CALC As String = "(berechnet)"
Public Const SRC_NOCOL As String = "(nicht zugeordnet)"

Private Const NS_INV As String = "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
Private Const NS_CAC As String = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
Private Const NS_CBC As String = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"

' Definition eines BT-Feldes (Kopf) bzw. einer Spaltenzuordnung (Position)
' Kind: G Gruppe, T Text, D Datum, A Betrag, N Zahl/Prozent, K berechnet,
'       I Kennung, U Einheit, C USt-Kategorie
Public Type tBTDef
    id As String
    Name As String
    req As String
    kind As String
    defVal As String
End Type

Public Type tLine
    id As String
    Note As String
    Qty As Double
    Unit As String
    Net As Double
    Price As Double
    BaseQty As Double
    Allow As Double
    AllowReason As String
    cat As String
    rate As Double
    Name As String
    Desc As String
    SellerID As String
    BuyerID As String
    StdID As String
    Origin As String
    OrderLine As String
    AcctRef As String
    PStart As String
    PEnd As String
    SrcRow As Long
End Type

Public Type tTotals
    LineSum As Double       ' BT-106
    AllowSum As Double      ' BT-107
    ChargeSum As Double     ' BT-108
    NetTotal As Double      ' BT-109
    VatTotal As Double      ' BT-110
    Gross As Double         ' BT-112
    Prepaid As Double       ' BT-113
    Rounding As Double      ' BT-114
    Due As Double           ' BT-115
    AllowAmt As Double
    AllowCat As String
    AllowRate As Double
    ChargeAmt As Double
    ChargeCat As String
    ChargeRate As Double
    nGroups As Long         ' BG-23
    gCat() As String
    gRate() As Double
    gBase() As Double
    gVat() As Double
End Type

Private Type tRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Windows-API (Word 2010 und neuer, 32 und 64 Bit)
Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long

Private mD() As tBTDef
Private mDN As Long
Private mVals As Collection
Private mCfgK() As String
Private mCfgV() As String
Private mCfgN As Long
Private mDom As Object
Private mErrs As Collection
Private mCur As String

'==============================================================================
' Einstiegspunkte
'==============================================================================

' Menüband-Schaltfläche "XRechnung erzeugen"
Public Sub XR_RibbonErzeugen(control As IRibbonControl)
    XRechnungErzeugen
End Sub

' Öffnet den Dialog für das aktive Dokument (auch über Alt+F8 aufrufbar)
Public Sub XRechnungErzeugen()
    Dim f As frmXRechnung
    If Documents.Count = 0 Then
        MsgBox "Bitte zuerst die Rechnung öffnen.", vbExclamation, "XRechnung"
        Exit Sub
    End If
    On Error GoTo Fehler
    Set f = New frmXRechnung
    f.Setup ActiveDocument
    f.Show
    Unload f
    Exit Sub
Fehler:
    MsgBox "Fehler beim Erzeugen der XRechnung:" & vbCr & Err.Description, vbCritical, "XRechnung"
    On Error Resume Next
    Unload f
End Sub

' Erzeugt die XRechnung ohne Dialog mit der im Dokument gespeicherten Zuordnung.
' Rückgabe: leer bei Erfolg (ggf. Hinweise), sonst Fehlermeldung beginnend mit "FEHLER".
Public Function XRechnungExportieren(Optional ByVal zielPfad As String = "") As String
    Dim f As frmXRechnung
    On Error GoTo Fehler
    If Documents.Count = 0 Then XRechnungExportieren = "FEHLER: kein Dokument geöffnet": Exit Function
    Set f = New frmXRechnung
    f.Setup ActiveDocument
    If Len(zielPfad) = 0 Then zielPfad = XR_TargetPath(ActiveDocument, False)
    If f.ExportXml(zielPfad, False) Then
        XRechnungExportieren = f.LastMessage
    Else
        XRechnungExportieren = "FEHLER: " & f.LastMessage
    End If
    Unload f
    Exit Function
Fehler:
    XRechnungExportieren = "FEHLER: " & Err.Description
    On Error Resume Next
    Unload f
End Function

'==============================================================================
' BT-Definitionen (XRechnung 3.0, EN 16931)
'==============================================================================

Private Sub AddDef(ByVal id As String, ByVal nm As String, Optional ByVal req As String = "O", _
                   Optional ByVal kind As String = "T", Optional ByVal defVal As String = "")
    mDN = mDN + 1
    If mDN > UBound(mD) Then ReDim Preserve mD(1 To mDN + 50)
    mD(mDN).id = id
    mD(mDN).Name = nm
    mD(mDN).req = req
    mD(mDN).kind = kind
    mD(mDN).defVal = defVal
End Sub

Private Sub AddGrp(ByVal id As String, ByVal nm As String)
    AddDef id, nm, "", "G"
End Sub

' Felder der Rechnung (ohne Positionen); P = Pflicht, O = optional, C = bedingt
Public Function XR_GetHeaderDefs(defs() As tBTDef) As Long
    mDN = 0
    ReDim mD(1 To 200)
    AddGrp "BG-0", "Rechnung"
    AddDef "BT-1", "Rechnungsnummer", "P"
    AddDef "BT-2", "Rechnungsdatum", "P", "D"
    AddDef "BT-3", "Code für den Rechnungstyp (380 = Rechnung)", "P", "T", "380"
    AddDef "BT-5", "Code für die Rechnungswährung", "P", "T", "EUR"
    AddDef "BT-6", "Code für die Umsatzsteuer-Buchungswährung", "O"
    AddDef "BT-7", "Abrechnungsdatum der Umsatzsteuer", "O", "D"
    AddDef "BT-8", "Code für das Abrechnungsdatum der Umsatzsteuer", "O"
    AddDef "BT-9", "Fälligkeitsdatum der Zahlung", "C", "D"
    AddDef "BT-10", "Käuferreferenz (Leitweg-ID)", "P"
    AddDef "BT-11", "Projektreferenz", "O"
    AddDef "BT-12", "Vertragsreferenz", "O"
    AddDef "BT-13", "Bestellreferenz", "O"
    AddDef "BT-14", "Auftragsreferenz", "O"
    AddDef "BT-15", "Wareneingangsmeldungsreferenz", "O"
    AddDef "BT-16", "Versandanzeigereferenz (Lieferschein)", "O"
    AddDef "BT-17", "Ausschreibungs- oder Losreferenz", "O"
    AddDef "BT-18", "Kennung des in Rechnung gestellten Objekts", "O"
    AddDef "BT-19", "Buchungsreferenz des Käufers", "O"
    AddDef "BT-20", "Zahlungsbedingungen", "C"
    AddGrp "BG-1", "Rechnungsanmerkung"
    AddDef "BT-21", "Betreff der Rechnungsanmerkung (Code)", "O"
    AddDef "BT-22", "Rechnungsanmerkung", "O"
    AddGrp "BG-2", "Prozesssteuerung"
    AddDef "BT-23", "Geschäftsprozesstyp", "O", "T", XR_PROFILE_ID
    AddDef "BT-24", "Spezifikationskennung", "P", "T", XR_SPEC_ID
    AddGrp "BG-3", "Vorausgegangene Rechnung"
    AddDef "BT-25", "Referenz auf die vorausgegangene Rechnung", "C"
    AddDef "BT-26", "Rechnungsdatum der vorausgegangenen Rechnung", "O", "D"
    AddGrp "BG-4", "Verkäufer"
    AddDef "BT-27", "Name des Verkäufers", "P"
    AddDef "BT-28", "Handelsname des Verkäufers", "O"
    AddDef "BT-29", "Kennung des Verkäufers", "C"
    AddDef "BT-30", "Kennung der rechtlichen Registrierung (Register-Nr.)", "C"
    AddDef "BT-31", "Umsatzsteuer-Identifikationsnummer des Verkäufers", "C"
    AddDef "BT-32", "Steuernummer des Verkäufers", "C"
    AddDef "BT-33", "Weitere rechtliche Informationen zum Verkäufer", "O"
    AddDef "BT-34", "Elektronische Adresse des Verkäufers", "P"
    AddGrp "BG-5", "Postanschrift des Verkäufers"
    AddDef "BT-35", "Anschriftenzeile 1 (Straße)", "O"
    AddDef "BT-36", "Anschriftenzeile 2", "O"
    AddDef "BT-162", "Anschriftenzeile 3", "O"
    AddDef "BT-37", "Stadt", "P"
    AddDef "BT-38", "Postleitzahl", "P"
    AddDef "BT-39", "Region (Bundesland)", "O"
    AddDef "BT-40", "Ländercode", "P", "T", "DE"
    AddGrp "BG-6", "Kontaktstelle des Verkäufers"
    AddDef "BT-41", "Kontaktstelle des Verkäufers", "P"
    AddDef "BT-42", "Telefonnummer der Kontaktstelle", "P"
    AddDef "BT-43", "E-Mail-Adresse der Kontaktstelle", "P"
    AddGrp "BG-7", "Käufer"
    AddDef "BT-44", "Name des Käufers", "P"
    AddDef "BT-45", "Handelsname des Käufers", "O"
    AddDef "BT-46", "Kennung des Käufers (Kundennummer)", "O"
    AddDef "BT-47", "Kennung der rechtlichen Registrierung des Käufers", "O"
    AddDef "BT-48", "Umsatzsteuer-Identifikationsnummer des Käufers", "C"
    AddDef "BT-49", "Elektronische Adresse des Käufers", "P"
    AddGrp "BG-8", "Postanschrift des Käufers"
    AddDef "BT-50", "Anschriftenzeile 1 (Straße)", "O"
    AddDef "BT-51", "Anschriftenzeile 2", "O"
    AddDef "BT-163", "Anschriftenzeile 3", "O"
    AddDef "BT-52", "Stadt", "P"
    AddDef "BT-53", "Postleitzahl", "P"
    AddDef "BT-54", "Region (Bundesland)", "O"
    AddDef "BT-55", "Ländercode", "P", "T", "DE"
    AddGrp "BG-9", "Kontaktstelle des Käufers"
    AddDef "BT-56", "Kontaktstelle des Käufers", "O"
    AddDef "BT-57", "Telefonnummer der Kontaktstelle", "O"
    AddDef "BT-58", "E-Mail-Adresse der Kontaktstelle", "O"
    AddGrp "BG-10", "Zahlungsempfänger (falls abweichend)"
    AddDef "BT-59", "Name des Zahlungsempfängers", "C"
    AddDef "BT-60", "Kennung des Zahlungsempfängers", "O"
    AddDef "BT-61", "Kennung der rechtlichen Registrierung", "O"
    AddGrp "BG-11", "Steuervertreter des Verkäufers"
    AddDef "BT-62", "Name des Steuervertreters", "C"
    AddDef "BT-63", "Umsatzsteuer-Identifikationsnummer des Steuervertreters", "C"
    AddGrp "BG-12", "Postanschrift des Steuervertreters"
    AddDef "BT-64", "Anschriftenzeile 1 (Straße)", "O"
    AddDef "BT-65", "Anschriftenzeile 2", "O"
    AddDef "BT-164", "Anschriftenzeile 3", "O"
    AddDef "BT-66", "Stadt", "O"
    AddDef "BT-67", "Postleitzahl", "O"
    AddDef "BT-68", "Region (Bundesland)", "O"
    AddDef "BT-69", "Ländercode", "C"
    AddGrp "BG-13", "Lieferinformationen"
    AddDef "BT-70", "Name des Empfängers der Lieferung", "O"
    AddDef "BT-71", "Kennung des Lieferorts", "O"
    AddDef "BT-72", "Tatsächliches Lieferdatum", "O", "D"
    AddGrp "BG-14", "Rechnungszeitraum"
    AddDef "BT-73", "Anfangsdatum des Rechnungszeitraums", "O", "D"
    AddDef "BT-74", "Enddatum des Rechnungszeitraums", "O", "D"
    AddGrp "BG-15", "Lieferanschrift"
    AddDef "BT-75", "Anschriftenzeile 1 (Straße)", "O"
    AddDef "BT-76", "Anschriftenzeile 2", "O"
    AddDef "BT-165", "Anschriftenzeile 3", "O"
    AddDef "BT-77", "Stadt", "C"
    AddDef "BT-78", "Postleitzahl", "C"
    AddDef "BT-79", "Region (Bundesland)", "O"
    AddDef "BT-80", "Ländercode", "C"
    AddGrp "BG-16", "Zahlungsanweisungen"
    AddDef "BT-81", "Code für die Zahlungsart (58 = SEPA-Überweisung)", "P", "T", "58"
    AddDef "BT-82", "Zahlungsart (Text)", "O"
    AddDef "BT-83", "Verwendungszweck", "O"
    AddGrp "BG-17", "Überweisung"
    AddDef "BT-84", "Kennung des Zahlungskontos (IBAN)", "C"
    AddDef "BT-85", "Name des Zahlungskontos (Kontoinhaber)", "O"
    AddDef "BT-86", "Kennung des Zahlungsdienstleisters (BIC)", "O"
    AddGrp "BG-18", "Zahlungskarte"
    AddDef "BT-87", "Kartennummer", "C"
    AddDef "BT-88", "Name des Karteninhabers", "O"
    AddGrp "BG-19", "Lastschrift"
    AddDef "BT-89", "Kennung der Mandatsreferenz", "C"
    AddDef "BT-90", "Gläubiger-Identifikationsnummer", "C"
    AddDef "BT-91", "Kennung des zu belastenden Kontos (IBAN)", "C"
    AddGrp "BG-20", "Abzüge auf Dokumentenebene (Nachlass)"
    AddDef "BT-92", "Betrag des Nachlasses", "C", "A"
    AddDef "BT-93", "Grundbetrag des Nachlasses", "O", "A"
    AddDef "BT-94", "Prozentsatz des Nachlasses", "O", "N"
    AddDef "BT-95", "Code der USt-Kategorie des Nachlasses", "C"
    AddDef "BT-96", "USt-Satz des Nachlasses", "C", "N"
    AddDef "BT-97", "Grund für den Nachlass", "C"
    AddDef "BT-98", "Code für den Grund des Nachlasses", "C"
    AddGrp "BG-21", "Zuschläge auf Dokumentenebene (z.B. Versandkosten)"
    AddDef "BT-99", "Betrag des Zuschlags", "C", "A"
    AddDef "BT-100", "Grundbetrag des Zuschlags", "O", "A"
    AddDef "BT-101", "Prozentsatz des Zuschlags", "O", "N"
    AddDef "BT-102", "Code der USt-Kategorie des Zuschlags", "C"
    AddDef "BT-103", "USt-Satz des Zuschlags", "C", "N"
    AddDef "BT-104", "Grund für den Zuschlag", "C"
    AddDef "BT-105", "Code für den Grund des Zuschlags", "C"
    AddGrp "BG-22", "Gesamtbeträge"
    AddDef "BT-106", "Summe der Nettobeträge aller Positionen", "P", "K"
    AddDef "BT-107", "Summe der Nachlässe auf Dokumentenebene", "C", "K"
    AddDef "BT-108", "Summe der Zuschläge auf Dokumentenebene", "C", "K"
    AddDef "BT-109", "Gesamtbetrag ohne Umsatzsteuer", "P", "K"
    AddDef "BT-110", "Gesamtbetrag der Umsatzsteuer", "C", "K"
    AddDef "BT-111", "Umsatzsteuer in Buchungswährung", "C", "A"
    AddDef "BT-112", "Gesamtbetrag einschließlich Umsatzsteuer", "P", "K"
    AddDef "BT-113", "Vorauszahlungsbetrag", "O", "A"
    AddDef "BT-114", "Rundungsbetrag", "O", "A"
    AddDef "BT-115", "Fälliger Zahlungsbetrag", "P", "K"
    AddGrp "BG-23", "Aufschlüsselung der Umsatzsteuer"
    AddDef "BT-116", "Steuerbasisbetrag der Kategorie", "P", "K"
    AddDef "BT-117", "Steuerbetrag der Kategorie", "P", "K"
    AddDef "BT-118", "Code der Umsatzsteuerkategorie", "P", "K"
    AddDef "BT-119", "Umsatzsteuersatz der Kategorie", "P", "K"
    AddDef "BT-120", "Befreiungsgrund (Text)", "C"
    AddDef "BT-121", "Code für den Befreiungsgrund", "C"
    ReDim Preserve mD(1 To mDN)
    defs = mD
    XR_GetHeaderDefs = mDN
End Function

' Felder einer Rechnungsposition (BG-25), die aus Tabellenspalten kommen
Public Function XR_GetLineDefs(defs() As tBTDef) As Long
    mDN = 0
    ReDim mD(1 To 30)
    AddDef "BT-126", "Positionsnummer", "P", "I"
    AddDef "BT-155", "Artikelnummer des Verkäufers", "O", "I"
    AddDef "BT-156", "Artikelnummer des Käufers", "O", "I"
    AddDef "BT-157", "Artikelkennung (GTIN)", "O", "I"
    AddDef "BT-153", "Artikelname", "P", "T"
    AddDef "BT-154", "Artikelbeschreibung", "O", "T"
    AddDef "BT-129", "Menge", "P", "N"
    AddDef "BT-130", "Einheit der Menge", "P", "U", "H87"
    AddDef "BT-146", "Nettopreis (Einzelpreis)", "P", "A"
    AddDef "BT-149", "Basismenge zum Preis", "O", "N"
    AddDef "BT-136", "Betrag des Positionsnachlasses", "O", "A"
    AddDef "BT-139", "Grund des Positionsnachlasses", "C", "T"
    AddDef "BT-131", "Nettobetrag der Position", "P", "A"
    AddDef "BT-151", "USt-Kategorie (S, Z, E, AE, K, G, O)", "P", "C", "S"
    AddDef "BT-152", "USt-Satz in %", "C", "N", "19"
    AddDef "BT-127", "Freitext zur Position", "O", "T"
    AddDef "BT-132", "Referenz zur Bestellposition", "O", "I"
    AddDef "BT-133", "Buchungsreferenz des Käufers", "O", "T"
    AddDef "BT-134", "Anfangsdatum Abrechnungszeitraum", "O", "D"
    AddDef "BT-135", "Enddatum Abrechnungszeitraum", "O", "D"
    AddDef "BT-159", "Ursprungsland des Artikels", "O", "T"
    ReDim Preserve mD(1 To mDN)
    defs = mD
    XR_GetLineDefs = mDN
End Function

'==============================================================================
' Werte-Speicher (vom Dialog befüllt, für Berechnung und XML)
'==============================================================================

Public Sub XR_ClearValues()
    Set mVals = New Collection
End Sub

Public Sub XR_SetValue(ByVal id As String, ByVal v As String)
    If mVals Is Nothing Then Set mVals = New Collection
    On Error Resume Next
    mVals.Remove id
    On Error GoTo 0
    mVals.Add v, id
End Sub

Public Function BTV(ByVal id As String) As String
    If mVals Is Nothing Then Exit Function
    On Error Resume Next
    BTV = mVals(id)
End Function

Private Function TX(ByVal id As String) As String
    TX = CleanText(BTV(id))
End Function

Private Function AnyV(ParamArray ids() As Variant) As Boolean
    Dim v As Variant
    For Each v In ids
        If Len(TX(CStr(v))) > 0 Then AnyV = True: Exit Function
    Next
End Function

Private Function NumV(ByVal id As String) As Double
    Dim ok As Boolean
    NumV = XR_ParseNum(TX(id), ok)
    If Not ok Then NumV = 0
End Function

'==============================================================================
' Text- und Zahlenfunktionen
'==============================================================================

Public Function CleanText(ByVal s As String) As String
    Dim i As Long, ch As Long, res As String
    s = Replace(s, Chr(13) & Chr(7), " ")
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, Chr(30), "-")
    s = Replace(s, Chr(31), "")
    s = Replace(s, ChrW(8203), "")
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch >= 0 And ch < 32 Then
            res = res & " "
        Else
            res = res & Mid$(s, i, 1)
        End If
    Next
    Do While InStr(res, "  ") > 0
        res = Replace(res, "  ", " ")
    Loop
    CleanText = Trim$(res)
End Function

' Text eines Bereichs mit Feldergebnissen statt Feldcodes (z.B. MERGEFIELD, TIME)
Public Function XR_RangeText(ByVal rng As Object) As String
    Dim r As Object
    Set r = rng.Duplicate
    On Error Resume Next
    r.TextRetrievalMode.IncludeFieldCodes = False
    r.TextRetrievalMode.IncludeHiddenText = False
    On Error GoTo 0
    XR_RangeText = r.Text
End Function

' Platzhalter wie "n/a" oder "-" in der Vorlage gelten als leeres Feld
Public Function XR_IsPlaceholder(ByVal s As String) As Boolean
    Select Case LCase$(CleanText(s))
        Case "n/a", "n.a.", "-", "--", ChrW(8211), ChrW(8212), "./.", "entfällt", "k.a.", "keine"
            XR_IsPlaceholder = True
    End Select
End Function

' Zerlegt einen Zelltext in nicht-leere Zeilen (Absätze / Zeilenumbrüche)
Public Function XR_CellLines(ByVal s As String) As Variant
    Dim parts() As String, outA() As String, i As Long, n As Long, t As String
    s = Replace(s, Chr(13) & Chr(7), "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(11), vbCr)
    s = Replace(s, vbLf, vbCr)
    parts = Split(s, vbCr)
    ReDim outA(0 To UBound(parts) + 1)
    For i = 0 To UBound(parts)
        t = CleanText(parts(i))
        If Len(t) > 0 Then outA(n) = t: n = n + 1
    Next
    If n = 0 Then
        XR_CellLines = Array()
    Else
        ReDim Preserve outA(0 To n - 1)
        XR_CellLines = outA
    End If
End Function

' Wählt Zeilen eines Zelltextes aus: "" = alle, "2" = 2. Zeile, "2-3", "3-" ab 3.
Public Function XR_PickLines(ByVal cellStr As String, ByVal spec As String) As String
    Dim arr() As String, a As Long, b As Long, i As Long, p As Long, res As String
    If Len(cellStr) = 0 Then Exit Function
    arr = Split(cellStr, vbLf)
    spec = Replace(Trim$(spec), " ", "")
    If Len(spec) = 0 Then
        a = 1: b = UBound(arr) + 1
    Else
        p = InStr(spec, "-")
        If p = 0 Then
            a = Val(spec): b = a
        Else
            a = Val(Left$(spec, p - 1))
            If a < 1 Then a = 1
            If p = Len(spec) Then b = UBound(arr) + 1 Else b = Val(Mid$(spec, p + 1))
        End If
    End If
    If b > UBound(arr) + 1 Then b = UBound(arr) + 1
    For i = a To b
        If i >= 1 Then
            If Len(res) > 0 Then res = res & " "
            res = res & arr(i - 1)
        End If
    Next
    XR_PickLines = res
End Function

' Liest die erste Zahl aus einem Text (deutsches und englisches Format)
Public Function XR_ParseNum(ByVal s As String, Optional ByRef ok As Boolean) As Double
    Dim i As Long, st As Long, en As Long, ch As String, t As String, pre As String
    Dim neg As Boolean, pc As Long, pd As Long
    ok = False
    s = CleanText(s)
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then st = i: Exit For
    Next
    If st = 0 Then Exit Function
    pre = Trim$(Left$(s, st - 1))
    If Right$(pre, 1) = "-" Or Right$(pre, 1) = ChrW(8722) Then neg = True
    en = st
    Do While en <= Len(s)
        ch = Mid$(s, en, 1)
        If ch Like "#" Or ch = "," Or ch = "." Or ch = "'" Then en = en + 1 Else Exit Do
    Loop
    t = Replace(Mid$(s, st, en - st), "'", "")
    If Mid$(s, en, 1) = "-" And en = Len(s) Then neg = True
    Do While Right$(t, 1) = "," Or Right$(t, 1) = "."
        t = Left$(t, Len(t) - 1)
    Loop
    pc = InStrRev(t, ",")
    pd = InStrRev(t, ".")
    If pc > 0 And pd > 0 Then
        If pc > pd Then
            t = Replace(t, ".", ""): t = Replace(t, ",", ".")
        Else
            t = Replace(t, ",", "")
        End If
    ElseIf pc > 0 Then
        If Len(t) - Len(Replace(t, ",", "")) > 1 Then t = Replace(t, ",", "") Else t = Replace(t, ",", ".")
    ElseIf pd > 0 Then
        If Len(t) - Len(Replace(t, ".", "")) > 1 Then
            t = Replace(t, ".", "")
        ElseIf Len(t) - pd = 3 Then
            t = Replace(t, ".", "")     ' 1.234 = Tausenderpunkt
        End If
    End If
    XR_ParseNum = Val(t)
    If neg Then XR_ParseNum = -XR_ParseNum
    ok = True
End Function

' Kaufmännisches Runden
Public Function XR_RoundHU(ByVal x As Double, ByVal decs As Long) As Double
    Dim f As Variant
    f = CDec(10 ^ decs)
    XR_RoundHU = CDbl(Sgn(x) * Int(Abs(CDec(x)) * f + CDec(0.5)) / f)
End Function

Public Function XR_FmtDec(ByVal x As Double, ByVal decs As Long) As String
    Dim s As String
    x = XR_RoundHU(x, decs)
    If decs > 0 Then s = Format$(x, "0." & String$(decs, "0")) Else s = Format$(x, "0")
    XR_FmtDec = Replace(s, ",", ".")
End Function

Public Function XR_FmtAmt(ByVal x As Double) As String
    XR_FmtAmt = XR_FmtDec(x, 2)
End Function

Public Function XR_FmtQty(ByVal x As Double) As String
    Dim s As String
    s = XR_FmtDec(x, 4)
    Do While Right$(s, 1) = "0" And InStr(s, ".") > 0
        s = Left$(s, Len(s) - 1)
    Loop
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    XR_FmtQty = s
End Function

Public Function XR_FmtPrice(ByVal x As Double) As String
    Dim s As String
    s = XR_FmtDec(x, 4)
    Do While Right$(s, 1) = "0" And Len(s) - InStr(s, ".") > 2
        s = Left$(s, Len(s) - 1)
    Loop
    XR_FmtPrice = s
End Function

' Datum in ISO-Format (JJJJ-MM-TT)
Public Function XR_IsoDate(ByVal s As String, Optional ByRef ok As Boolean) As String
    Dim t As String, i As Long, st As Long, ch As String, p() As String
    Dim d As Long, m As Long, y As Long, dt As Date
    ok = False
    s = CleanText(s)
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then st = i: Exit For
    Next
    If st = 0 Then GoTo TryDate
    For i = st To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Or ch = "." Or ch = "-" Or ch = "/" Then t = t & ch Else Exit For
    Next
    If t Like "####-##-##*" Then
        y = Val(Left$(t, 4)): m = Val(Mid$(t, 6, 2)): d = Val(Mid$(t, 9, 2))
    ElseIf t Like "########" Then
        y = Val(Left$(t, 4)): m = Val(Mid$(t, 5, 2)): d = Val(Mid$(t, 7, 2))
    Else
        p = Split(Replace(Replace(t, "/", "."), "-", "."), ".")
        If UBound(p) < 2 Then GoTo TryDate
        If Len(p(2)) = 0 Then GoTo TryDate
        d = Val(p(0)): m = Val(p(1)): y = Val(p(2))
        If y < 100 Then y = y + 2000
    End If
    If m >= 1 And m <= 12 And d >= 1 And d <= 31 And y >= 1900 And y <= 2200 Then
        If Day(DateSerial(y, m, d)) = d Then
            XR_IsoDate = CStr(y) & "-" & Right$("0" & m, 2) & "-" & Right$("0" & d, 2)
            ok = True
        End If
    End If
    Exit Function
TryDate:
    If IsDate(s) Then
        dt = CDate(s)
        XR_IsoDate = Format$(dt, "yyyy\-mm\-dd")
        ok = True
    End If
End Function

' Einheiten nach UN/ECE Recommendation 20/21
Public Function XR_UnitCode(ByVal s As String, ByVal defCode As String) As String
    Dim u As String, orig As String
    orig = CleanText(s)
    u = LCase$(Replace(orig, ".", ""))
    u = Trim$(u)
    Select Case u
        Case "": XR_UnitCode = defCode
        Case "st", "stk", "stck", "stück", "stueck", "pc", "pcs", "piece", "x": XR_UnitCode = "H87"
        Case "std", "h", "stunde", "stunden", "hr", "hrs": XR_UnitCode = "HUR"
        Case "min", "minute", "minuten": XR_UnitCode = "MIN"
        Case "tag", "tage", "tg", "d": XR_UnitCode = "DAY"
        Case "wo", "woche", "wochen": XR_UnitCode = "WEE"
        Case "mon", "monat", "monate", "mo": XR_UnitCode = "MON"
        Case "jahr", "jahre": XR_UnitCode = "ANN"
        Case "kg": XR_UnitCode = "KGM"
        Case "g", "gr": XR_UnitCode = "GRM"
        Case "t", "to", "tonne", "tonnen": XR_UnitCode = "TNE"
        Case "m", "lfm", "meter": XR_UnitCode = "MTR"
        Case "km": XR_UnitCode = "KMT"
        Case "cm": XR_UnitCode = "CMT"
        Case "mm": XR_UnitCode = "MMT"
        Case "m2", "m²", "qm": XR_UnitCode = "MTK"
        Case "m3", "m³", "cbm": XR_UnitCode = "MTQ"
        Case "l", "ltr", "liter": XR_UnitCode = "LTR"
        Case "ml": XR_UnitCode = "MLT"
        Case "kwh": XR_UnitCode = "KWH"
        Case "psch", "pschl", "pausch", "pauschal", "pau", "ls": XR_UnitCode = "LS"
        Case "satz", "set": XR_UnitCode = "SET"
        Case "paar", "pr": XR_UnitCode = "PR"
        Case "pck", "pkt", "paket", "packung", "pack": XR_UnitCode = "XPK"
        Case "rolle", "rol": XR_UnitCode = "XRO"
        Case "karton", "ktn", "krt": XR_UnitCode = "XCT"
        Case "dose", "ds": XR_UnitCode = "XCX"
        Case "fl", "flasche": XR_UnitCode = "XBO"
        Case "pal", "palette": XR_UnitCode = "XPX"
        Case Else
            If (orig Like "[A-Z][A-Z0-9]" Or orig Like "[A-Z][A-Z0-9][A-Z0-9]") Then
                XR_UnitCode = orig
            Else
                XR_UnitCode = defCode
            End If
    End Select
End Function

Public Function XR_NormCountry(ByVal s As String) As String
    Dim u As String
    u = UCase$(Replace(CleanText(s), ".", ""))
    Select Case u
        Case "D", "DEU", "DEUTSCHLAND", "GERMANY", "BRD": XR_NormCountry = "DE"
        Case "A", "AUT", "ÖSTERREICH", "OESTERREICH", "AUSTRIA": XR_NormCountry = "AT"
        Case "SCHWEIZ", "SWITZERLAND", "CHE": XR_NormCountry = "CH"
        Case "F", "FRANKREICH", "FRANCE": XR_NormCountry = "FR"
        Case "NIEDERLANDE", "NETHERLANDS": XR_NormCountry = "NL"
        Case "B", "BELGIEN": XR_NormCountry = "BE"
        Case "L", "LUXEMBURG": XR_NormCountry = "LU"
        Case "I", "ITALIEN": XR_NormCountry = "IT"
        Case "POLEN": XR_NormCountry = "PL"
        Case "TSCHECHIEN": XR_NormCountry = "CZ"
        Case "DÄNEMARK", "DAENEMARK": XR_NormCountry = "DK"
        Case Else: XR_NormCountry = u
    End Select
End Function

Public Function XR_NormCurrency(ByVal s As String) As String
    Dim u As String
    u = UCase$(CleanText(s))
    If u = "€" Or u = "EURO" Then u = "EUR"
    XR_NormCurrency = u
End Function

Public Function XR_NormInvType(ByVal s As String) As String
    Dim u As String
    u = UCase$(CleanText(s))
    If u Like "###" Then XR_NormInvType = u: Exit Function
    If InStr(u, "KORREKTUR") > 0 Or InStr(u, "STORNO") > 0 Then
        XR_NormInvType = "384"
    ElseIf InStr(u, "SCHLUSSRECHNUNG") > 0 Then
        XR_NormInvType = "380"
    ElseIf InStr(u, "TEILRECHNUNG") > 0 Or InStr(u, "ABSCHLAG") > 0 Then
        XR_NormInvType = "326"
    ElseIf InStr(u, "GUTSCHRIFT") > 0 Then
        XR_NormInvType = "381"
    ElseIf InStr(u, "RECHNUNG") > 0 Then
        XR_NormInvType = "380"
    Else
        XR_NormInvType = u
    End If
End Function

Public Function XR_NormPayMeans(ByVal s As String) As String
    Dim u As String
    u = UCase$(CleanText(s))
    If u Like "#*" Then XR_NormPayMeans = Trim$(Str$(Val(u))): Exit Function
    If InStr(u, "LASTSCHRIFT") > 0 Then
        XR_NormPayMeans = "59"
    ElseIf InStr(u, "ÜBERWEISUNG") > 0 Or InStr(u, "UEBERWEISUNG") > 0 Then
        XR_NormPayMeans = "58"
    ElseIf InStr(u, "KARTE") > 0 Then
        XR_NormPayMeans = "54"
    Else
        XR_NormPayMeans = u
    End If
End Function

Public Function XR_NormCat(ByVal s As String) As String
    Dim u As String
    u = UCase$(CleanText(s))
    Select Case u
        Case "S", "Z", "E", "AE", "K", "G", "O", "L", "M", "": XR_NormCat = u
        Case "STANDARD", "NORMAL", "REGULÄR", "ERMÄSSIGT", "ERMÄßIGT": XR_NormCat = "S"
        Case Else
            If InStr(u, "REVERSE") > 0 Or InStr(u, "13B") > 0 Then
                XR_NormCat = "AE"
            ElseIf InStr(u, "INNERGEMEINSCHAFT") > 0 Then
                XR_NormCat = "K"
            ElseIf InStr(u, "AUSFUHR") > 0 Or InStr(u, "EXPORT") > 0 Then
                XR_NormCat = "G"
            ElseIf InStr(u, "NICHT STEUERBAR") > 0 Then
                XR_NormCat = "O"
            ElseIf InStr(u, "STEUERFREI") > 0 Or InStr(u, "BEFREIT") > 0 Then
                XR_NormCat = "E"
            Else
                XR_NormCat = u
            End If
    End Select
End Function

' Kennungen wie USt-IdNr., IBAN, BIC ohne Leerzeichen
Public Function XR_NormID(ByVal s As String) As String
    XR_NormID = UCase$(Replace(CleanText(s), " ", ""))
End Function

Public Function XR_NormPostcode(ByVal s As String) As String
    s = CleanText(s)
    If s Like "[A-Za-z]-#*" Then
        s = Mid$(s, 3)
    ElseIf s Like "[A-Za-z][A-Za-z]-#*" Then
        s = Mid$(s, 4)
    End If
    XR_NormPostcode = s
End Function

' Entfernt Bezeichner wie "Ihre Artikelnummer: " vor Kennungen
Public Function XR_StripLabel(ByVal s As String) As String
    Dim p As Long
    s = CleanText(s)
    p = InStrRev(s, ":")
    If p > 0 And p < Len(s) Then s = Trim$(Mid$(s, p + 1))
    XR_StripLabel = s
End Function

Private Function IsExemptCat(ByVal cat As String) As Boolean
    Select Case cat
        Case "E", "AE", "K", "G", "O": IsExemptCat = True
    End Select
End Function

Private Function CountDigits(ByVal s As String) As Long
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then CountDigits = CountDigits + 1
    Next
End Function

Private Function IbanOk(ByVal iban As String) As Boolean
    Dim s As String, i As Long, j As Long, ch As String, d As String, r As Long
    If Len(iban) < 15 Then Exit Function
    s = Mid$(iban, 5) & Left$(iban, 4)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            d = ch
        ElseIf ch Like "[A-Z]" Then
            d = CStr(Asc(ch) - 55)
        Else
            Exit Function
        End If
        For j = 1 To Len(d)
            r = (r * 10 + CLng(Mid$(d, j, 1))) Mod 97
        Next
    Next
    IbanOk = (r = 1)
End Function

' Sortierschlüssel, damit BT_2 vor BT_10 einsortiert wird
Public Function XR_NatKey(ByVal s As String) As String
    Dim i As Long, ch As String, num As String, res As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            num = num & ch
        Else
            If Len(num) > 0 Then res = res & Right$("0000000000" & num, 10): num = ""
            res = res & LCase$(ch)
        End If
    Next
    If Len(num) > 0 Then res = res & Right$("0000000000" & num, 10)
    XR_NatKey = res
End Function

Public Function XR_JoinCol(ByVal col As Collection, ByVal maxItems As Long) As String
    Dim i As Long, s As String
    For i = 1 To col.Count
        If i > maxItems Then
            s = s & vbCr & "... (" & (col.Count - maxItems) & " weitere)"
            Exit For
        End If
        If Len(s) > 0 Then s = s & vbCr
        s = s & "- " & col(i)
    Next
    XR_JoinCol = s
End Function

'==============================================================================
' Konfiguration als benutzerdefinierte Dokumenteigenschaft "XRechnung"
' (Word begrenzt Texteigenschaften auf 255 Zeichen, daher ggf. Fortsetzung
'  in XRechnung_2, XRechnung_3, ...)
'==============================================================================

Public Function XR_Esc(ByVal s As String) As String
    s = Replace(s, "%", "%25")
    s = Replace(s, ";", "%3B")
    s = Replace(s, "=", "%3D")
    s = Replace(s, "|", "%7C")
    s = Replace(s, "+", "%2B")
    s = Replace(s, vbCr, "%0D")
    s = Replace(s, vbLf, "%0A")
    s = Replace(s, " ", "+")
    XR_Esc = s
End Function

Public Function XR_Unesc(ByVal s As String) As String
    s = Replace(s, "+", " ")
    s = Replace(s, "%0A", vbLf)
    s = Replace(s, "%0D", vbCr)
    s = Replace(s, "%2B", "+")
    s = Replace(s, "%7C", "|")
    s = Replace(s, "%3D", "=")
    s = Replace(s, "%3B", ";")
    s = Replace(s, "%25", "%")
    XR_Unesc = s
End Function

Public Function XR_ReadConfig(ByVal doc As Object) As String
    Dim s As String, i As Long, v As Variant
    On Error Resume Next
    v = doc.CustomDocumentProperties(XR_PROP).Value
    If Err.Number <> 0 Then Exit Function
    s = CStr(v)
    For i = 2 To 999
        Err.Clear
        v = doc.CustomDocumentProperties(XR_PROP & "_" & i).Value
        If Err.Number <> 0 Then Exit For
        s = s & CStr(v)
    Next
    XR_ReadConfig = s
End Function

Public Sub XR_WriteConfig(ByVal doc As Object, ByVal cfg As String)
    Const CHUNK As Long = 240
    Dim props As Object, i As Long, n As Long, nm As String
    Set props = doc.CustomDocumentProperties
    XR_DeleteConfig doc
    n = (Len(cfg) + CHUNK - 1) \ CHUNK
    For i = 1 To n
        If i = 1 Then nm = XR_PROP Else nm = XR_PROP & "_" & i
        props.Add nm, False, 4, Mid$(cfg, (i - 1) * CHUNK + 1, CHUNK)
    Next
    doc.Saved = False
End Sub

Public Sub XR_DeleteConfig(ByVal doc As Object)
    Dim i As Long
    On Error Resume Next
    doc.CustomDocumentProperties(XR_PROP).Delete
    For i = 2 To 999
        Err.Clear
        doc.CustomDocumentProperties(XR_PROP & "_" & i).Delete
        If Err.Number <> 0 Then Exit For
    Next
End Sub

Public Function XR_ParseConfig(ByVal cfg As String) As Boolean
    Dim parts() As String, i As Long, p As Long
    mCfgN = 0
    ReDim mCfgK(0 To 0)
    ReDim mCfgV(0 To 0)
    If Left$(cfg, 2) <> "v=" Then Exit Function
    parts = Split(cfg, ";")
    ReDim mCfgK(0 To UBound(parts))
    ReDim mCfgV(0 To UBound(parts))
    For i = 0 To UBound(parts)
        p = InStr(parts(i), "=")
        If p > 1 Then
            mCfgK(mCfgN) = Left$(parts(i), p - 1)
            mCfgV(mCfgN) = Mid$(parts(i), p + 1)
            mCfgN = mCfgN + 1
        End If
    Next
    XR_ParseConfig = True
End Function

' Liefert den (noch maskierten) Wert eines Konfigurationsschlüssels
Public Function XR_CfgGet(ByVal key As String, Optional ByRef found As Boolean) As String
    Dim i As Long
    found = False
    For i = 0 To mCfgN - 1
        If StrComp(mCfgK(i), key, vbTextCompare) = 0 Then
            XR_CfgGet = mCfgV(i)
            found = True
            Exit Function
        End If
    Next
End Function

'==============================================================================
' Tabellen im Dokument
'==============================================================================

Public Sub XR_CollectTables(ByVal tbls As Object, ByVal prefix As String, _
                            paths() As String, labels() As String, n As Long)
    Dim i As Long, t As Object, p As String
    For i = 1 To tbls.Count
        Set t = tbls(i)
        If Len(prefix) = 0 Then p = CStr(i) Else p = prefix & "." & i
        n = n + 1
        ReDim Preserve paths(1 To n)
        ReDim Preserve labels(1 To n)
        paths(n) = p
        labels(n) = "Tabelle " & p & " (" & t.Rows.Count & " Zeilen): " & XR_TableHead(t, 90)
        If t.Tables.Count > 0 Then XR_CollectTables t.Tables, p, paths, labels, n
    Next
End Sub

Public Function XR_GetTable(ByVal doc As Object, ByVal path As String) As Object
    Dim parts() As String, i As Long, t As Object
    On Error GoTo Fail
    parts = Split(path, ".")
    Set t = doc.Tables(CLng(parts(0)))
    For i = 1 To UBound(parts)
        Set t = t.Tables(CLng(parts(i)))
    Next
    Set XR_GetTable = t
Fail:
End Function

Public Function XR_TableHead(ByVal t As Object, ByVal maxLen As Long) As String
    Dim c As Object, s As String, lvl As Long
    On Error Resume Next
    lvl = t.NestingLevel
    For Each c In t.Range.Cells
        If c.NestingLevel = lvl Then
            If c.RowIndex > 1 Then Exit For
            If Len(s) > 0 Then s = s & " | "
            s = s & CleanText(XR_RangeText(c.Range))
            If Len(s) > maxLen Then Exit For
        End If
    Next
    XR_TableHead = Left$(s, maxLen)
End Function

' Bewertet, wie sehr eine Tabelle nach einer Positionstabelle aussieht
Public Function XR_TableScore(ByVal t As Object) As Long
    Dim h As String, s As Long
    h = LCase$(XR_TableHead(t, 500))
    If InStr(h, "pos") > 0 Then s = s + 1
    If InStr(h, "menge") > 0 Or InStr(h, "anzahl") > 0 Then s = s + 2
    If InStr(h, "preis") > 0 Then s = s + 2
    If InStr(h, "betrag") > 0 Or InStr(h, "summe") > 0 Or InStr(h, "gesamt") > 0 Then s = s + 1
    If InStr(h, "bezeichnung") > 0 Or InStr(h, "beschreibung") > 0 Or InStr(h, "artikel") > 0 Or InStr(h, "leistung") > 0 Then s = s + 1
    XR_TableScore = s
End Function

' Liest eine Tabelle in ein Raster (Zeile, Spalte); Zellzeilen durch vbLf getrennt.
' Funktioniert auch bei verbundenen Zellen und unterschiedlich langen Zeilen.
Public Sub XR_ReadGrid(ByVal t As Object, grid() As String, nRows As Long, nCols As Long)
    Dim c As Object, r As Long, k As Long, maxC As Long, lvl As Long
    nRows = t.Rows.Count
    lvl = t.NestingLevel
    ReDim grid(1 To nRows, 1 To 64)
    For Each c In t.Range.Cells
        If c.NestingLevel = lvl Then
            r = c.RowIndex
            k = c.ColumnIndex
            If r <= nRows And k <= 64 Then
                grid(r, k) = Join(XR_CellLines(XR_RangeText(c.Range)), vbLf)
                If k > maxC Then maxC = k
            End If
        End If
    Next
    If maxC = 0 Then maxC = 1
    ReDim Preserve grid(1 To nRows, 1 To maxC)
    nCols = maxC
End Sub

Private Function HasLetter(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If LCase$(ch) <> UCase$(ch) Then HasLetter = True: Exit Function
    Next
End Function

Private Function LV(vals() As String, ids() As String, ByVal n As Long, ByVal id As String) As String
    Dim k As Long
    For k = 1 To n
        If ids(k) = id Then LV = vals(k): Exit Function
    Next
End Function

' Baut die Rechnungspositionen aus dem Tabellenraster und der Spaltenzuordnung
Public Function XR_BuildLines(grid() As String, ByVal nRows As Long, ByVal nCols As Long, _
                              ids() As String, cols() As Long, specs() As String, defs() As String, _
                              ByVal nMap As Long, ByVal rFrom As Long, ByVal rTo As Long, _
                              lines() As tLine) As Long
    Dim r As Long, k As Long, n As Long, z As tLine, vals() As String, dv() As String
    Dim okQ As Boolean, okP As Boolean, okN As Boolean, okB As Boolean, okA As Boolean, okR As Boolean
    Dim q As Double, pr As Double, nt As Double, bq As Double, al As Double, rt As Double
    Dim sName As String, sDesc As String, sSell As String

    If nRows < 1 Then Exit Function
    If rFrom < 1 Then rFrom = 1
    If rTo > nRows Then rTo = nRows
    ReDim lines(1 To nRows)
    ReDim vals(1 To nMap)
    ReDim dv(1 To nMap)
    For r = rFrom To rTo
        For k = 1 To nMap
            vals(k) = ""
            If cols(k) >= 1 And cols(k) <= nCols Then vals(k) = XR_PickLines(grid(r, cols(k)), specs(k))
            dv(k) = vals(k)
            If Len(dv(k)) = 0 Then dv(k) = defs(k)
        Next
        sName = LV(vals, ids, nMap, "BT-153")
        sDesc = LV(vals, ids, nMap, "BT-154")
        sSell = XR_StripLabel(LV(vals, ids, nMap, "BT-155"))
        q = XR_ParseNum(LV(vals, ids, nMap, "BT-129"), okQ)
        pr = XR_ParseNum(LV(vals, ids, nMap, "BT-146"), okP)
        nt = XR_ParseNum(LV(vals, ids, nMap, "BT-131"), okN)
        ' Position nur, wenn ein Text mit Buchstaben und eine Zahl vorhanden sind
        ' (Summenzeilen wie "384,84 | 9,90 | ..." werden so übersprungen)
        If (HasLetter(sName) Or HasLetter(sDesc) Or HasLetter(sSell)) And (okQ Or okP Or okN) Then
            n = n + 1
            lines(n) = z
            If Not okQ Then q = XR_ParseNum(LV(dv, ids, nMap, "BT-129"), okQ)
            If Not okQ Then q = 1
            bq = XR_ParseNum(LV(dv, ids, nMap, "BT-149"), okB)
            If Not okB Or bq = 0 Then bq = 1
            al = XR_ParseNum(LV(dv, ids, nMap, "BT-136"), okA)
            If Not okA Then al = 0
            al = Abs(al)
            If Not okN Then nt = XR_RoundHU(q * pr / bq - al, 2)
            If Not okP Then
                If q <> 0 Then pr = XR_RoundHU((nt + al) / q * bq, 4) Else pr = 0
            End If
            With lines(n)
                .SrcRow = r
                .id = XR_StripLabel(LV(vals, ids, nMap, "BT-126"))
                If Len(.id) = 0 Then .id = CStr(n)
                .Qty = q
                .Price = pr
                .BaseQty = bq
                .Allow = al
                .AllowReason = LV(dv, ids, nMap, "BT-139")
                .Net = XR_RoundHU(nt, 2)
                .Unit = XR_UnitCode(LV(dv, ids, nMap, "BT-130"), "H87")
                .cat = XR_NormCat(LV(dv, ids, nMap, "BT-151"))
                If Len(.cat) = 0 Then .cat = "S"
                rt = XR_ParseNum(LV(dv, ids, nMap, "BT-152"), okR)
                If Not okR Then rt = 0
                If IsExemptCat(.cat) Or .cat = "Z" Then rt = 0
                .rate = rt
                .Name = sName
                .Desc = sDesc
                If Len(.Name) = 0 Then .Name = .Desc: .Desc = ""
                If Len(.Name) = 0 Then .Name = sSell
                If .Desc = .Name Then .Desc = ""
                .SellerID = sSell
                .BuyerID = XR_StripLabel(LV(vals, ids, nMap, "BT-156"))
                .StdID = XR_StripLabel(LV(vals, ids, nMap, "BT-157"))
                .Origin = LV(dv, ids, nMap, "BT-159")
                .Note = LV(dv, ids, nMap, "BT-127")
                .OrderLine = XR_StripLabel(LV(vals, ids, nMap, "BT-132"))
                .AcctRef = LV(dv, ids, nMap, "BT-133")
                .PStart = LV(dv, ids, nMap, "BT-134")
                .PEnd = LV(dv, ids, nMap, "BT-135")
            End With
        End If
    Next
    XR_BuildLines = n
End Function

'==============================================================================
' Berechnung der Summen (BG-22) und der USt-Aufschlüsselung (BG-23)
'==============================================================================

Private Sub AddGroup(tot As tTotals, ByVal cat As String, ByVal rate As Double, ByVal Amt As Double)
    Dim k As Long
    For k = 1 To tot.nGroups
        If tot.gCat(k) = cat And Abs(tot.gRate(k) - rate) < 0.00001 Then
            tot.gBase(k) = tot.gBase(k) + Amt
            Exit Sub
        End If
    Next
    tot.nGroups = tot.nGroups + 1
    ReDim Preserve tot.gCat(1 To tot.nGroups)
    ReDim Preserve tot.gRate(1 To tot.nGroups)
    ReDim Preserve tot.gBase(1 To tot.nGroups)
    ReDim Preserve tot.gVat(1 To tot.nGroups)
    tot.gCat(tot.nGroups) = cat
    tot.gRate(tot.nGroups) = rate
    tot.gBase(tot.nGroups) = Amt
End Sub

Private Function ACAmount(ByVal idAmt As String, ByVal idBase As String, ByVal idPct As String) As Double
    Dim ok As Boolean, okB As Boolean, okP As Boolean, b As Double, p As Double
    ACAmount = Abs(XR_ParseNum(TX(idAmt), ok))
    If ok Then ACAmount = XR_RoundHU(ACAmount, 2): Exit Function
    b = XR_ParseNum(TX(idBase), okB)
    p = XR_ParseNum(TX(idPct), okP)
    If okB And okP Then ACAmount = XR_RoundHU(Abs(b * p / 100), 2) Else ACAmount = 0
End Function

' USt-Kategorie/-Satz eines Nachlasses/Zuschlags; ohne Angabe wird bei nur
' einer Steuergruppe deren Kategorie und Satz übernommen.
Private Sub ResolveCat(tot As tTotals, ByVal idCat As String, ByVal idRate As String, _
                       cat As String, rate As Double)
    Dim k As Long, ok As Boolean
    cat = XR_NormCat(TX(idCat))
    If Len(cat) = 0 And tot.nGroups = 1 Then cat = tot.gCat(1)
    rate = XR_ParseNum(TX(idRate), ok)
    If Not ok Then
        rate = 0
        For k = 1 To tot.nGroups
            If tot.gCat(k) = cat Then rate = tot.gRate(k): Exit For
        Next
    End If
    If IsExemptCat(cat) Or cat = "Z" Then rate = 0
End Sub

Public Sub XR_CalcTotals(lines() As tLine, ByVal n As Long, tot As tTotals)
    Dim z As tTotals, i As Long, k As Long
    tot = z
    ReDim tot.gCat(1 To 1)
    ReDim tot.gRate(1 To 1)
    ReDim tot.gBase(1 To 1)
    ReDim tot.gVat(1 To 1)
    For i = 1 To n
        tot.LineSum = tot.LineSum + XR_RoundHU(lines(i).Net, 2)
        AddGroup tot, lines(i).cat, lines(i).rate, XR_RoundHU(lines(i).Net, 2)
    Next
    tot.AllowAmt = ACAmount("BT-92", "BT-93", "BT-94")
    If tot.AllowAmt <> 0 Then
        ResolveCat tot, "BT-95", "BT-96", tot.AllowCat, tot.AllowRate
        If Len(tot.AllowCat) > 0 Then AddGroup tot, tot.AllowCat, tot.AllowRate, -tot.AllowAmt
    End If
    tot.ChargeAmt = ACAmount("BT-99", "BT-100", "BT-101")
    If tot.ChargeAmt <> 0 Then
        ResolveCat tot, "BT-102", "BT-103", tot.ChargeCat, tot.ChargeRate
        If Len(tot.ChargeCat) > 0 Then AddGroup tot, tot.ChargeCat, tot.ChargeRate, tot.ChargeAmt
    End If
    For k = 1 To tot.nGroups
        tot.gBase(k) = XR_RoundHU(tot.gBase(k), 2)
        tot.gVat(k) = XR_RoundHU(tot.gBase(k) * tot.gRate(k) / 100, 2)
        tot.VatTotal = tot.VatTotal + tot.gVat(k)
    Next
    tot.LineSum = XR_RoundHU(tot.LineSum, 2)
    tot.AllowSum = tot.AllowAmt
    tot.ChargeSum = tot.ChargeAmt
    tot.NetTotal = XR_RoundHU(tot.LineSum - tot.AllowSum + tot.ChargeSum, 2)
    tot.VatTotal = XR_RoundHU(tot.VatTotal, 2)
    tot.Gross = XR_RoundHU(tot.NetTotal + tot.VatTotal, 2)
    tot.Prepaid = XR_RoundHU(NumV("BT-113"), 2)
    tot.Rounding = XR_RoundHU(NumV("BT-114"), 2)
    tot.Due = XR_RoundHU(tot.Gross - tot.Prepaid + tot.Rounding, 2)
End Sub

'==============================================================================
' Prüfung der Geschäftsregeln (EN 16931 / XRechnung BR-DE)
'==============================================================================

Public Sub XR_Check(lines() As tLine, ByVal n As Long, tot As tTotals, _
                    errs As Collection, warns As Collection)
    Dim defs() As tBTDef, dn As Long, i As Long, k As Long, v As String, code As String
    Dim cat As String, calcNet As Double, hasO As Boolean

    dn = XR_GetHeaderDefs(defs)
    For i = 1 To dn
        If defs(i).req = "P" And defs(i).kind <> "K" And defs(i).kind <> "G" Then
            If Len(TX(defs(i).id)) = 0 Then errs.Add defs(i).id & " (" & defs(i).Name & ") fehlt."
        End If
    Next
    If n = 0 Then errs.Add "BG-25: Die Rechnung enthält keine Positionen."

    code = XR_NormInvType(TX("BT-3"))
    Select Case code
        Case "326", "380", "381", "384", "389", "875", "876", "877"
        Case Else
            errs.Add "BT-3: Rechnungstyp """ & TX("BT-3") & """ ist nicht zulässig (326, 380, 381, 384, 389, 875, 876, 877)."
    End Select
    If code = "384" And Len(TX("BT-25")) = 0 Then errs.Add "BT-25: Bei einer Rechnungskorrektur (384) muss die vorausgegangene Rechnung angegeben werden (BR-DE-26)."
    If Not XR_NormCurrency(TX("BT-5")) Like "[A-Z][A-Z][A-Z]" Then errs.Add "BT-5: ungültiger Währungscode """ & TX("BT-5") & """."

    If Len(TX("BT-31")) = 0 And Len(TX("BT-32")) = 0 And Len(TX("BT-63")) = 0 Then
        errs.Add "BT-31/BT-32: USt-IdNr. oder Steuernummer des Verkäufers muss angegeben werden (BR-DE-16)."
    End If
    If Len(TX("BT-29")) = 0 And Len(TX("BT-30")) = 0 And Len(TX("BT-31")) = 0 Then
        errs.Add "BT-29/BT-30/BT-31: Kennung, Registernummer oder USt-IdNr. des Verkäufers erforderlich (BR-CO-26)."
    End If
    v = TX("BT-43")
    If Len(v) > 0 And Not (v Like "*?@?*.?*") Then errs.Add "BT-43: ungültige E-Mail-Adresse """ & v & """ (BR-DE-28)."
    v = TX("BT-42")
    If Len(v) > 0 And CountDigits(v) < 3 Then errs.Add "BT-42: Die Telefonnummer muss mindestens 3 Ziffern enthalten (BR-DE-27)."

    code = XR_NormPayMeans(TX("BT-81"))
    Select Case code
        Case "30", "58"
            If Len(TX("BT-84")) = 0 Then
                errs.Add "BT-84: Bei Überweisung (" & code & ") muss die IBAN angegeben werden (BR-DE-23)."
            ElseIf code = "58" And Not IbanOk(XR_NormID(TX("BT-84"))) Then
                warns.Add "BT-84: Die Prüfziffer der IBAN ist ungültig (BR-DE-19)."
            End If
        Case "59"
            If Len(TX("BT-89")) = 0 Then errs.Add "BT-89: Bei Lastschrift muss die Mandatsreferenz angegeben werden."
            If Len(TX("BT-90")) = 0 Then errs.Add "BT-90: Bei Lastschrift muss die Gläubiger-ID angegeben werden."
            If Len(TX("BT-91")) = 0 Then errs.Add "BT-91: Bei Lastschrift muss das zu belastende Konto angegeben werden."
        Case "48", "54", "55"
            If Len(TX("BT-87")) = 0 Then errs.Add "BT-87: Bei Kartenzahlung muss die Kartennummer angegeben werden."
        Case ""
        Case Else
            If code Like "#*" Then
                warns.Add "BT-81: Zahlungsart " & code & " - XRechnung erwartet Überweisung (30/58), Lastschrift (59) oder Karte (48/54/55) (BR-DE-13)."
            Else
                errs.Add "BT-81: """ & TX("BT-81") & """ ist kein Code für die Zahlungsart (z.B. 58 = SEPA-Überweisung)."
            End If
    End Select
    If tot.Due > 0 And Len(TX("BT-9")) = 0 And Len(TX("BT-20")) = 0 Then
        errs.Add "BT-9/BT-20: Fälligkeitsdatum oder Zahlungsbedingungen müssen angegeben werden (BR-CO-25)."
    End If

    If AnyV("BT-75", "BT-76", "BT-165", "BT-77", "BT-78", "BT-79", "BT-80") Then
        If Len(TX("BT-77")) = 0 Then errs.Add "BT-77: Stadt der Lieferanschrift fehlt (BR-DE-10)."
        If Len(TX("BT-78")) = 0 Then errs.Add "BT-78: Postleitzahl der Lieferanschrift fehlt (BR-DE-11)."
        If Len(TX("BT-80")) = 0 Then errs.Add "BT-80: Ländercode der Lieferanschrift fehlt (BR-57)."
    End If
    If AnyV("BT-62", "BT-63") Then
        If Len(TX("BT-62")) = 0 Or Len(TX("BT-63")) = 0 Or Len(TX("BT-69")) = 0 Then
            errs.Add "BG-11: Für den Steuervertreter sind Name (BT-62), USt-IdNr. (BT-63) und Ländercode (BT-69) erforderlich."
        End If
    End If
    If Len(TX("BT-6")) > 0 And Len(TX("BT-111")) = 0 Then errs.Add "BT-111: Bei abweichender Buchungswährung (BT-6) muss der USt-Betrag in dieser Währung angegeben werden."

    If tot.AllowAmt <> 0 Then
        If Len(tot.AllowCat) = 0 Then errs.Add "BT-95: USt-Kategorie des Nachlasses fehlt."
        If Len(TX("BT-97")) = 0 And Len(TX("BT-98")) = 0 Then warns.Add "BT-97: Kein Grund für den Nachlass angegeben - es wird ""Nachlass"" verwendet."
    End If
    If tot.ChargeAmt <> 0 Then
        If Len(tot.ChargeCat) = 0 Then errs.Add "BT-102: USt-Kategorie des Zuschlags fehlt."
        If Len(TX("BT-104")) = 0 And Len(TX("BT-105")) = 0 Then warns.Add "BT-104: Kein Grund für den Zuschlag angegeben - es wird ""Zuschlag"" verwendet."
    End If

    For k = 1 To tot.nGroups
        cat = tot.gCat(k)
        Select Case cat
            Case "S"
                If tot.gRate(k) <= 0 Then errs.Add "USt-Kategorie S (Normalsatz) erfordert einen Steuersatz größer 0 %."
            Case "Z", "L", "M"
            Case "E", "AE", "K", "G", "O"
                If cat = "O" Then hasO = True
                If Len(TX("BT-120")) = 0 And Len(TX("BT-121")) = 0 Then
                    errs.Add "BT-120/BT-121: Für die USt-Kategorie " & cat & " muss ein Befreiungsgrund angegeben werden."
                End If
                If (cat = "AE" Or cat = "K") And Len(TX("BT-48")) = 0 Then
                    errs.Add "BT-48: Für die USt-Kategorie " & cat & " ist die USt-IdNr. des Käufers erforderlich."
                End If
            Case Else
                errs.Add "Unbekannte USt-Kategorie """ & cat & """ (zulässig: S, Z, E, AE, K, G, O, L, M)."
        End Select
    Next
    If hasO And tot.nGroups > 1 Then errs.Add "USt-Kategorie O (nicht steuerbar) darf nicht mit anderen Kategorien kombiniert werden (BR-O-11)."

    For i = 1 To n
        With lines(i)
            If Len(.Name) = 0 Then errs.Add "Position " & .id & ": Artikelname (BT-153) fehlt."
            If Len(.Unit) = 0 Then errs.Add "Position " & .id & ": Einheit (BT-130) fehlt."
            If .Price < 0 Then errs.Add "Position " & .id & ": Der Nettopreis darf nicht negativ sein (BR-27)."
            calcNet = XR_RoundHU(.Qty * .Price / .BaseQty - .Allow, 2)
            If Abs(calcNet - .Net) > 0.011 Then
                warns.Add "Position " & .id & ": Menge x Preis ergibt " & XR_FmtAmt(calcNet) & ", Nettobetrag ist " & XR_FmtAmt(.Net) & "."
            End If
            If Len(.PStart) > 0 Or Len(.PEnd) > 0 Then
                If Not IsIsoOk(.PStart) Or Not IsIsoOk(.PEnd) Then errs.Add "Position " & .id & ": ungültiges Datum im Abrechnungszeitraum."
            End If
        End With
    Next
End Sub

Private Function IsIsoOk(ByVal s As String) As Boolean
    Dim ok As Boolean
    If Len(s) = 0 Then IsIsoOk = True: Exit Function
    XR_IsoDate s, ok
    IsIsoOk = ok
End Function

'==============================================================================
' XML-Erzeugung (UBL 2.1 Invoice) mit MSXML2.DOMDocument
'==============================================================================

Public Function XR_NewDom() As Object
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("MSXML2.DOMDocument.6.0")
    If d Is Nothing Then Set d = CreateObject("MSXML2.DOMDocument")
    On Error GoTo 0
    If d Is Nothing Then Err.Raise vbObjectError + 1, "XRechnung", "MSXML2.DOMDocument ist nicht verfügbar."
    d.async = False
    Set XR_NewDom = d
End Function

Private Function NewEl(ByVal qn As String) As Object
    Dim ns As String
    Select Case Left$(qn, 4)
        Case "cac:": ns = NS_CAC
        Case "cbc:": ns = NS_CBC
        Case Else: ns = NS_INV
    End Select
    Set NewEl = mDom.createNode(1, qn, ns)
End Function

Private Function Node(ByVal parent As Object, ByVal qn As String) As Object
    Set Node = parent.appendChild(NewEl(qn))
End Function

' Fügt ein Element mit Text nur hinzu, wenn der Text nicht leer ist
Private Function Leaf(ByVal parent As Object, ByVal qn As String, ByVal txt As String, _
                      Optional ByVal attr As String = "", Optional ByVal attrVal As String = "") As Object
    Dim e As Object
    If Len(txt) = 0 Then Exit Function
    Set e = NewEl(qn)
    e.Text = txt
    If Len(attr) > 0 And Len(attrVal) > 0 Then e.setAttribute attr, attrVal
    Set Leaf = parent.appendChild(e)
End Function

Private Sub Amt(ByVal parent As Object, ByVal qn As String, ByVal v As Double)
    Leaf parent, qn, XR_FmtAmt(v), "currencyID", mCur
End Sub

Private Function DX(ByVal id As String) As String
    Dim v As String, ok As Boolean
    v = TX(id)
    If Len(v) = 0 Then Exit Function
    DX = XR_IsoDate(v, ok)
    If Not ok Then mErrs.Add id & ": """ & v & """ ist kein gültiges Datum."
End Function

Private Function AX(ByVal id As String) As String
    Dim v As String, ok As Boolean, x As Double
    v = TX(id)
    If Len(v) = 0 Then Exit Function
    x = XR_ParseNum(v, ok)
    If ok Then AX = XR_FmtAmt(x) Else mErrs.Add id & ": """ & v & """ ist kein gültiger Betrag."
End Function

Private Function PX(ByVal id As String) As String
    Dim v As String, ok As Boolean, x As Double
    v = TX(id)
    If Len(v) = 0 Then Exit Function
    x = XR_ParseNum(v, ok)
    If ok Then PX = XR_FmtDec(x, 2) Else mErrs.Add id & ": """ & v & """ ist keine gültige Zahl."
End Function

Private Function CX(ByVal id As String) As String
    Dim v As String
    v = TX(id)
    If Len(v) = 0 Then Exit Function
    CX = XR_NormCountry(v)
    If Not CX Like "[A-Z][A-Z]" Then mErrs.Add id & ": """ & v & """ ist kein Ländercode nach ISO 3166-1 (z.B. DE)."
End Function

Private Sub Endpoint(ByVal party As Object, ByVal id As String)
    Dim v As String, sch As String, p As Long, pre As String
    v = TX(id)
    If Len(v) = 0 Then Exit Sub
    ' explizites Schema "Schema:Kennung", z.B. EM:rechnung@firma.de oder 0204:991-...
    p = InStr(v, ":")
    If p >= 3 And p <= 5 Then
        pre = UCase$(Left$(v, p - 1))
        If Not pre Like "*[!A-Z0-9]*" Then
            sch = pre
            v = Trim$(Mid$(v, p + 1))
            If sch = "EM" Then v = Replace(v, " ", "")
        End If
    End If
    If Len(sch) > 0 Then
        If Len(v) = 0 Then mErrs.Add id & ": Nach dem Schema """ & sch & ":"" fehlt die Kennung.": Exit Sub
    ElseIf InStr(v, "@") > 0 Then
        sch = "EM"
        v = Replace(Replace(v, "mailto:", "", 1, -1, vbTextCompare), " ", "")
    ElseIf v Like "#*-*" Then
        sch = "0204"                    ' Leitweg-ID
    ElseIf XR_NormID(v) Like "DE#########" Then
        sch = "9930"                    ' deutsche USt-IdNr.
        v = XR_NormID(v)
    Else
        mErrs.Add id & ": Art der elektronischen Adresse """ & v & """ unbekannt - E-Mail-Adresse, Leitweg-ID oder ""Schema:Kennung"" (z.B. 0088:4000001000005) angeben."
        Exit Sub
    End If
    Leaf party, "cbc:EndpointID", v, "schemeID", sch
End Sub

Private Sub Address(ByVal parent As Object, ByVal qn As String, ByVal s1 As String, ByVal s2 As String, _
                    ByVal s3 As String, ByVal city As String, ByVal pc As String, _
                    ByVal region As String, ByVal country As String)
    Dim x As Object, c As String
    Set x = Node(parent, qn)
    Leaf x, "cbc:StreetName", TX(s1)
    Leaf x, "cbc:AdditionalStreetName", TX(s2)
    Leaf x, "cbc:CityName", TX(city)
    Leaf x, "cbc:PostalZone", XR_NormPostcode(TX(pc))
    Leaf x, "cbc:CountrySubentity", TX(region)
    If AnyV(s3) Then Leaf Node(x, "cac:AddressLine"), "cbc:Line", TX(s3)
    c = CX(country)
    If Len(c) > 0 Then Leaf Node(x, "cac:Country"), "cbc:IdentificationCode", c
End Sub

Private Sub TaxScheme(ByVal party As Object, ByVal companyID As String, ByVal scheme As String)
    Dim x As Object
    If Len(companyID) = 0 Then Exit Sub
    Set x = Node(party, "cac:PartyTaxScheme")
    Leaf x, "cbc:CompanyID", companyID
    Leaf Node(x, "cac:TaxScheme"), "cbc:ID", scheme
End Sub

Private Sub Contact(ByVal party As Object, ByVal idName As String, ByVal idTel As String, ByVal idMail As String)
    Dim x As Object
    If Not AnyV(idName, idTel, idMail) Then Exit Sub
    Set x = Node(party, "cac:Contact")
    Leaf x, "cbc:Name", TX(idName)
    Leaf x, "cbc:Telephone", TX(idTel)
    Leaf x, "cbc:ElectronicMail", TX(idMail)
End Sub

Private Sub RefDoc(ByVal root As Object, ByVal qn As String, ByVal id As String)
    If AnyV(id) Then Leaf Node(root, qn), "cbc:ID", TX(id)
End Sub

Private Sub TaxCat(ByVal parent As Object, ByVal qn As String, ByVal cat As String, _
                   ByVal rate As Double, ByVal withExemption As Boolean)
    Dim x As Object
    Set x = Node(parent, qn)
    Leaf x, "cbc:ID", cat
    If cat <> "O" Then Leaf x, "cbc:Percent", XR_FmtDec(rate, 2)
    If withExemption And IsExemptCat(cat) Then
        Leaf x, "cbc:TaxExemptionReasonCode", TX("BT-121")
        Leaf x, "cbc:TaxExemptionReason", TX("BT-120")
    End If
    Leaf Node(x, "cac:TaxScheme"), "cbc:ID", "VAT"
End Sub

Private Sub DocAC(ByVal root As Object, ByVal isCharge As Boolean, ByVal amount As Double, _
                  ByVal idBase As String, ByVal idPct As String, ByVal cat As String, _
                  ByVal rate As Double, ByVal idReason As String, ByVal idCode As String)
    Dim x As Object, reason As String
    Set x = Node(root, "cac:AllowanceCharge")
    Leaf x, "cbc:ChargeIndicator", IIf(isCharge, "true", "false")
    Leaf x, "cbc:AllowanceChargeReasonCode", TX(idCode)
    reason = TX(idReason)
    If Len(reason) = 0 And Len(TX(idCode)) = 0 Then reason = IIf(isCharge, "Zuschlag", "Nachlass")
    Leaf x, "cbc:AllowanceChargeReason", reason
    Leaf x, "cbc:MultiplierFactorNumeric", PX(idPct)
    Amt x, "cbc:Amount", amount
    Leaf x, "cbc:BaseAmount", AX(idBase), "currencyID", mCur
    TaxCat x, "cac:TaxCategory", cat, rate, False
End Sub

Private Function LineDate(ByVal s As String) As String
    Dim ok As Boolean
    If Len(s) > 0 Then LineDate = XR_IsoDate(s, ok)
End Function

Private Sub AddLine(ByVal root As Object, ln As tLine)
    Dim p As Object, x As Object, it As Object
    Set p = Node(root, "cac:InvoiceLine")
    Leaf p, "cbc:ID", ln.id
    Leaf p, "cbc:Note", ln.Note
    Leaf p, "cbc:InvoicedQuantity", XR_FmtQty(ln.Qty), "unitCode", ln.Unit
    Amt p, "cbc:LineExtensionAmount", ln.Net
    Leaf p, "cbc:AccountingCost", ln.AcctRef
    If Len(ln.PStart) > 0 Or Len(ln.PEnd) > 0 Then
        Set x = Node(p, "cac:InvoicePeriod")
        Leaf x, "cbc:StartDate", LineDate(ln.PStart)
        Leaf x, "cbc:EndDate", LineDate(ln.PEnd)
    End If
    If Len(ln.OrderLine) > 0 Then Leaf Node(p, "cac:OrderLineReference"), "cbc:LineID", ln.OrderLine
    If ln.Allow <> 0 Then
        Set x = Node(p, "cac:AllowanceCharge")
        Leaf x, "cbc:ChargeIndicator", "false"
        Leaf x, "cbc:AllowanceChargeReason", IIf(Len(ln.AllowReason) > 0, ln.AllowReason, "Rabatt")
        Amt x, "cbc:Amount", ln.Allow
    End If
    Set it = Node(p, "cac:Item")
    Leaf it, "cbc:Description", ln.Desc
    Leaf it, "cbc:Name", ln.Name
    If Len(ln.BuyerID) > 0 Then Leaf Node(it, "cac:BuyersItemIdentification"), "cbc:ID", ln.BuyerID
    If Len(ln.SellerID) > 0 Then Leaf Node(it, "cac:SellersItemIdentification"), "cbc:ID", ln.SellerID
    If Len(ln.StdID) > 0 Then Leaf Node(it, "cac:StandardItemIdentification"), "cbc:ID", ln.StdID, "schemeID", "0160"
    If Len(ln.Origin) > 0 Then Leaf Node(it, "cac:OriginCountry"), "cbc:IdentificationCode", XR_NormCountry(ln.Origin)
    TaxCat it, "cac:ClassifiedTaxCategory", ln.cat, ln.rate, False
    Set x = Node(p, "cac:Price")
    Leaf x, "cbc:PriceAmount", XR_FmtPrice(ln.Price), "currencyID", mCur
    If ln.BaseQty <> 1 Then Leaf x, "cbc:BaseQuantity", XR_FmtQty(ln.BaseQty), "unitCode", ln.Unit
End Sub

' Erzeugt das XML-Dokument; Formatfehler werden in errs gesammelt
Public Function XR_BuildXml(lines() As tLine, ByVal n As Long, tot As tTotals, errs As Collection) As Object
    Dim root As Object, p As Object, x As Object, e As Object, i As Long, k As Long

    Set mErrs = errs
    Set mDom = XR_NewDom()
    mCur = XR_NormCurrency(TX("BT-5"))
    If Len(mCur) = 0 Then mCur = "EUR"

    Set root = mDom.createNode(1, "Invoice", NS_INV)
    mDom.appendChild root
    root.setAttribute "xmlns:cac", NS_CAC
    root.setAttribute "xmlns:cbc", NS_CBC

    Leaf root, "cbc:CustomizationID", IIf(Len(TX("BT-24")) > 0, TX("BT-24"), XR_SPEC_ID)
    Leaf root, "cbc:ProfileID", TX("BT-23")
    Leaf root, "cbc:ID", TX("BT-1")
    Leaf root, "cbc:IssueDate", DX("BT-2")
    Leaf root, "cbc:DueDate", DX("BT-9")
    Leaf root, "cbc:InvoiceTypeCode", XR_NormInvType(TX("BT-3"))
    If AnyV("BT-22") Then
        If AnyV("BT-21") Then
            Leaf root, "cbc:Note", "#" & UCase$(TX("BT-21")) & "#" & TX("BT-22")
        Else
            Leaf root, "cbc:Note", TX("BT-22")
        End If
    End If
    Leaf root, "cbc:TaxPointDate", DX("BT-7")
    Leaf root, "cbc:DocumentCurrencyCode", mCur
    Leaf root, "cbc:TaxCurrencyCode", XR_NormCurrency(TX("BT-6"))
    Leaf root, "cbc:AccountingCost", TX("BT-19")
    Leaf root, "cbc:BuyerReference", TX("BT-10")
    If AnyV("BT-73", "BT-74", "BT-8") Then
        Set p = Node(root, "cac:InvoicePeriod")
        Leaf p, "cbc:StartDate", DX("BT-73")
        Leaf p, "cbc:EndDate", DX("BT-74")
        Leaf p, "cbc:DescriptionCode", TX("BT-8")
    End If
    If AnyV("BT-13", "BT-14") Then
        Set p = Node(root, "cac:OrderReference")
        Leaf p, "cbc:ID", IIf(Len(TX("BT-13")) > 0, TX("BT-13"), "NA")
        Leaf p, "cbc:SalesOrderID", TX("BT-14")
    End If
    If AnyV("BT-25") Then
        Set p = Node(Node(root, "cac:BillingReference"), "cac:InvoiceDocumentReference")
        Leaf p, "cbc:ID", TX("BT-25")
        Leaf p, "cbc:IssueDate", DX("BT-26")
    End If
    RefDoc root, "cac:DespatchDocumentReference", "BT-16"
    RefDoc root, "cac:ReceiptDocumentReference", "BT-15"
    RefDoc root, "cac:OriginatorDocumentReference", "BT-17"
    RefDoc root, "cac:ContractDocumentReference", "BT-12"
    If AnyV("BT-18") Then
        Set p = Node(root, "cac:AdditionalDocumentReference")
        Leaf p, "cbc:ID", TX("BT-18")
        Leaf p, "cbc:DocumentTypeCode", "130"
    End If
    RefDoc root, "cac:ProjectReference", "BT-11"

    ' BG-4 Verkäufer
    Set p = Node(Node(root, "cac:AccountingSupplierParty"), "cac:Party")
    Endpoint p, "BT-34"
    If AnyV("BT-29") Then Leaf Node(p, "cac:PartyIdentification"), "cbc:ID", TX("BT-29")
    If AnyV("BT-90") Then Leaf Node(p, "cac:PartyIdentification"), "cbc:ID", XR_NormID(TX("BT-90")), "schemeID", "SEPA"
    If AnyV("BT-28") Then Leaf Node(p, "cac:PartyName"), "cbc:Name", TX("BT-28")
    Address p, "cac:PostalAddress", "BT-35", "BT-36", "BT-162", "BT-37", "BT-38", "BT-39", "BT-40"
    TaxScheme p, XR_NormID(TX("BT-31")), "VAT"
    TaxScheme p, TX("BT-32"), "FC"
    Set x = Node(p, "cac:PartyLegalEntity")
    Leaf x, "cbc:RegistrationName", TX("BT-27")
    Leaf x, "cbc:CompanyID", TX("BT-30")
    Leaf x, "cbc:CompanyLegalForm", TX("BT-33")
    Contact p, "BT-41", "BT-42", "BT-43"

    ' BG-7 Käufer
    Set p = Node(Node(root, "cac:AccountingCustomerParty"), "cac:Party")
    Endpoint p, "BT-49"
    If AnyV("BT-46") Then Leaf Node(p, "cac:PartyIdentification"), "cbc:ID", TX("BT-46")
    If AnyV("BT-45") Then Leaf Node(p, "cac:PartyName"), "cbc:Name", TX("BT-45")
    Address p, "cac:PostalAddress", "BT-50", "BT-51", "BT-163", "BT-52", "BT-53", "BT-54", "BT-55"
    TaxScheme p, XR_NormID(TX("BT-48")), "VAT"
    Set x = Node(p, "cac:PartyLegalEntity")
    Leaf x, "cbc:RegistrationName", TX("BT-44")
    Leaf x, "cbc:CompanyID", TX("BT-47")
    Contact p, "BT-56", "BT-57", "BT-58"

    ' BG-10 Zahlungsempfänger
    If AnyV("BT-59", "BT-60", "BT-61") Then
        Set p = Node(root, "cac:PayeeParty")
        If AnyV("BT-60") Then Leaf Node(p, "cac:PartyIdentification"), "cbc:ID", TX("BT-60")
        If AnyV("BT-59") Then Leaf Node(p, "cac:PartyName"), "cbc:Name", TX("BT-59")
        If AnyV("BT-61") Then Leaf Node(p, "cac:PartyLegalEntity"), "cbc:CompanyID", TX("BT-61")
    End If

    ' BG-11 Steuervertreter
    If AnyV("BT-62", "BT-63") Then
        Set p = Node(root, "cac:TaxRepresentativeParty")
        Leaf Node(p, "cac:PartyName"), "cbc:Name", TX("BT-62")
        Address p, "cac:PostalAddress", "BT-64", "BT-65", "BT-164", "BT-66", "BT-67", "BT-68", "BT-69"
        TaxScheme p, XR_NormID(TX("BT-63")), "VAT"
    End If

    ' BG-13 Lieferung
    If AnyV("BT-70", "BT-71", "BT-72", "BT-75", "BT-76", "BT-165", "BT-77", "BT-78", "BT-79", "BT-80") Then
        Set p = Node(root, "cac:Delivery")
        Leaf p, "cbc:ActualDeliveryDate", DX("BT-72")
        If AnyV("BT-71", "BT-75", "BT-76", "BT-165", "BT-77", "BT-78", "BT-79", "BT-80") Then
            Set x = Node(p, "cac:DeliveryLocation")
            Leaf x, "cbc:ID", TX("BT-71")
            If AnyV("BT-75", "BT-76", "BT-165", "BT-77", "BT-78", "BT-79", "BT-80") Then
                Address x, "cac:Address", "BT-75", "BT-76", "BT-165", "BT-77", "BT-78", "BT-79", "BT-80"
            End If
        End If
        If AnyV("BT-70") Then Leaf Node(Node(p, "cac:DeliveryParty"), "cac:PartyName"), "cbc:Name", TX("BT-70")
    End If

    ' BG-16 Zahlungsanweisungen
    Set p = Node(root, "cac:PaymentMeans")
    Set e = Leaf(p, "cbc:PaymentMeansCode", XR_NormPayMeans(TX("BT-81")))
    If Not e Is Nothing And AnyV("BT-82") Then e.setAttribute "name", TX("BT-82")
    Leaf p, "cbc:PaymentID", TX("BT-83")
    If AnyV("BT-87") Then
        Set x = Node(p, "cac:CardAccount")
        Leaf x, "cbc:PrimaryAccountNumberID", TX("BT-87")
        Leaf x, "cbc:NetworkID", "NA"
        Leaf x, "cbc:HolderName", TX("BT-88")
    End If
    If AnyV("BT-84") Then
        Set x = Node(p, "cac:PayeeFinancialAccount")
        Leaf x, "cbc:ID", XR_NormID(TX("BT-84"))
        Leaf x, "cbc:Name", TX("BT-85")
        If AnyV("BT-86") Then Leaf Node(x, "cac:FinancialInstitutionBranch"), "cbc:ID", XR_NormID(TX("BT-86"))
    End If
    If AnyV("BT-89", "BT-91") Then
        Set x = Node(p, "cac:PaymentMandate")
        Leaf x, "cbc:ID", TX("BT-89")
        If AnyV("BT-91") Then Leaf Node(x, "cac:PayerFinancialAccount"), "cbc:ID", XR_NormID(TX("BT-91"))
    End If
    If AnyV("BT-20") Then Leaf Node(root, "cac:PaymentTerms"), "cbc:Note", TX("BT-20")

    ' BG-20 / BG-21 Nachlässe und Zuschläge
    If tot.AllowAmt <> 0 Then DocAC root, False, tot.AllowAmt, "BT-93", "BT-94", tot.AllowCat, tot.AllowRate, "BT-97", "BT-98"
    If tot.ChargeAmt <> 0 Then DocAC root, True, tot.ChargeAmt, "BT-100", "BT-101", tot.ChargeCat, tot.ChargeRate, "BT-104", "BT-105"

    ' BG-23 Umsatzsteuer
    Set p = Node(root, "cac:TaxTotal")
    Amt p, "cbc:TaxAmount", tot.VatTotal
    For k = 1 To tot.nGroups
        Set x = Node(p, "cac:TaxSubtotal")
        Amt x, "cbc:TaxableAmount", tot.gBase(k)
        Amt x, "cbc:TaxAmount", tot.gVat(k)
        TaxCat x, "cac:TaxCategory", tot.gCat(k), tot.gRate(k), True
    Next
    If AnyV("BT-111") Then
        Set p = Node(root, "cac:TaxTotal")
        Leaf p, "cbc:TaxAmount", AX("BT-111"), "currencyID", XR_NormCurrency(TX("BT-6"))
    End If

    ' BG-22 Gesamtbeträge
    Set p = Node(root, "cac:LegalMonetaryTotal")
    Amt p, "cbc:LineExtensionAmount", tot.LineSum
    Amt p, "cbc:TaxExclusiveAmount", tot.NetTotal
    Amt p, "cbc:TaxInclusiveAmount", tot.Gross
    If tot.AllowAmt <> 0 Then Amt p, "cbc:AllowanceTotalAmount", tot.AllowSum
    If tot.ChargeAmt <> 0 Then Amt p, "cbc:ChargeTotalAmount", tot.ChargeSum
    If tot.Prepaid <> 0 Then Amt p, "cbc:PrepaidAmount", tot.Prepaid
    If tot.Rounding <> 0 Then Amt p, "cbc:PayableRoundingAmount", tot.Rounding
    Amt p, "cbc:PayableAmount", tot.Due

    ' BG-25 Positionen
    For i = 1 To n
        AddLine root, lines(i)
    Next
    Set XR_BuildXml = mDom
End Function

' Speichert eingerückt als UTF-8; fällt bei Problemen auf die einfache Ausgabe zurück
Public Sub XR_SaveXml(ByVal dom As Object, ByVal path As String)
    Dim wr As Object, rd As Object, outDoc As Object
    On Error GoTo Plain
    Set wr = CreateObject("MSXML2.MXXMLWriter.6.0")
    Set rd = CreateObject("MSXML2.SAXXMLReader.6.0")
    wr.omitXMLDeclaration = True
    wr.Indent = True
    Set rd.contentHandler = wr
    rd.parse dom
    Set outDoc = XR_NewDom()
    outDoc.preserveWhiteSpace = True
    If Not outDoc.LoadXML(wr.output) Then GoTo Plain
    outDoc.InsertBefore outDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8"""), outDoc.ChildNodes(0)
    outDoc.Save path
    Exit Sub
Plain:
    On Error GoTo 0
    If dom.ChildNodes(0).NodeType <> 7 Then
        dom.InsertBefore dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8"""), dom.ChildNodes(0)
    End If
    dom.Save path
End Sub

' Größe des Bildschirm-Arbeitsbereichs in Punkt (für die Dialoggröße)
Public Function XR_ScreenPoints(wPt As Single, hPt As Single) As Boolean
    Dim r As tRect, dpi As Long, hdc As LongPtr
    On Error GoTo Fail
    If SystemParametersInfo(48, 0, r, 0) = 0 Then Exit Function     ' SPI_GETWORKAREA
    hdc = GetDC(0)
    dpi = GetDeviceCaps(hdc, 90)                                     ' LOGPIXELSY
    ReleaseDC 0, hdc
    If dpi <= 0 Then dpi = 96
    wPt = (r.Right - r.Left) * 72 / dpi
    hPt = (r.Bottom - r.Top) * 72 / dpi
    XR_ScreenPoints = (wPt > 0 And hPt > 0)
Fail:
End Function

' Zielpfad: gleicher Name wie das Dokument, Endung .xml
Public Function XR_TargetPath(ByVal doc As Object, Optional ByVal interactive As Boolean = True) As String
    Dim base As String, p As Long, folder As String
    base = doc.Name
    p = InStrRev(base, ".")
    If p > 1 Then base = Left$(base, p - 1)
    folder = doc.path
    If LCase$(Left$(folder, 4)) = "http" Then folder = ""
    If Len(folder) = 0 Then
        If Not interactive Then
            folder = Environ$("USERPROFILE") & "\Documents"
        Else
            With Application.FileDialog(4)
                .Title = "Ordner für die XRechnung """ & base & ".xml"" wählen"
                If .Show <> -1 Then Exit Function
                folder = .SelectedItems(1)
            End With
        End If
    End If
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    XR_TargetPath = folder & base & ".xml"
End Function

