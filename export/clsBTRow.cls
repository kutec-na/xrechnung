VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBTRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsBTRow - eine Zeile der Tabelle "Rechnungsdaten (BT-Felder)"
' BT-Nr. | Bezeichnung | Pflicht | Textmarke (Auswahl) | Inhalt
'==============================================================================
Option Explicit

Private Const X_ID As Single = 2
Private Const X_NAME As Single = 50
Private Const X_REQ As Single = 330
Private Const X_BM As Single = 360
Private Const X_VAL As Single = 536
Private Const CTL_H As Single = 15

Public Owner As Object
Public Index As Long
Public id As String
Public BTName As String
Public req As String
Public kind As String
Public BmValue As String
Public CalcValue As String
Public FallbackValue As String              ' Ersatzwert, wenn das Feld leer ist
Public FallbackFrom As String               ' Herkunft des Ersatzwerts (für die Anzeige)

Public lblID As MSForms.Label
Public lblName As MSForms.Label
Public lblReq As MSForms.Label
Public WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1

Private mBusy As Boolean

Public Sub Create(ByVal frm As Object, ByVal container As Object, ByVal idx As Long, _
                  ByVal sID As String, ByVal sName As String, ByVal sReq As String, _
                  ByVal sKind As String, ByVal items As Variant, ByVal valWidth As Single)
    Set Owner = frm
    Index = idx
    id = sID
    BTName = sName
    req = sReq
    kind = sKind

    Set lblID = container.Controls.Add("Forms.Label.1")
    lblID.Caption = sID
    lblID.Left = X_ID: lblID.Width = X_NAME - X_ID - 2: lblID.Height = CTL_H - 2

    Set lblName = container.Controls.Add("Forms.Label.1")
    lblName.Caption = sName
    lblName.WordWrap = False
    lblName.ControlTipText = sID & " - " & sName
    lblName.Left = X_NAME: lblName.Width = X_REQ - X_NAME - 4: lblName.Height = CTL_H - 2

    Set lblReq = container.Controls.Add("Forms.Label.1")
    lblReq.Caption = sReq
    lblReq.TextAlign = 2                    ' fmTextAlignCenter
    lblReq.Left = X_REQ: lblReq.Width = X_BM - X_REQ - 4: lblReq.Height = CTL_H - 2
    Select Case sReq
        Case "P": lblReq.Font.Bold = True: lblReq.ControlTipText = "Pflichtfeld"
        Case "C": lblReq.ForeColor = RGB(160, 100, 0): lblReq.ControlTipText = "bedingt erforderlich"
        Case Else: lblReq.ForeColor = RGB(110, 110, 110): lblReq.ControlTipText = "optional"
    End Select

    Set cbo = container.Controls.Add("Forms.ComboBox.1")
    cbo.Style = 2                           ' fmStyleDropDownList
    cbo.Left = X_BM: cbo.Width = X_VAL - X_BM - 4: cbo.Height = CTL_H
    cbo.ListRows = 20
    cbo.ListWidth = 240
    cbo.List = items
    cbo.ListIndex = 0

    Set txt = container.Controls.Add("Forms.TextBox.1")
    txt.Left = X_VAL: txt.Width = valWidth: txt.Height = CTL_H
End Sub

Public Sub SetTop(ByVal y As Single)
    lblID.Top = y + 2
    lblName.Top = y + 2
    lblReq.Top = y + 2
    cbo.Top = y
    txt.Top = y
End Sub

Public Sub SetVisible(ByVal v As Boolean)
    lblID.Visible = v
    lblName.Visible = v
    lblReq.Visible = v
    cbo.Visible = v
    txt.Visible = v
End Sub

' M = manuelle Eingabe, C = berechnet, B = Textmarke
Public Function SourceKind() As String
    If cbo.ListIndex <= 0 Then
        If kind = "K" Then SourceKind = "C" Else SourceKind = "M"
    Else
        SourceKind = "B"
    End If
End Function

Public Function BookmarkName() As String
    If cbo.ListIndex > 0 Then BookmarkName = cbo.List(cbo.ListIndex)
End Function

' Eigener Inhalt des Feldes (ohne Ersatzwert)
Public Function OwnValue() As String
    Select Case SourceKind
        Case "M": OwnValue = CleanText(txt.Text)
        Case "C": OwnValue = CalcValue
        Case "B": If kind = "K" Then OwnValue = CalcValue Else OwnValue = BmValue
    End Select
    If kind <> "K" Then
        If XR_IsPlaceholder(OwnValue) Then OwnValue = ""
    End If
End Function

' Inhalt, der in die XRechnung übernommen wird
Public Function Value() As String
    Value = OwnValue
    If Len(Value) = 0 Then Value = FallbackValue
End Function

Public Function UsesFallback() As Boolean
    UsesFallback = (Len(OwnValue) = 0 And Len(FallbackValue) > 0)
End Function

' Bei berechneten Feldern: stimmt der Wert der Textmarke mit der Berechnung überein?
Public Function CalcMatches() As Boolean
    Dim a As Double, b As Double, okA As Boolean, okB As Boolean
    a = XR_ParseNum(BmValue, okA)
    b = XR_ParseNum(CalcValue, okB)
    If okA And okB Then
        CalcMatches = (Abs(a - b) < 0.005)
    Else
        CalcMatches = (CleanText(BmValue) = CleanText(CalcValue))
    End If
End Function

Public Sub SetManualText(ByVal s As String)
    mBusy = True
    txt.Text = s
    mBusy = False
End Sub

Public Sub UpdateDisplay()
    mBusy = True
    Select Case SourceKind
        Case "M"
            txt.Locked = False
            txt.BackColor = &H80000005
            txt.ForeColor = &H80000008
        Case "C"
            txt.Text = CalcValue
            txt.Locked = True
            txt.BackColor = RGB(232, 240, 254)
            txt.ForeColor = &H80000008
        Case "B"
            txt.Locked = True
            txt.BackColor = RGB(240, 240, 240)
            If kind = "K" Then
                If CalcMatches Then
                    txt.Text = CalcValue & "   (Textmarke: " & BmValue & " - stimmt überein)"
                    txt.ForeColor = &H80000008
                Else
                    txt.Text = CalcValue & "   (Textmarke: " & BmValue & " - ABWEICHUNG)"
                    txt.ForeColor = RGB(200, 0, 0)
                End If
            Else
                txt.Text = BmValue
                txt.ForeColor = &H80000008
            End If
    End Select
    If UsesFallback Then
        lblName.Caption = BTName & "  [leer: " & FallbackFrom & " wird verwendet]"
        lblName.ForeColor = RGB(160, 100, 0)
        lblName.ControlTipText = id & " - " & BTName & ": leer, verwendet wird " & FallbackFrom & " = " & FallbackValue
    Else
        lblName.Caption = BTName
        lblName.ForeColor = &H80000012
        lblName.ControlTipText = id & " - " & BTName
    End If
    If req = "P" Then
        If Len(Value) > 0 Then lblReq.ForeColor = RGB(0, 150, 0) Else lblReq.ForeColor = RGB(220, 0, 0)
    End If
    mBusy = False
End Sub

Private Sub cbo_Change()
    If mBusy Then Exit Sub
    If Owner.Loading Then Exit Sub
    Owner.OnBTSourceChanged Index
End Sub

Private Sub txt_Change()
    If mBusy Then Exit Sub
    If Owner.Loading Then Exit Sub
    If SourceKind = "M" Then Owner.OnBTValueChanged Index
End Sub

