VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form cgr_sumasysaldosp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "BALANCE de COMPROBANCION DE SUMAS y SALDOS"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fechas  Corte"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      Begin VB.TextBox t_f2 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_f1 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR016.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CGR016.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   8550
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "06/06/2025"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:11 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_sumasysaldosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
Call limpia
msf1.SetFocus
End Sub



Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
'saldo anterior
espere.Show
espere.Refresh
Call armagrid
Dim q As String
q = "select * from c_01 where [tipo] = 'C'"
If Option6 = True Then
  q = q & " and [pos1] <= 2 "
End If

q = q & " order by [id_cuenta]"

Set rs = New ADODB.Recordset
rs.Open q, cn1
p = ""
If t_f1 <> "" Then
   p = " and datevalue([fecha]) >= datevalue('" & t_f1 & "')"
End If
If t_f2 <> "" Then
   p = p & " and datevalue([fecha]) <= datevalue('" & t_f2 & "')"
End If
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 1000
c = 0
While Not rs.EOF
  c = c + 1
  espere.ProgressBar1.Value = c
  Set rs1 = New ADODB.Recordset
  q = "select * from c_02,  c_03 where c_02.[num_interno] = c_03.[num_interno] and  [id_cuenta] = " & rs("id_cuenta")
  q = q & p
  rs1.Open q, cn1
  i = 0
  td = 0
  th = 0
  While Not rs1.EOF
    If rs1("ubicacion") = "D" Then
       td = td + rs1("importe")
    Else
       th = th + rs1("importe")
    End If
    rs1.MoveNext
  Wend
  Set rs1 = Nothing
  
  If td > 0 Or th > 0 Then
    s = td - th
    If s > 0 Then
      sd = Format$(s, "######0.00")
      sh = ""
    Else
      If s < 0 Then
        sh = Format$(-s, "######0.00")
        sd = ""
      Else
        sd = ""
        sh = ""
      End If
    End If
    
    msf1.AddItem rs("id_cuenta") & Chr$(9) & rs("DESCRIPCION") & Chr$(9) & Format$(td, "######0.00") & Chr$(9) & Format$(th, "######0.00") & Chr$(9) & Format$(sd, "######0.00") & Chr$(9) & Format$(sh, "######0.00")
  
  End If
  rs.MoveNext
Wend
Set rs = Nothing
l = "------------------------------"
msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & l & Chr$(9) & l & Chr$(9) & l & Chr$(9) & l
msf1.AddItem "" & Chr$(9) & "TOTALES....." & Chr$(9) & Format$(suma_msflexgrid(msf1, 2), "#######0.00") & Chr$(9) & Format$(suma_msflexgrid(msf1, 3), "#######0.00") & Chr$(9) & Format$(suma_msflexgrid(msf1, 4), "#######0.00") & Chr$(9) & Format$(suma_msflexgrid(msf1, 5), "#######0.00")





Unload espere
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 3500
msf1.ColWidth(2) = 1500
msf1.ColWidth(3) = 1500
msf1.ColWidth(4) = 1500
msf1.ColWidth(5) = 1500

msf1.TextMatrix(0, 0) = "Nro."
msf1.TextMatrix(0, 1) = "Cuenta"
msf1.TextMatrix(0, 2) = "Sumas Debe"
msf1.TextMatrix(0, 3) = "Sumas Haber"
msf1.TextMatrix(0, 4) = "Saldo Deudor"
msf1.TextMatrix(0, 5) = "Saldo Acred."

For i = 0 To 1
 msf1.ColAlignment(i) = 1 'izq
Next i

For i = 2 To 5
 msf1.ColAlignment(i) = 9 'izq
Next i

End Sub

Private Sub Form_Load()
Call barracgr(Me)
Call armagrid
Option1 = True
Option3 = True
Option6 = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_asientos
End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime -  [F11] Excel "
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5
    
    For i = 6 To 14
      c(i) = -1
    Next i
    If Option1 Then
      t = "Detallado"
    Else
      t = "Resumido"
    End If
    Call imprimegrid(msf1, c(), "Balance de Sumas y Saldos", "", "Periodo......: " & t_f1 & "  " & t_f2, "Tipo.........:" & t, 85, 7, True, False)
  End If
    
    
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 'On Error GoTo e1
 'If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
 'Call nivel_acceso(1)
 'If para.id_grupo_modulo_actual >= 5 Then
 '  Load abm_asientos0
 '  abm_asientos0.t_f3 = msf1.TextMatrix(msf1.Row, 2)
 '  abm_asientos0.t_f4 = msf1.TextMatrix(msf1.Row, 2)
 '  abm_asientos0.limpia
 '  abm_asientos0.Show
 'Else
 '  Call sinpermisos
 'End If
End If

Exit Sub
e1:
 Exit Sub
'End If
End Sub

Private Sub t_f1_GotFocus()
t_f1 = ""

End Sub

Private Sub t_f1_LostFocus()
If t_f1 <> "" Then
 If Not IsDate(t_f1) Then
   t_f1 = ""
 End If
End If
 
End Sub

Private Sub t_f2_GotFocus()
t_f2 = ""

End Sub

Private Sub t_f2_LostFocus()
If t_f2 <> "" Then
 If Not IsDate(t_f2) Then
   t_f2 = ""
 End If
End If

End Sub


