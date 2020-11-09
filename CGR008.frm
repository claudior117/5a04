VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_balance 
   BackColor       =   &H00E0E0E0&
   Caption         =   "BALANCE DE SUMAS Y SALDOS"
   ClientHeight    =   8805
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opcion"
      Height          =   615
      Left            =   4320
      TabIndex        =   10
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detallado"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha"
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
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
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
         Picture         =   "CGR008.frx":0000
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
         Picture         =   "CGR008.frx":0882
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_balance"
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
q = "select * from c_01 where [tipo] = 'C' order by [id_cuenta]"
Set rs = New ADODB.Recordset
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
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
  'para cada cuenta totalizo
  c = c + 1
  espere.ProgressBar1.Value = c
  Set rs1 = New ADODB.Recordset
  q = "select * from c_11, c_12 where c_11.[id_asiento] = c_12.[id_asiento] and [id_cuenta] = " & rs("id_cuenta")
  q = q & p
  rs1.Open q, cn1
  i = 0
  While Not rs1.EOF
    If rs1("ubicacion") = "D" Then
       i = i + rs1("c_12.importe")
    Else
       i = i - rs1("c_12.importe")
    End If
    rs1.MoveNext
  Wend
  Set rs1 = Nothing
  rs("importe") = i
  rs.Update
  rs.MoveNext
Wend
Set rs = Nothing
q = "select * from c_01 where [tipo] = 'T' order by [id_cuenta]"
Set rs = New ADODB.Recordset
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
  c = c + 1
  espere.ProgressBar1.Value = c
  ro = Val(Mid$(Format$(rs("ID_CUENTA"), "000000"), 2, 5))
  If ro = 0 Then
    pi = Val(Mid$(rs("id_cuenta"), 1, 1) & "00000")
    pf = Val(Mid$(rs("id_cuenta"), 1, 1) & "99999")
    ic = " and  c_12.[id_cuenta] >= " & pi & " and c_12.[id_cuenta] <= " & pf
  Else
    ro = Val(Mid$(Format$(rs("ID_CUENTA"), "000000"), 3, 4))
    If ro = 0 Then
      pi = Val(Mid$(rs("id_cuenta"), 1, 2) & "0000")
      pf = Val(Mid$(rs("id_cuenta"), 1, 2) & "9999")
      ic = " and c_12.[id_cuenta] >= " & pi & " and c_12.[id_cuenta] <= " & pf
    Else
      pi = Val(Mid$(rs("id_cuenta"), 1, 4) & "00")
      pf = Val(Mid$(rs("id_cuenta"), 1, 4) & "99")
      ic = " and c_12.[id_cuenta] >= " & pi & " and c_12.[id_cuenta] <= " & pf
    End If
  End If
  q = "select * from c_11, c_12 where c_11.[id_asiento] = c_12.[id_asiento] "
  q = q & p & ic
  Set rs1 = New ADODB.Recordset
  rs1.Open q, cn1
  i = 0
  While Not rs1.EOF
    If rs1("ubicacion") = "D" Then
       i = i + rs1("c_12.importe")
    Else
       i = i - rs1("c_12.importe")
    End If
    rs1.MoveNext
  Wend
  Set rs1 = Nothing
  rs("importe") = i
  rs.Update
  rs.MoveNext
Wend
Set rs = Nothing



'muestro
q = "select * from c_01 order by [id_cuenta]"
Set rs = New ADODB.Recordset
rs.Open q, cn1
l = "---------------------"
lf = "-------------------->"
l2 = "---------------------------------------------------------------------------"
T2 = "               "
While Not rs.EOF
  If rs("tipo") = "C" Then
    If Option1 Then
       msf1.AddItem rs("id_cuenta") & Chr$(9) & T2 & "      " & rs("descripcion") & Chr$(9) & Format$(rs("importe"), "######0.00")
    End If
  Else
    ro = Val(Mid$(Format$(rs("ID_CUENTA"), "000000"), 2, 5))
    If ro = 0 Then
      msf1.AddItem ""
      msf1.AddItem rs("descripcion") & Chr$(9) & l2 & Chr$(9) & l & Chr$(9) & l & Chr$(9) & lf & Chr$(9) & Format$(rs("importe"), "######0.00")
    Else
       ro = Val(Mid$(Format$(rs("ID_CUENTA"), "000000"), 3, 5))
       If ro = 0 Then
          msf1.AddItem ""
          msf1.AddItem "" & Chr$(9) & rs("descripcion") & l2 & Chr$(9) & l & Chr$(9) & lf & Chr$(9) & Format$(rs("importe"), "######0.00")
       Else
          msf1.AddItem ""
          msf1.AddItem "" & Chr$(9) & T2 & rs("descripcion") & l2 & Chr$(9) & lf & Chr$(9) & Format$(rs("importe"), "######0.00")
       End If
    End If
  End If
  rs.MoveNext
Wend
Set rs = Nothing
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
msf1.ColWidth(0) = 2000
msf1.ColWidth(1) = 4500
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 1200

msf1.TextMatrix(0, 0) = "Cuenta"
msf1.TextMatrix(0, 1) = "Detalle"
msf1.TextMatrix(0, 2) = "Importe"
msf1.TextMatrix(0, 3) = ""
msf1.TextMatrix(0, 4) = ""
msf1.TextMatrix(0, 5) = ""

For i = 0 To 1
 msf1.ColAlignment(i) = 1 'izq
Next i

For i = 2 To 5
 msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()
Call barracgr(Me)
Call armagrid
Option1 = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_asientos
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
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 On Error GoTo e1
 If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
 Call nivel_acceso(1)
 If para.id_grupo_modulo_actual >= 5 Then
   Load abm_asientos0
   abm_asientos0.t_f3 = msf1.TextMatrix(msf1.Row, 2)
   abm_asientos0.t_f4 = msf1.TextMatrix(msf1.Row, 2)
   abm_asientos0.limpia
   abm_asientos0.Show
 Else
   Call sinpermisos
 End If
End If

Exit Sub
e1:
 Exit Sub
End If
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


