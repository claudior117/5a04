VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_mayores 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MAYORES POR CUENTA"
   ClientHeight    =   8775
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   12135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha"
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_id 
         Height          =   285
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox t_f2 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox t_f1 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Cuenta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar por Descripcion/Periodo"
      Height          =   735
      Left            =   8520
      TabIndex        =   11
      Top             =   840
      Width           =   3255
      Begin VB.ComboBox c_periodo 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   615
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   8
      Top             =   1680
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
      TabIndex        =   5
      Top             =   7440
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR007.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CGR007.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   4
      Top             =   8520
      Width           =   12135
      _ExtentX        =   21405
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
Attribute VB_Name = "cgr_mayores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
Call limpia
msf1.SetFocus
End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex >= 0 Then
  t_id = c_cuenta.ItemData(c_cuenta.ListIndex)
Else
  t_id = ""
End If
End Sub

Private Sub c_periodo_LostFocus()
If c_periodo.ListIndex < 0 Then
  c_periodo.ListIndex = buscaindice(c_periodo, para.id_periodo_contable)
End If
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
Call armagrid
Dim q As String
d = 0
h = 0
sa = 0
dt = 0
ht = 0
st = 0
If t_f1 <> "" Then
  q = "select * from c_11, c_12 where c_11.[id_asiento] = c_12.[id_asiento] and [id_cuenta] = " & Val(t_id)
  q = q & " and datevalue([fecha]) < datevalue('" & t_f1 & "')"
  c = " and "
  If c_periodo.ListIndex > 0 Then
   q = q & c & " c_11.[id_periodo] = " & c_periodo.ItemData(c_periodo.ListIndex)
   c = " and "
  End If
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  While Not rs.EOF
   If rs("ubicacion") = "D" Then
    d = d + rs("c_12.importe")
   Else
    h = h + rs("c_12.importe")
   End If
   rs.MoveNext
  Wend
  sa = d - h
End If
msf1.AddItem "" & Chr$(9) & t_f1 & Chr$(9) & "" & Chr$(9) & "Saldo Anterior" & Chr$(9) & Format$(d, "######0.00") & Chr$(9) & Format$(h, "######0.00") & Chr$(9) & Format$(sa, "######0.00")
dt = d
ht = h

q = "select * from c_11, c_12 where c_11.[id_asiento] = c_12.[id_asiento] and [id_cuenta] = " & Val(t_id)
c = " and "
If t_f1 <> "" Then
   q = q & c & " datevalue([fecha]) >= datevalue('" & t_f1 & "')"
   c = " and "
End If

If t_f2 <> "" Then
   q = q & c & " datevalue([fecha]) <= datevalue('" & t_f2 & "')"
   c = " and "
End If


If c_periodo.ListIndex > 0 Then
   q = q & c & " c_11.[id_periodo] = " & c_periodo.ItemData(c_periodo.ListIndex)
   c = " and "
End If


If Option2 = True Then
    q = q & " order by [fecha], c_11.[id_asiento]"
Else
  If Option3 = True Then
     q = q & " order by [num_asiento]"
  Else
      q = q & " order by c_11.[descripcion], c_11.[id_asiento]"
  End If
End If

Set rs = New ADODB.Recordset
rs.Open q, cn1
d = 0
h = 0
s = sa
While Not rs.EOF
  If rs("ubicacion") = "D" Then
    d = rs("c_12.importe")
  Else
    h = rs("c_12.importe")
  End If
  s = s + (d - h)
  msf1.AddItem Format$(rs("c_11.id_asiento"), "00000") & Chr$(9) & Format$(rs("fecha"), "dd/mm/yyyy") & Chr$(9) & rs("num_asiento") & Chr$(9) & rs("c_12.descripcion") & Chr$(9) & Format$(d, "######0.00") & Chr$(9) & Format$(h, "######0.00") & Chr$(9) & Format$(s, "######0.00")
  rs.MoveNext
  dt = dt + d
  ht = ht + h
  d = 0
  h = 0

Wend
Set rs = Nothing
msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "-----------------" & Chr$(9) & "-----------------" & Chr$(9) & "-----------------"
msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "Totales--->" & Chr$(9) & Format$(dt, "######0.00") & Chr$(9) & Format$(ht, "######0.00") & Chr$(9) & Format$(dt - ht, "######0.00")

End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 1200
msf1.ColWidth(2) = 1400
msf1.ColWidth(3) = 4000
msf1.ColWidth(4) = 1400
msf1.ColWidth(5) = 1400
msf1.ColWidth(6) = 1400
msf1.ColWidth(7) = 700
msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Fecha"
msf1.TextMatrix(0, 2) = "Asiento"
msf1.TextMatrix(0, 3) = "Detalle"
msf1.TextMatrix(0, 4) = "Debe"
msf1.TextMatrix(0, 5) = "Haber"
msf1.TextMatrix(0, 6) = "Saldo"
msf1.TextMatrix(0, 7) = ""
For i = 0 To 3
 msf1.ColAlignment(i) = 1 'izq
Next i

For i = 4 To 6
 msf1.ColAlignment(i) = 9 'izq
Next i

End Sub

Private Sub Form_Load()
Call barracgr(Me)
Option2 = True
Call armagrid

Call carga_periodos(c_periodo)
c_periodo.AddItem "<Todos>", 0
c_periodo.ListIndex = buscaindice(c_periodo, para.id_periodo_contable)

Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_asientos
End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F11] Excel - [ENTER] Muestra Asiento  "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual >= 4 Then
    J = MsgBox("Prepare Impresora y Confirme", 4)
    If J = 6 Then
     Dim c(15) As Double
      c(0) = 7
      c(1) = 1
      c(2) = 2
      c(3) = 3
      c(4) = 4
      c(5) = 5
      c(6) = 6
      For i = 7 To 14
        c(i) = -1
      Next i
      Call imprimegrid(msf1, c(), "MAYORES POR CUENTA", "", "Cuenta: (" & t_id & ") " & c_cuenta, "Periodo: " & t_f1 & "  " & t_f2, 87, 7, True, False)
    End If
         
  End If
  
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
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


Private Sub t_id_LostFocus()
If t_id <> "" Then
  Set rs = New ADODB.Recordset
  q = "select * from c_01 where [id_cuenta] = " & Val(t_id)
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
      c_cuenta.ListIndex = buscaindice(c_cuenta, Val(t_id))
  Else
     t_id = ""
  End If
  Set rs = Nothing
End If
End Sub
