VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_asientos0 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ASIENTOS CONTABLES"
   ClientHeight    =   8790
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   11985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Numero"
      Height          =   975
      Left            =   5400
      TabIndex        =   20
      Top             =   7200
      Width           =   1695
      Begin VB.TextBox t_f3 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_f4 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha"
      Height          =   975
      Left            =   3600
      TabIndex        =   15
      Top             =   7200
      Width           =   1695
      Begin VB.TextBox t_f2 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_f1 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar por Descripcion/Periodo"
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   7200
      Width           =   3375
      Begin VB.ComboBox c_periodo 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox t_razon 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   14
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   1095
      Left            =   9000
      TabIndex        =   10
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id."
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   9
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton Command6 
         Caption         =   "Cierrre y Apert."
         Height          =   735
         Left            =   6720
         Picture         =   "CGR006A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Renumerar"
         Height          =   735
         Left            =   5400
         Picture         =   "CGR006A.frx":141E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "CGR006A.frx":1728
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "CGR006A.frx":1A32
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "CGR006A.frx":1D3C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "CGR006A.frx":2046
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR006A.frx":2350
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CGR006A.frx":2BD2
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   8535
      Width           =   11985
      _ExtentX        =   21140
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
Attribute VB_Name = "abm_asientos0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
Call limpia
msf1.SetFocus
End Sub

Private Sub c_periodo_LostFocus()
If c_periodo.ListIndex < 0 Then
  c_periodo.ListIndex = buscaindice(c_periodo, para.id_periodo_contable)
End If
End Sub

Private Sub Command4_Click()
Dim c(15) As Double
Call nivel_acceso(1)
If para.id_grupo_modulo_actual >= 3 Then
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5
    c(6) = 7
    For i = 7 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE ASIENTOS", "", "Periodo......: " & t_f1 & "  " & t_f2, "Intervalo....: " & t_f3 & "  " & t_f4, 75, 8, True, False, "V")
  End If
Else
 Call sinpermisos
End If

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Private Sub Command1_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
 abm_asientos.limpia
 abm_asientos.t_funcion = "A"
 abm_asientos.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
 Call nivel_acceso(7)
 If para.id_grupo_modulo_actual >= 5 Then
   abm_asientos.limpia
   abm_asientos!t_funcion = "M"
   Call LLENACAMPOS
 Else
   Call sinpermisos
 End If
End If

Exit Sub
e1:
 Exit Sub
End Sub

Sub LLENACAMPOS()
'On Error GoTo ERROR1
q = "select * from c_11 where [id_asiento] = " & Val(msf1.TextMatrix(msf1.Row, 0))
Set rs = New ADODB.Recordset
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  abm_asientos.t_id = rs("id_asiento")
  abm_asientos.t_numero = rs("num_asiento")
  abm_asientos.t_descripciong = rs("descripcion")
  abm_asientos.t_f1 = rs("fecha")
  abm_asientos.armagrid
  abm_asientos.armagrid2
  q = "select * from c_12, c_01 where [id_asiento] = " & Val(msf1.TextMatrix(msf1.Row, 0)) & " and c_12.[id_cuenta] = c_01.[id_cuenta]  order by [secuencia]"
  Set rs1 = New ADODB.Recordset
  rs1.Open q, cn1
  While Not rs1.EOF
    If rs1("ubicacion") = "D" Then
       abm_asientos.msf1.AddItem abm_asientos.msf1.Row & Chr$(9) & rs1("c_12.id_cuenta") & Chr$(9) & rs1("c_12.descripcion") & Chr$(9) & rs1("c_12.importe") & Chr$(9) & rs1("c_01.descripcion")
    Else
       abm_asientos.msf2.AddItem abm_asientos.msf2.Row & Chr$(9) & rs1("c_12.id_cuenta") & Chr$(9) & rs1("c_12.descripcion") & Chr$(9) & rs1("c_12.importe") & Chr$(9) & rs1("c_01.descripcion")
    End If
    rs1.MoveNext
  Wend
End If
Set rs = Nothing
Set rs1 = Nothing
abm_asientos.calcula_totales
abm_asientos.Show

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Asientos. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
 Call nivel_acceso(7)
 If para.id_grupo_modulo_actual >= 7 Then
   J = MsgBox("Confirma borrar asiento Nro: " & msf1.TextMatrix(msf1.Row, 3), 4)
   If J = 6 Then
      QUERY = "DELETE FROM c_11 WHERE [id_asiento] = " & Val(msf1.TextMatrix(msf1.Row, 0))
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      Call limpia
   End If
 
 Else
  Call sinpermisos
 End If
End If

Exit Sub
e1:
 Exit Sub
End Sub


Private Sub Command5_Click()
If c_periodo.ListIndex > 0 Then
  J = MsgBox("El proceso reenumera todos los asientos del periodo contable selecionado. ¿Confirma?", 4)
  If J = 6 Then
    Call nivel_acceso(7)
    If para.id_grupo_modulo_actual >= 8 Then
      espere.Show
      espere.Refresh
       Call renumera
      Unload espere
    End If
  End If
Else
  MsgBox ("Debe tener seleccionado un periodo contable para poder renumerar asientos")
End If
End Sub
Sub renumera()
q = "select * from c_11 where [id_periodo] = " & c_periodo.ItemData(c_periodo.ListIndex) & " order by [fecha], [num_asiento]"
Set rs = New ADODB.Recordset
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
p = 0
s = 1
While Not rs.EOF
  pa = Val(Mid$(Format$(rs("num_asiento"), "000000000"), 1, 6))
  If pa <> p Then
     s = 1
     p = pa
  End If
  na = Val(Format$(pa, "000000") & Format$(s, "000"))
  rs("num_asiento") = na
  rs.Update
  s = s + 1
  rs.MoveNext
  espere.Label1 = s
  espere.Label1.Refresh
Wend


End Sub

Private Sub Command6_Click()
cgr_cierreyapertura.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
Dim q As String
q = "select * from c_11, c_10 where c_11.[id_periodo] = c_10.[id_periodo] "
c = " and "
If t_razon <> "" Then
   q = q & c & " c_11.[descripcion] like '%" & t_razon & "%'"
   c = " and "
End If

If t_f1 <> "" Then
   q = q & c & " datevalue([fecha]) >= datevalue('" & t_f1 & "')"
   c = " and "
End If

If t_f2 <> "" Then
   q = q & c & " datevalue([fecha]) <= datevalue('" & t_f2 & "')"
   c = " and "
End If

If t_f3 <> "" Then
   q = q & c & " [num_asiento] >= " & Val(t_f3)
   c = " and "
End If

If t_f4 <> "" Then
   q = q & c & " [num_asiento] <= " & Val(t_f4)
   c = " and "
End If

If c_periodo.ListIndex > 0 Then
   q = q & c & " c_11.[id_periodo] = " & c_periodo.ItemData(c_periodo.ListIndex)
   c = " and "
End If

If Option1 = True Then
  q = q & " order by [id_asiento]"
Else
  If Option2 = True Then
    q = q & " order by [fecha], [id_asiento]"
  Else
    If Option3 = True Then
      q = q & " order by [num_asiento]"
    Else
      q = q & " order by c_11.[descripcion], [id_asiento]"
    End If
  End If
End If

Call armagrid
Set rs = New ADODB.Recordset
rs.Open q, cn1
t = 0
ca = 0
While Not rs.EOF
  
  msf1.AddItem Format$(rs("id_asiento"), "00000") & Chr$(9) & rs("fecha") & Chr$(9) & rs("c_11.descripcion") & Chr$(9) & Format$(rs("num_asiento"), "000000000") & Chr$(9) & Format$(rs("importe"), "######0.00") & Chr$(9) & rs("c_11.id_periodo") & Chr$(9) & rs("c_10.descripcion")
  t = t + rs("importe")
  ca = ca + 1
  rs.MoveNext
Wend
msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "=====================" & Chr$(9) & "" & Chr$(9) & ""
msf1.AddItem "" & Chr$(9) & "Asientos --->" & Chr$(9) & ca & Chr$(9) & "" & Chr$(9) & Format$(t, "######0.00") & Chr$(9) & "" & Chr$(9) & ""
If Not rs.EOF And Not rs.BOF Then
  msf1.RowSel = 1
End If
Set rs = Nothing
Call INICIALIZA2(abm_asientos)
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 1200
msf1.ColWidth(2) = 4000
msf1.ColWidth(3) = 1400
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 600
msf1.ColWidth(6) = 2000
msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Fecha"
msf1.TextMatrix(0, 2) = "Descripcion"
msf1.TextMatrix(0, 3) = "Numero"
msf1.TextMatrix(0, 4) = "Importe"
msf1.TextMatrix(0, 5) = ""
msf1.TextMatrix(0, 6) = "Periodo"



For i = 0 To 6
 msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(4) = 9

End Sub

Private Sub Form_Load()
Call barracgr(Me)
Option3 = True
Call armagrid

Call carga_periodos(c_periodo)
c_periodo.AddItem "<Todos>", 0
c_periodo.ListIndex = buscaindice(c_periodo, para.id_periodo_contable)

Load abm_asientos
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_asientos
End Sub


Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 On Error GoTo e1
 If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
 Call nivel_acceso(1)
 If para.id_grupo_modulo_actual >= 5 Then
   abm_asientos!t_funcion = "C"
   Call LLENACAMPOS
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

Private Sub t_f3_GotFocus()
t_f3 = ""

End Sub

Private Sub t_f4_GotFocus()
t_f2 = ""
End Sub

Private Sub t_razon_GotFocus()
t_razon = ""
End Sub

