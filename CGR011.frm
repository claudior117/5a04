VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cgr_diario 
   BackColor       =   &H00E0E0E0&
   Caption         =   "LIBRO DIARIO"
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   11985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Numero"
      Height          =   975
      Left            =   5400
      TabIndex        =   15
      Top             =   120
      Width           =   1695
      Begin VB.TextBox t_f3 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_f4 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha"
      Height          =   975
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   1695
      Begin VB.TextBox t_f2 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_f1 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar por Descripcion/Periodo"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox c_periodo 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox t_razon 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   615
      Left            =   7440
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
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
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR011.frx":0000
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
         Picture         =   "CGR011.frx":0882
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
      Top             =   8415
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
            TextSave        =   "21/01/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:23 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_diario"
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
ca = 0
dt = 0
ht = 0
While Not rs.EOF
  nra = "Nro. As:  " & rs("num_asiento")
  f = rs("fecha")
  Desc = "Descrip:  " & rs("c_11.descripcion")
  msf1.AddItem nra & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "Fecha: " & f
  msf1.AddItem Desc
  msf1.AddItem ""
  
  Set rs2 = New ADODB.Recordset
  q = "select * from c_12, C_01 where [id_asiento] = " & rs("id_asiento") & " AND c_12.[id_cuenta] = c_01.[id_cuenta]  order by [ubicacion], [secuencia]"
  rs2.Open q, cn1
  da = 0
  ha = 0
  While Not rs2.EOF
     If rs2("ubicacion") = "D" Then
         e = ""
         d = Format$(rs2("c_12.importe"), "######0.00")
         h = ""
         da = da + Val(d)
         
     Else
         e = "               "
         h = Format$(rs2("c_12.importe"), "######0.00")
         ha = ha + Val(h)
         d = ""
     End If
     desca = rs2("c_12.descripcion")
     dc = rs2("c_01.descripcion")
     cc = rs2("c_12.id_cuenta")
     msf1.AddItem e & cc & "  " & dc & Chr$(9) & d & Chr$(9) & h & Chr$(9) & desca
     rs2.MoveNext
  Wend
  lc = "---------------------------"
  dat = Format$(da, "#######0.00")
  hat = Format$(ha, "#######0.00")
  msf1.AddItem "  " & Chr$(9) & lc & Chr$(9) & lc
  msf1.AddItem "  " & Chr$(9) & dat & Chr$(9) & hat
  msf1.AddItem "  "
  ht = ht + Val(hat)
  dt = dt + Val(dat)
  ca = ca + 1
  rs.MoveNext
  Set rs2 = Nothing
  
Wend
lc = "==========================="
msf1.AddItem "  " & Chr$(9) & lc & Chr$(9) & lc
msf1.AddItem "Asientos --->  " & ca & Chr$(9) & Format$(dt, "#######0.00") & Chr$(9) & Format$(ht, "#######0.00")
msf1.AddItem "  " & Chr$(9) & lc & Chr$(9) & lc
Set rs = Nothing
Set rs2 = Nothing
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 5
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 6500
msf1.ColWidth(1) = 1500
msf1.ColWidth(2) = 1500
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 500
msf1.TextMatrix(0, 0) = "ASIENTO"
msf1.TextMatrix(0, 1) = "DEBE"
msf1.TextMatrix(0, 2) = "HABER"
msf1.TextMatrix(0, 3) = ""
msf1.TextMatrix(0, 4) = ""
For i = 1 To 2
 msf1.ColAlignment(i) = 9 'DER
Next i
msf1.ColAlignment(0) = 1
msf1.ColAlignment(3) = 1
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


Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 4
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    For i = 5 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "Libro Diario", "", "Periodo......: " & t_f1 & "  " & t_f2, "Ejercicio.........:" & c_periodo, 87, 7, True, False)
  End If
    
    
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

