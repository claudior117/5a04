VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form exp_lista 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "BUSCADOR DE OPERACIONES DE EXPORTACION"
   ClientHeight    =   8640
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8640
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Embarque"
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   4335
      Begin VB.TextBox t_fechaf 
         Height          =   285
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Posterior a"
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Anterior a"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar"
      Height          =   855
      Left            =   10080
      TabIndex        =   17
      Top             =   7320
      Width           =   1695
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   960
         Picture         =   "exp003.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   495
         Left            =   120
         Picture         =   "exp003.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   9360
      TabIndex        =   15
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   120
         Picture         =   "exp003.frx":1104
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   4815
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   4695
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   8281
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenado por"
      Height          =   1335
      Left            =   7680
      TabIndex        =   10
      Top             =   840
      Width           =   2175
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Expedicion"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Basico"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Registros"
      Height          =   735
      Left            =   7680
      TabIndex        =   8
      Top             =   0
      Width           =   1575
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   13
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox c_grupo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox t_detalle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox t_basico 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Num. OP."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8280
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   26458
            MinWidth        =   26458
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "exp_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub carga()
espere.Show
espere.Label1 = "Cargando lista de precios...."
espere.Refresh
 Call armagrid
 ct = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select * from exp01 "
 c = " where "
 
 If t_basico <> "" Then
   q = q & c & "[num_exp] = " & Val(t_basico)
   c = " and "
 End If
 
 
 If t_detalle <> "" Then
   q = q & c & "[detalle] like  '%" & t_detalle & "%'"
   c = " and "
 End If
 
 If c_grupo.ListIndex > 0 Then
   q = q & c & "[id_cliente] = " & c_grupo.ItemData(c_grupo.ListIndex)
   c = " and "
 End If
 
  
 If t_fechaf <> "" Then
    q = q & c & " datevalue([fecha_embarque]) "
    If Option3 = True Then
       q = q & " <= "
    Else
       q = q & " >= "
    End If
    q = q & " datevalue('" & t_fechaf & "')"
    c = " and "
 End If
 
 
 
 If Option1 = True Then
   q = q & " order by [num_exp]"
 Else
  If Option2 = True Then
     q = q & " order by [detalle]"
  Else
    If Option5 = True Then
       q = q & " order by [cliente]"
    Else
    q = q & " order by [fecha_embarque]"
    End If
  End If
 End If

rs.Open q, cn1
t_encontrados = 0
While Not rs.EOF
    tr = sacareintegro(rs("num_exp"))
    b = Format$(rs("num_exp"), "00000")
    d = rs("detalle")
    p = Format$(rs("importe"), "######0.00")
    fp = rs("fecha_embarque")
    c = rs("cliente")
    msf1.AddItem b & Chr$(9) & d & Chr$(9) & c & Chr$(9) & fp & Chr$(9) & p & Chr$(9) & tr & Chr$(9) & rs("num_exp") & Chr$(9) & rs("id_cliente")
    t_encontrados = Val(t_encontrados) + 1
    rs.MoveNext
 Wend
msf1.SetFocus
Set rs = Nothing
Unload espere
End Sub


Private Sub btnacepta_Click()

Call carga

End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub


Private Sub c_grupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub



Private Sub Command2_Click()
exp_exporta.Show
End Sub

Private Sub Form_Activate()
para.exporta_sel = 0
If msf1.Rows > 1 Then
 msf1.SetFocus
Else
 t_detalle.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Me.Hide
End If

End Sub

Private Sub Form_Load()
  Call carga_clientes(c_grupo)
  c_grupo.AddItem "<Todos>", 0
  c_grupo.ListIndex = 0
  Call armagrid
  Check1 = 0
  Option6 = True
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 4500
msf1.ColWidth(2) = 2400
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 1000
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 0
msf1.ColWidth(7) = 0
msf1.ColWidth(8) = 400




msf1.TextMatrix(0, 0) = "Basico"
msf1.TextMatrix(0, 1) = "Descripcion"
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = "Embarque"
msf1.TextMatrix(0, 4) = "Reintegro"
msf1.TextMatrix(0, 5) = "Ingresado"
msf1.TextMatrix(0, 6) = "Id.Exportacion"
msf1.TextMatrix(0, 7) = "Id.Cliente"
msf1.TextMatrix(0, 8) = ""


For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 6
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

  




Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(1) = "[F3] Reintegro -  [F4] Saca - [F7] Imprime  - [Esc] Cancela"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
 If msf1.Rows > 0 Then
   Load exp_prodreintegro
   exp_prodreintegro.c_prov.ListIndex = buscaindice(exp_prodreintegro.c_prov, Val(msf1.TextMatrix(msf1.Row, 7)))
   Call exp_prodreintegro.carga_exportaciones(exp_prodreintegro.c_vend)
   exp_prodreintegro.c_vend.ListIndex = buscaindice(exp_prodreintegro.c_vend, Val(msf1.TextMatrix(msf1.Row, 6)))
   exp_prodreintegro.carga
   exp_prodreintegro.Show
   
   
   
  End If
End If



If KeyCode = vbKeySpace Then
  t_detalle.SetFocus
End If





If KeyCode = vbKeyF4 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    msf1.RemoveItem r
    t_encontrados = Val(t_encontrados) - 1
  End If
End If

If KeyCode = vbKeyF6 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    vta_listaprecios4.Show
    vta_listaprecios4.t_idprod = msf1.TextMatrix(r, 0)
    vta_listaprecios4.t_prod = msf1.TextMatrix(r, 1)
    
  End If
End If


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
    
    If t_detalle <> "" Then
      t = "Detalle: " & t_detalle
    Else
      t = ""
    End If
    
    If c_grupo.ListIndex > 0 Then
       t1 = "Cliente: " & c_grupo
    End If
    Call imprimegrid(msf1, c(), "LISTADO DE EXPORTACIONES", "", t, t1, 80, 8, True, False, "V")
  End If
End If








End Sub

Private Sub msf1_LostFocus()
'Call barra(Me)
End Sub

Private Sub t_basico_GotFocus()
t_basico = ""
End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub


Private Sub T_detalle_GotFocus()
t_detalle = ""
End Sub

Private Sub t_detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub t_fechaf_GotFocus()
t_fechaf = ""
End Sub

Private Sub t_fechaf_LostFocus()
If t_fechaf <> "" Then
 If Not IsDate(t_fechaf) Then
   t_fechaf = ""
 End If
End If
End Sub
