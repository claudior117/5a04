VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form con_verprod 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PRODUCTOS INGRESADOS"
   ClientHeight    =   8595
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   11535
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Text            =   "c_prod"
         Top             =   1080
         Width           =   4815
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Productos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tipo Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6960
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   6
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Con006.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Con006.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   7
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
      TabIndex        =   5
      Top             =   8340
      Width           =   12270
      _ExtentX        =   21643
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
            TextSave        =   "09:43"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "con_verprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer


Sub carga()
  Call armagrid
  q = "select * from a5, a6, g2, a1, g1 where a5.[num_int] = a6.[num_int] and [id_tipocomp] = [id_tipo_comp] and a5.[id_proveedor] = a1.[id_proveedor] and a5.[id_usuario] = g1.[id_usuario] "
  c = " and "
  If c_prov.ListIndex > 0 Then
     q = q & c & " a5.[id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  If c_tipocomp.ListIndex > 0 Then
    q = q & c & " [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex)
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue(a5.[fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue(a5.[fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  If c_prod.ListIndex < 0 Then
    If c_prod <> "" Then
     q = q & c & " [detalle] like '%" & c_prod & "%'"
    End If
  Else
    q = q & c & " [id_producto] = " & c_prod.ItemData(c_prod.ListIndex)
  End If
  
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  t = 0
  While Not rs.EOF
     F = rs("a5.fecha")
     CTC = Format$(rs("ID_TIPOCOMP"), "000")
     tc = rs("descripcion")
     nc = rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     d = rs("detalle")
     cp = Format$(rs("a5.id_proveedor"), "0000")
     p = rs("denominacion")
     u = rs("usuario")
     ni = rs("a5.num_int")
     pu = Format$(rs("pu"), "#####0.00")
     c = Format$(rs("cantidad"), "#####0.00")
     msf1.AddItem F & Chr(9) & cp & Chr(9) & p & Chr(9) & CTC & Chr(9) & tc & Chr(9) & nc & Chr(9) & d & Chr(9) & rs("a5.num_int") & Chr(9) & u & Chr(9) & pu & Chr(9) & c
     rs.MoveNext
  Wend
  
   
End Sub

Private Sub btnacepta_Click()
Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 11
msf1.ColWidth(0) = 1300
msf1.ColWidth(1) = 700 'cod prov
msf1.ColWidth(2) = 3500
msf1.ColWidth(3) = 500
msf1.ColWidth(4) = 1700
msf1.ColWidth(5) = 1700
msf1.ColWidth(6) = 2500
msf1.ColWidth(7) = 1000
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 1000
msf1.ColWidth(10) = 1000
msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = ""
msf1.TextMatrix(0, 2) = "Proveedor"
msf1.TextMatrix(0, 3) = ""
msf1.TextMatrix(0, 4) = "Operacion"
msf1.TextMatrix(0, 5) = "Nro.Comprobante"
msf1.TextMatrix(0, 6) = "Producto"
msf1.TextMatrix(0, 7) = "Num.Int."
msf1.TextMatrix(0, 8) = "Usuario"
msf1.TextMatrix(0, 9) = "Precio s/ iva"
msf1.TextMatrix(0, 10) = "Cantidad"

End Sub









Private Sub c_prod_LostFocus()
If c_prod.ListIndex < 0 Then
  If Val(c_prod) > 0 Then
    c_prod.ListIndex = buscaindice(c_prod, Val(c_prod))
  Else
    c_prod.ListIndex = 0
  End If
End If
End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
  End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_Load()

Call carga_proveedores(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0

Call carga_productos(c_prod)
c_prod.ListIndex = 0

Call carga_tipocomp(c_tipocomp)
c_tipocomp.AddItem "<Todos>", 0
c_tipocomp.ListIndex = 0

Call armagrid
Call barraesag(Me)


End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [ENTER] Visualiza Comprobante - [F6] Exporta "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

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
    c(6) = 6
    c(7) = 7
    c(8) = 8
    For i = 9 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "PRODUCTOS INGRESADOS", "", "", "", 72, 8, True, False)
  End If

End If


If KeyCode = vbKeyF6 Then
  Dim c2(15) As Double
    c2(0) = 0
    c2(1) = 1
    c2(2) = 2
    c2(3) = 3
    c2(4) = 4
    c2(5) = 5
    c2(6) = 6
    c2(7) = 7
    c2(8) = 8
    For i = 9 To 14
      c2(i) = -1
    Next i
    Call exportagrid(msf1, c2(), "PRODUCTOS INGRESADOS", "", "", "", True, False, para.archivo_exportacion)

End If

End Sub


Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    cc_detalle.T_IDPROV = msf1.TextMatrix(msf1.Row, 1)
    cc_detalle.t_prov = msf1.TextMatrix(msf1.Row, 2)
    cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 5), 3, 4)
    cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 5), 1, 1)
    cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 5), 8, 8)
    cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 3)
    cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 7)
    cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub
