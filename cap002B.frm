VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form com_faltantes1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO DE ARTICULOS AL REGISTRO DE FALTANTES"
   ClientHeight    =   2175
   ClientLeft      =   135
   ClientTop       =   4815
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11910
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   11775
      Begin VB.TextBox t_costo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         Picture         =   "cap002B.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1680
         MaxLength       =   49
         TabIndex        =   1
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   720
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox c_prov 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7320
         TabIndex        =   12
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   4815
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "21/06/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:34 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "com_faltantes1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984




Private Sub c_prov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If c_prov.ListIndex < 0 Then
     c_prov.ListIndex = 0
   End If
   Set cl_prod = New productos
   If Val(t_renglon) = 0 Then
     Call cl_prod.cargafaltante(Val(t_basico), Val(t_cantidad), c_prov.ItemData(c_prov.ListIndex))
   Else
     Call cl_prod.modificafaltante(Val(t_renglon), Val(t_cantidad), c_prov.ItemData(c_prov.ListIndex))
   End If
   Set cl_prod = Nothing
   Call com_faltantes.carga
   Me.Hide
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

Private Sub Command1_Click()
ABM_PROD.Show
End Sub

Private Sub Form_Activate()
If Val(t_renglon) = 0 Then
  t_basico.Enabled = True
  t_detalle.Enabled = True
  t_basico.SetFocus
Else
  t_basico.Enabled = False
  t_detalle.Enabled = False
  t_cantidad.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_proveedores(c_prov)
c_prov.ListIndex = 0

End Sub


Private Sub t_basico_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Acepta - [ESC] Sale - [F6] Dto1 - [F7] Dto2  "

t_detalle.Enabled = False
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If

If KeyCode = vbKeyF6 And t_renglon = "" Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  d1 = rs("descuento1")
  Set rs = Nothing
  t_basico = 1
  t_detalle = "Descuento " & Format$(d1, "##0.00") & "%"
  t_pu = "0.00"
  c_tasa.ListIndex = 1
  t_ip = 1
  t_cantidad = "1.00"
  t_pu = Format$(-Val(vta_facturacion.t_subtotal) * d1 / 100, "######0.00")
  t_importe = t_pu
  
End If

If KeyCode = vbKeyF7 And t_renglon = "" Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  d2 = rs("descuento2")
  Set rs = Nothing
  t_basico = 1
  t_detalle = "Descuento " & Format$(d2, "##0.00") & "%"
  t_pu = "0.00"
  c_tasa.ListIndex = 1
  t_ip = 1
  t_cantidad = "1.00"
  t_pu = Format$(-Val(vta_facturacion.t_subtotal) * d2 / 100, "######0.00")
  t_importe = t_pu

End If

End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Sub carga()
If IsNumeric(t_basico) Then
 If Val(t_basico) <= 1 Then
    t_basico = 1
    t_ip = 1
    t_detalle.Enabled = True
    t_detalle.SetFocus
 Else
    If Len(t_basico) <= 5 Then
       Call busca("I") 'busca por id. producto
    Else
       Call busca("B") 'busca por cod. barra
    End If
 End If
Else
  Call busca("B") 'busca por cod. barra
End If
End Sub
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2, g5 where a2.[id_unidad] = g5.[id_unidad] "
If tipo = "I" Then
  q = q & " and [id_producto] = " & Val(t_basico)
Else
  q = q & " and [cod_barra] = '" & RTrim$(t_basico) & "'"
End If
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  If para.tipoprecioventa = 1 Then
    t_pu = rs("precio_final")
  Else
    t_pu = rs("pu")
  End If
  t_ip = rs("id_producto")
  t_unidad = rs("unidad")
  t_costo = rs("costoreal")
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
  t_costo = 0
End If
Set rs = Nothing
End Sub

Private Sub t_basico_LostFocus()
Call barraesag(Me)
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)
End Sub


 
  
Sub limpia()
t_cantidad = ""
t_basico = ""
t_detalle = ""
t_pu = ""
t_importe = ""
t_ip = ""
t_unidad = ""
End Sub


