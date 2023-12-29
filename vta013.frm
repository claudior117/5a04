VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_remitos1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "INGRESO DE ARTICULOS"
   ClientHeight    =   2175
   ClientLeft      =   945
   ClientTop       =   6825
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11880
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   11655
      Begin VB.TextBox t_tr 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox t_envase 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   8040
         MaxLength       =   5
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_bultos 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   7320
         MaxLength       =   8
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t_unidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   6480
         MaxLength       =   8
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         MaxLength       =   5
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   960
         MaxLength       =   69
         TabIndex        =   1
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   10560
         MaxLength       =   11
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox c_tasa 
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
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox t_pu 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   5640
         MaxLength       =   8
         TabIndex        =   2
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Bultos"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7320
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6480
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10440
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Tasa Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Pu"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8160
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5520
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   4695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   11880
      _ExtentX        =   20955
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
            TextSave        =   "14/11/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:28 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_remitos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim grecargocc As Single


Private Sub c_tasa_LostFocus()
If c_tasa.ListIndex < 0 Then
  c_tasa.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()

t_basico.SetFocus

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
    Call TabEnter2(Me, 7)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)

For i = 0 To 9
  c_tasa.AddItem para.tasaiva(i)
Next i
c_tasa.ListIndex = 0

Set rs = New ADODB.Recordset
q = "select [recargo_cc] from g0 where [sucursal] = 0"
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  grecargocc = rs("recargo_cc")
Else
  grecargocc = 0
End If
Set rs = Nothing
End Sub


Private Sub t_basico_GotFocus()
t_detalle.Enabled = False
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If
End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga

End If
End Sub

Sub carga()
If t_basico = "" Then
  t_basico = 1
End If

If IsNumeric(t_basico) Then
 If Val(t_basico) <= 1 Then
    t_basico = 1
    t_ip = 1
    t_unidad = "U."
    t_bultos = 1
    t_envase = 1
    c_tasa.ListIndex = buscaindice2(c_tasa, para.tasageneral)
    t_detalle.Enabled = True
    t_detalle.SetFocus
    t_tr = 0
 Else
    If vta_remitos.c_tipocomp.ItemData(vta_remitos.c_tipocomp.ListIndex) = 46 Then
       t_tr.Visible = True
       Call BUSCAREMITIDO
    Else
       t_tr = 0
    End If
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
q = "select * from a2, g5, g12 where a2.[id_unidad] = g5.[id_unidad] and a2.[id_tasaib] = g12.[id_tasaib] "

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
  
  If grecargocc > 0 Then
       r = (Val(t_pu) * grecargocc) / 100
       t_pu = Format(Val(t_pu) + r, "#####0.00")
   End If
  c_tasa.ListIndex = rs("cod_tasaiva")
  t_ip = rs("id_producto")
  t_unidad = rs("unidad")
  t_bultos = 1
  t_envase = rs("envase")
 
 

Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
End If
Set rs = Nothing
End Sub

Private Sub t_cantidad_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F5] Comvierte Unidades x Envase"
End Sub

Private Sub t_cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
  c = InputBox$("Ingrese Cantidad a Convertir (formula:cantidad x envase)")
  If Val(c) > 0 Then
     t_cantidad = Format$(Val(c) * Val(t_envase), "#####0.00")
  End If
End If
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
  Call solonum(KeyAscii, 1)
Else
 If Val(t_envase) <= 0 Then
   t_envase = "1"
 End If
  t_bultos = Val(t_cantidad) / Val(t_envase)
End If
End Sub

Sub cargarenglon(t As String)
  ip = Val(t_ip)
  d = t_detalle
  cu = Format$(Val(t_cantidad), "######0.00")
  ti = Format$(c_tasa, "####0.00")
  u = Left$(t_unidad, 8)
  b = Format$(Val(t_bultos), "#####0")
  If para.tipoprecioventa = 1 Then
    pu = Format$(Val(t_pu) / (1 + Val(c_tasa) / 100), "#####0.000")
    im = Format$(Val(pu) * Val(cu), "#####0.00")
    puf = Format$(Val(t_pu), "#####0.00")
  Else
    pu = Format$(Val(t_pu), "#####0.000")
    im = Format$(Val(pu) * Val(cu), "#####0.00")
    puf = Format$(Val(t_pu) * (1 + Val(c_tasa) / 100), "#####0.00")
 End If
  If t = "A" Then
    r = vta_remitos.msf1.Rows
    If r <= Val(vta_remitos.t_cantlineas) Then
       vta_remitos.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr(9) & pu & Chr(9) & ti & Chr(9) & b & Chr(9) & im & Chr$(9) & t_tr & Chr$(9) & puf
    Else
       MsgBox ("Imposible seguir agregando productos al documento. Supera el maximo de renglones permitidos")
    End If
  Else
    r = t_renglon
    vta_remitos.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr(9) & pu & Chr(9) & ti & Chr(9) & b & Chr(9) & im & Chr$(9) & t_tr & Chr$(9) & puf, r
    vta_remitos.msf1.RemoveItem r + 1
  End If
  vta_remitos.CALCULATOTALES
  para.producto_sel = 0
End Sub
 
  
Sub limpia()
t_cantidad = ""
t_basico = ""
t_detalle = ""
t_pu = ""
t_importe = ""
t_ip = ""
t_envase = 1
t_bultos = ""
t_unidad = ""
t_renglon = ""
t_tr = ""
t_tr.Visible = False
End Sub

Private Sub t_cantidad_LostFocus()
If vta_remitos.c_tipocomp.ItemData(vta_remitos.c_tipocomp.ListIndex) = 46 Then
 If Val(t_basico) > 1 Then
  If Val(t_cantidad) > Val(t_tr) Then
    MsgBox ("El cliente tiene " & (t_tr) & " unidades del producto seleccionado sin facturar. La cantidad a devolver no puede superar esa cantidad")
    't_cantidad.SetFocus
  End If
 Else
  tr = ""
 End If
End If

End Sub

Private Sub t_importe_GotFocus()
t_importe = Format$(Val(t_cantidad) * Val(t_pu), "#####0.00")
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If vta_remitos.c_tipocomp.ItemData(vta_remitos.c_tipocomp.ListIndex) = 46 Then
   If Val(t_cantidad) > Val(t_tr) Then
     MsgBox ("El cliente tiene " & t_tr & " unidades del producto seleccionado sin facturar. La cantidad a devolver no puede superar esa cantidad")
     t_cantidad.SetFocus
    Else
     Call pasa
   End If
 Else
   Call pasa
 End If
Else
  Call solonum(KeyAscii, 1)
End If
End Sub
Sub BUSCAREMITIDO()
q = "SELECT * FROM VTA_02, VTA_03 WHERE VTA_02.[NUM_INT] = VTA_03.[NUM_INT] AND [ID_TIPOCOMP] = 45 AND [ID_PRODUCTO] = " & Val(t_basico) & " AND [CANTIDAD] > 0 and [id_cliente] = " & vta_remitos.c_prov.ItemData(vta_remitos.c_prov.ListIndex) & " and [estado] = 'S'"
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
c = 0
While Not rs2.EOF
  c = c + rs2("CANTIDAD")
  rs2.MoveNext
Wend
Set rs2 = Nothing
t_tr = c
End Sub
Sub pasa()
  If t_renglon = "" Then
   Call cargarenglon("A")
   t_basico.SetFocus
   
  Else
   Call cargarenglon("M")
   Me.Hide
  End If
  Call limpia
  
End Sub

