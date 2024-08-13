VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_listaprecios_3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "LISTA DE PRECIOS (Formato 3)"
   ClientHeight    =   8595
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12390
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8595
   ScaleWidth      =   12390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tasa IB"
      Height          =   615
      Left            =   8760
      TabIndex        =   42
      Top             =   7440
      Width           =   1215
      Begin VB.ComboBox c_tasaib 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "vta050.frx":0000
         Left            =   120
         List            =   "vta050.frx":0019
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tasa Iva"
      Height          =   615
      Left            =   7200
      TabIndex        =   40
      Top             =   7440
      Width           =   1575
      Begin VB.ComboBox c_tasa 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "vta050.frx":0092
         Left            =   120
         List            =   "vta050.frx":00AB
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Vigentes"
      Height          =   735
      Left            =   7680
      TabIndex        =   35
      Top             =   1680
      Width           =   4575
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducados"
         Height          =   195
         Left            =   2640
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vigentes"
         Height          =   195
         Left            =   1320
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Etiquetas"
      Height          =   735
      Left            =   4560
      TabIndex        =   33
      Top             =   7440
      Width           =   2415
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solo Pendientes de Impresion"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha ultima actualizacion"
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   7440
      Width           =   4335
      Begin VB.TextBox t_fechaf 
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Posterior a"
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Anterior a"
         Height          =   255
         Left            =   120
         TabIndex        =   30
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
      TabIndex        =   27
      Top             =   7320
      Width           =   1695
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   960
         Picture         =   "vta050.frx":0124
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   495
         Left            =   120
         Picture         =   "vta050.frx":09A6
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtros Varios"
      Height          =   735
      Left            =   9600
      TabIndex        =   25
      Top             =   840
      Width           =   2535
      Begin VB.ComboBox c_tipo 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "vta050.frx":1228
         Left            =   120
         List            =   "vta050.frx":1241
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   9720
      TabIndex        =   23
      Top             =   120
      Width           =   1095
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         Picture         =   "vta050.frx":12BA
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   4815
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   12255
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   4695
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   8281
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
      Height          =   1575
      Left            =   7680
      TabIndex        =   18
      Top             =   0
      Width           =   1815
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   1440
         TabIndex        =   46
         Top             =   1080
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.TextBox t_mixto 
         Height          =   285
         Left            =   960
         TabIndex        =   45
         Text            =   "10"
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mixto"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Basico"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Registros"
      Height          =   615
      Left            =   10920
      TabIndex        =   16
      Top             =   120
      Width           =   1215
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   13
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   7215
      Begin VB.TextBox t_codprodprov 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6120
         MaxLength       =   20
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Codigo de producto en la lista del proveedor"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox t_idgrupo 
         Height          =   285
         Left            =   6120
         TabIndex        =   47
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   2040
         Width           =   4575
      End
      Begin VB.ComboBox c_marca 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.ComboBox c_depto 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox c_grupo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox t_detalle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox t_codbarra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_basico 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00800080&
         Caption         =   "C Prod Prov"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800080&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00800080&
         Caption         =   "Marca"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800080&
         Caption         =   "Departamento"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Grupo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Cod. Barra"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
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
      Top             =   8235
      Width           =   12390
      _ExtentX        =   21855
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_listaprecios_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gtipoprecio As Integer
Dim gfila, gcol As Integer
Dim gestilo As Integer
Dim galtofila As Integer
Dim gtamañofuente As Integer
Dim grecargocc3 As Single


Sub carga()
 Call armagrid
 ct = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select [id_producto], a2.[descripcion],  [precio_final], [pu], [reg_faltante], [pedidos], [fecha_actu_precio_venta] from a2, g4 where [cod_tasaiva] = [id_tasaiva]"
 c = " and "
 filtro = 0
 If t_basico <> "" Then
   q = q & c & "[id_producto] = " & Val(t_basico)
   c = " and "
   filtro = 1
 End If
 
 If t_codbarra <> "" Then
   s = " = '"
   q = q & c & "[cod_barra] " & s & RTrim$(t_codbarra) & "'"
   c = " and "
   filtro = 1
 End If
 
 If t_codprodprov <> "" Then
   s = " = '"
   q = q & c & "[id_prod_prov] " & s & RTrim$(t_codprodprov) & "'"
   c = " and "
   filtro = 1
 End If
 
 If t_detalle <> "" Then
   q = q & c & "a2.[descripcion] like  '%" & t_detalle & "%'"
   c = " and "
   filtro = 1
 End If
 
 If c_grupo.ListIndex > 0 Then
   q = q & c & "[id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
   c = " and "
   filtro = 1
 End If
 
 If c_depto.ListIndex > 0 Then
   q = q & c & "[id_departamento] = " & c_depto.ItemData(c_depto.ListIndex)
   c = " and "
   filtro = 1
 End If
 
  If c_marca.ListIndex > 0 Then
   q = q & c & "[id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
   c = " and "
   filtro = 1
 End If
 
  
If c_tasa.ListIndex > 0 Then
   q = q & c & "[cod_tasaiva] = " & c_tasa.ItemData(c_tasa.ListIndex)
   c = " and "
   filtro = 1
 End If
  
 If c_tasaib.ListIndex > 0 Then
   q = q & c & "[id_tasaib] = " & c_tasaib.ItemData(c_tasaib.ListIndex)
   c = " and "
   filtro = 1
 End If
  
  If c_prov.ListIndex > 0 Then
   q = q & c & "[id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex)
   c = " and "
  filtro = 1
 End If
 
  Select Case c_tipo.ListIndex
  
  Case Is = 1, Is = 2
   q = q & c & "[tipo_producto] = '" & Mid$(c_tipo, 1, 1) & "'"
   c = " and "
   filtro = 1
  Case Is = 3
   q = q & c & "[reg_faltante] > 0"
   c = " and "
   filtro = 1
  Case Is = 4
   q = q & c & "[pedidos] > 0"
   c = " and "
   filtro = 1
  Case Is = 5
   q = q & c & "[reg_faltante] > 0 or [pedidos] > 0"
   c = " and "
   filtro = 1
  Case Is = 6
   q = q & c & "[stock] < [stock_minimo]"
   c = " and "
  filtro = 1
  
  Case Else
   
  End Select
 
  If t_fechaf <> "" Then
    q = q & c & " datevalue([fecha_actu_precio_venta]) "
    If Option3 = True Then
       q = q & " <= "
    Else
       q = q & " >= "
    End If
    q = q & " datevalue('" & t_fechaf & "')"
    c = " and "
    filtro = 1
 End If
 
If Check1 = 1 Then
  q = q & c & " [emite_etiqueta] = 'S'"
  c = " and "
  filtro = 1
End If
 
If Option5 = False Then
 If Option6 = True Then
    q = q & c & " [vigente] = true"
    c = " and "
 Else
    q = q & c & " [vigente] = false"
    c = " and "
 End If
 'filtro = 1
End If
 
 
 If Option1 = True Then
   q = q & " order by [id_producto]"
 Else
  If Option2 = True Then
     q = q & " order by a2.[descripcion]"
  Else
     q = q & " order by left$(a2.[descripcion], " & Val(t_mixto) & "), [id_producto]"
     
  End If
 End If

If filtro = 0 Then
  J = MsgBox("Mostrar lista de precios completa? (S/N)", 4)
  If J = 6 Then
    muestra = 1
  Else
    muestra = 0
  End If
Else
 muestra = 1
End If
    

If muestra = 1 Then
espere.Show
espere.Label1 = "Cargando lista de precios...."
espere.Refresh

rs.Open q, cn1
 t_encontrados = 0
 Set cl_stock = New STOCK
 ls = 1
 While Not rs.EOF
    b = Format$(rs("id_producto"), "00000")
    'd = Format$(Left$(rs("a2.descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    d = rs("descripcion")
    
    If gtipoprecio = 0 Then
      p = Format$(rs("precio_final"), "######0.00")
     
    Else
        p = Format$(rs("pu"), "######0.00")
    End If
    p2 = Format$(Val(p) * (1 + (grecargocc3 / 100)), "######0.00")
    
    F = rs("reg_faltante")
    msf1.AddItem b & Chr$(9) & d & Chr$(9) & p & Chr$(9) & p2 & Chr$(9) & rs("fecha_actu_precio_venta") & Chr$(9) & rs("pedidos")
    t_encontrados = Val(t_encontrados) + 1
    
    If gestilo >= 2 Then
     If ls = 0 Then
      Call cambiacolor("&HFFDEC8", msf1.Rows - 1)
      ls = 1
     Else
      Call cambiacolor("&H80000005", msf1.Rows - 1)
      ls = 0
     End If
    End If
    rs.MoveNext
 Wend
 Set cl_stock = Nothing
 Set rs = Nothing
 Unload espere
End If
If msf1.Rows > 1 Then
  msf1.Row = 1
  msf1.col = 2
  msf1.SetFocus
  gfila = 1
  gcol = 2
Else
  gfila = 0
  gcol = 0
End If

End Sub

Sub cambiacolor(ByVal c As String, ByVal F As Integer)
 'color
   msf1.Row = F
   For i = 0 To 9
    msf1.col = i
    msf1.CellBackColor = c
    
   Next i
 End Sub
 
 Sub resalta(ByVal F As Integer)
   msf1.HighLight = flexHighlightNever
   'msf1.FocusRect = flexFocusLight
   If gfila > 0 Then
     msf1.Row = gfila
     msf1.RowHeight(gfila) = galtofila
     For i = 0 To 9
      msf1.col = i
      msf1.CellFontBold = False
      msf1.CellFontSize = gtamañofuente
      msf1.CellForeColor = vbBlack
     Next i
   End If
   msf1.Row = F
   msf1.RowHeight(F) = galtofila + 100
   For i = 0 To 9
    msf1.col = i
    msf1.CellFontBold = True
    msf1.CellFontSize = gtamañofuente + 2
    msf1.CellForeColor = vbRed
   Next i
   msf1.col = gcol
   gfila = F
   msf1.HighLight = flexHighlightWithFocus
 End Sub
 
  
Private Sub btnacepta_Click()

Call carga

End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub

Private Sub c_depto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_grupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_prov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If
End Sub

Private Sub Command2_Click()
On Error GoTo err1
ABM_PROD.Show

Exit Sub
err1:
'Call errormod
End Sub

Private Sub Form_Activate()
para.producto_sel = 0
If para.tipolistaprecios = 1 Then
 If msf1.Rows > 1 Then
  ' msf1.SetFocus
  Else
  t_detalle.SetFocus
  End If
Else
  Select Case para.tipolistaprecios
  Case Is = 2
     vta_listaprecios_2.Show
     Unload Me
  
  Case Else
    If msf1.Rows > 1 Then
      ' msf1.SetFocus
    Else
      t_detalle.SetFocus
    End If
  
  End Select
 End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Me.Hide
End If

End Sub

Private Sub Form_Load()
  Set rs = New ADODB.Recordset
  q = "select * from g1 where [id_usuario] = " & para.id_usuario
  rs.MaxRecords = 1
  rs.Open q, cn1
  gtipoprecio = rs("tipo_precio_lista")
  gestilo = rs("estilo_lista_precios")
  Set rs = Nothing
  
  Call carga_grupos(c_grupo)
  c_grupo.AddItem "<Todos>", 0
  c_grupo.ListIndex = 0
  Call carga_deptos_venta(c_depto)
  c_depto.AddItem "<Todos>", 0
  c_depto.ListIndex = 0
  Call carga_marcas(c_marca)
  c_marca.AddItem "<Todas>", 0
  c_marca.ListIndex = 0
  Call carga_proveedores(c_prov)
  c_prov.AddItem "<Todos>", 0
  c_prov.ListIndex = 0
  Option2 = True
  Load vta_listaprecios2
  Load vta_listaprecios3
  c_tipo.ListIndex = 0
  Call armagrid
  Check1 = 0
  Option6 = True
  Call carga_tasaiva(c_tasa)
  c_tasa.AddItem "<Todas>", 0
  c_tasa.ListIndex = 0
  
  Call carga_tasaib(c_tasaib)
  c_tasaib.AddItem "<Todas>", 0
  c_tasaib.ListIndex = 0
  
 
Set rs = New ADODB.Recordset
q = "select [recargo_cc] from g0 where [sucursal] = 0"
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  grecargocc3 = rs("recargo_cc")
Else
  grecargocc3 = 0
End If
Set rs = Nothing
  
  
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
'If gestilo = 0 Then
'    msf1.FocusRect = flexFocusLight
'Else
    msf1.FocusRect = flexFocusNone
'End If
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 5000
msf1.ColWidth(2) = 2000
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 1500
msf1.ColWidth(5) = 1000


msf1.TextMatrix(0, 0) = "Basico"
msf1.TextMatrix(0, 1) = "Descripcion"
If gtipoprecio = 0 Then
  msf1.TextMatrix(0, 2) = "Final Contado"
  'msf1.TextMatrix(0, 6) = "P.s/Iva"
  msf1.TextMatrix(0, 3) = "Final Ctacte"
Else
  msf1.TextMatrix(0, 2) = "Sin Iva Cdo"
  msf1.TextMatrix(0, 3) = "Sin Iva CtaCte"
End If
msf1.TextMatrix(0, 4) = "Actualizado"
msf1.TextMatrix(0, 5) = "En O.C"

For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 5
  msf1.ColAlignment(i) = 9 'der
Next i

galtofila = msf1.RowHeight(0)
gtamañofuente = 12


End Sub

  




Private Sub Form_Unload(Cancel As Integer)
Unload vta_listaprecios2
Unload vta_listaprecios3
End Sub

Sub muestra2()
 r = msf1.Row
 c = Val(msf1.TextMatrix(r, 0))
 If c > 1 Then
   Set rs2 = New ADODB.Recordset
   q = " select * from a2 where [id_producto] = " & c
   rs2.Open q, cn1
   If Not rs2.BOF And Not rs2.EOF Then
     vta_listaprecios2.t_linea = r
     
     vta_listaprecios2.t_basico = rs2("id_producto")
     vta_listaprecios2.t_codbarra = rs2("cod_barra")
     vta_listaprecios2.t_detalle = rs2("descripcion")
     vta_listaprecios2.c_grupo.ListIndex = buscaindice(vta_listaprecios2.c_grupo, rs2("id_grupo"))
     vta_listaprecios2.c_depto.ListIndex = buscaindice(vta_listaprecios2.c_depto, rs2("id_departamento"))
     vta_listaprecios2.c_marca.ListIndex = buscaindice(vta_listaprecios2.c_marca, rs2("id_marca"))
     vta_listaprecios2.c_prov.ListIndex = buscaindice(vta_listaprecios2.c_prov, rs2("id_proveedor"))
     vta_listaprecios2.c_unidad.ListIndex = buscaindice(vta_listaprecios2!c_unidad, rs2("id_unidad"))
     vta_listaprecios2.t_envase = rs2("envase")
     vta_listaprecios2.t_pu = Format$(rs2("pu"), "###0.00")
     vta_listaprecios2.c_iva.ListIndex = buscaindice(vta_listaprecios2!c_iva, rs2("cod_tasaiva"))
     vta_listaprecios2.t_stockminimo = rs2("stock_minimo")
     vta_listaprecios2.t_utilidad = rs2("porc_utilidad")
     vta_listaprecios2.t_costo = rs2("costoreal")
     vta_listaprecios2.t_fletecompra = rs2("flete_compra")
     vta_listaprecios2.t_dtocompra = rs2("dto_compra")
     vta_listaprecios2.t_dtocompra2 = rs2("dto_compra2")
     vta_listaprecios2.t_final = Format$(rs2("precio_final"), "####0.00")
     vta_listaprecios2.t_tasaimpint = rs2("tasa_imp_interno")
     vta_listaprecios2.t_tipo = rs2("tipo_producto")
     vta_listaprecios2.t_moneda = rs2("moneda")
     vta_listaprecios2.t_impuesto = rs2("impuesto")
     vta_listaprecios2.t_observaciones = rs2("observaciones")
     vta_listaprecios2.t_preciocompra = rs2("precio_ult_compra")
     vta_listaprecios2.t_ultvta = rs2("ultima_venta")
     vta_listaprecios2.t_ultimacompra = rs2("ultima_compra")
     vta_listaprecios2.t_fechaactu = rs2("fecha_actu_precio_venta")
     vta_listaprecios2.t_stock = rs2("stock")
     vta_listaprecios2.t_oc = rs2("pedidos")
     vta_listaprecios2.t_pedidos = rs2("requeridos")
     vta_listaprecios2.t_fechaactuc = rs2("fecha_ult_compra")
     vta_listaprecios2.t_textocentral = rs2("texto_central")
     vta_listaprecios2.t_tipocarga = rs2("tipo_carga_tique")
     vta_listaprecios2.c_tasaib.ListIndex = buscaindice(vta_listaprecios2!c_tasaib, rs2("id_tasaib"))
     vta_listaprecios2.t_idprodprov = rs2("id_prod_prov")
     vta_listaprecios2.t_percibe5329 = rs2("percibe_5329")
     
     If rs2("vigente") = True Then
      vta_listaprecios2.Check1 = 1
     Else
      vta_listaprecios2.Check1 = 0
     End If
     vta_listaprecios2.Show
     If gtipoprecio = 0 Then
       vta_listaprecios2.t_final.SetFocus
     Else
        vta_listaprecios2.t_pu.SetFocus
     End If
   
   End If
   Set rs2 = Nothing
   
    Call vta_listaprecios2.actualiza
 End If
 
End Sub


Private Sub msf1_DblClick()

If para.id_grupo_modulo_actual > 7 Then
     Call muestra2
  End If
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(1) = "[F1] P.F - [F2] Sel. - [F3] A Faltantes -  [F4] Saca - [F5] Grupal - [F6] Op.  - [F7] Imprime - [F10] Imp.x Grupo - [ENTER] Detalle  - [Esc] Cancela"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = vbKeySpace Then
  t_detalle.SetFocus
End If

If KeyCode = vbKeyF1 Then
 If para.id_grupo_modulo_actual > 7 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    precio = InputBox("Ingrese Precio")
    If Val(precio) > 0 Then
      Set rs = New ADODB.Recordset
      q = "select * from a2 where [id_producto] = " & p
      rs.MaxRecords = 1
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rs.BOF And Not rs.EOF Then
        rs("precio_final") = Val(precio)
        rs("fecha_actu_precio_venta") = Format$(Now, "dd/mm/yyyy")
        rs.Update
        msf1.TextMatrix(r, 2) = Format$(Val(precio), "#####0.00")
              End If
      Set rs = Nothing
    End If
    
  End If
 End If
End If


If KeyCode = vbKeyF2 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    para.producto_sel = p
    Me.Hide
  Else
    para.producto_sel = 0
  End If
End If


If KeyCode = vbKeyF3 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    cant = InputBox("Ingrese Cantidad a Ingresar en el registro de Faltantes")
    If Val(cant) > 0 Then
      Set cl_prod = New productos
      Call cl_prod.cargafaltante(p, cant, 0)
      Set cl_prod = Nothing
    End If
    
  End If
End If


If KeyCode = vbKeyF5 Then
  If para.id_grupo_modulo_actual > 7 Then
   vta_listaprecios3.Show
  End If
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
      
    For i = 3 To 14
      c(i) = -1
    Next i
    
    If t_detalle <> "" Then
      t = "Detalle: " & t_detalle
    Else
      t = ""
    End If
    
    If c_grupo.ListIndex > 0 Then
       t1 = "Grupo: " & c_grupo
    End If
    Call imprimegrid(msf1, c(), "LISTA DE PRECIOS", "", t, t1, 80, 8, True, False, "V")
  End If
End If


If KeyCode = vbKeyF10 Then
 If msf1.Rows > 1 Then
  Call imprimegrupo
 End If
End If



End Sub


Sub imprimegrupo()
Dim dp As String
If c_grupo.ListIndex > 0 Then
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
     pagina = 1
     'Call membreteg
     Call cabecera(pagina)
     linea = 10
     k = 1
     b = 0
     While k < msf1.Rows
     If linea < 65 Then
       If b = 0 Then
         Printer.FontBold = True
         b = 1
       Else
         Printer.FontBold = False
         b = 0
       End If
       dp = Left$(RTrim$(msf1.TextMatrix(k, 1)) & String$(80 - Len(RTrim$(msf1.TextMatrix(k, 1))), "."), 70)
       pc = Space$(8)
       pcc = Space$(8)
       RSet pc = Format$(Val(msf1.TextMatrix(k, 2)), "####0.00")
       RSet pcc = Format$(Val(msf1.TextMatrix(k, 3)), "####0.00")
       renglon = "     " & msf1.TextMatrix(k, 0) & "  " & dp & " " & pc & " " & pcc

      Printer.Print renglon
      k = k + 1
      linea = linea + 1
     Else
      linea = 10
      pagina = pagina + 1
      Printer.NewPage
      'Call membrete
      Call cabecera(pagina)
    End If
  Wend
  Printer.EndDoc
  
     
  End If
Else
  MsgBox ("No tiene grupo seleccionado")
End If
End Sub
Sub cabecera(ByVal p As Integer)
 Call imprimeempresa(14)
 Printer.FontName = "Courier New"
 Printer.FontSize = 10
 Printer.FontUnderline = True
 Printer.Print Tab(35); "LISTA DE PRECIOS"
 Printer.FontUnderline = False
 Printer.Print Tab(70); "Hoja Nro:" & p
 Printer.Print Tab(70); "Fecha Imp:" & Format$(Now, "dd/mm/yyyy")
 Printer.Print
 Printer.FontSize = 16
 Printer.FontBold = True
 Printer.Print Tab(5); "(" & Format$(c_grupo.ItemData(c_grupo.ListIndex), "000") & ")  " & c_grupo
 Printer.FontSize = 10
 Printer.FontBold = False
 Printer.Print
 Printer.Print Tab(5); "======================================================================================================"
 Printer.Print Tab(5); "Basico  PRODUCTO                                                                 Ctdo.  Cta.Cte.   "
 Printer.Print Tab(5); "======================================================================================================"

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If para.id_grupo_modulo_actual > 7 Then
  Call muestra2
  End If
Else
 
  
 If KeyAscii <> 27 And KeyAscii <> 38 Then
  vta_listaprecios5.t_texto = Chr$(KeyAscii)
  vta_listaprecios5.Show
 End If
End If


End Sub

Private Sub msf1_LostFocus()
'Call barra(Me)
End Sub

Private Sub msf1_SelChange()
If gestilo = 1 Or gestilo = 3 Then
  r = msf1.Row
  gcol = msf1.col
  Call resalta(r)
End If
End Sub

Private Sub t_basico_GotFocus()
t_basico = ""
End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub t_codbarra_GotFocus()
t_codbarra = ""
End Sub

Private Sub t_codbarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub t_codprodprov_GotFocus()
t_codprodprov = ""
End Sub

Private Sub t_codprodprov_KeyPress(KeyAscii As Integer)
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

Private Sub t_idgrupo_GotFocus()
t_idgrupo = ""
End Sub

Private Sub t_idgrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  c_grupo.ListIndex = buscaindice(c_grupo, Val(t_idgrupo))
  btnacepta.SetFocus
  
 
End If

End Sub

Private Sub t_idgrupo_LostFocus()
c_grupo.ListIndex = buscaindice(c_grupo, Val(t_idgrupo))
End Sub

Private Sub UpDown1_DownClick()
If Val(t_mixto) > 1 Then
  t_mixto = Val(t_mixto) - 1
End If

End Sub

Private Sub UpDown1_UpClick()
If Val(t_mixto) < 25 Then
  t_mixto = Val(t_mixto) + 1
End If
End Sub
