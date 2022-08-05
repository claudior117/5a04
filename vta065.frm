VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_listaprecios_5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "LISTA DE PRECIOS (Formato Tiendal)"
   ClientHeight    =   8595
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8595
   ScaleWidth      =   11880
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
         Height          =   315
         ItemData        =   "vta065.frx":0000
         Left            =   120
         List            =   "vta065.frx":0019
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
         Height          =   315
         ItemData        =   "vta065.frx":0092
         Left            =   120
         List            =   "vta065.frx":00AB
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
      Height          =   615
      Left            =   7680
      TabIndex        =   35
      Top             =   1680
      Width           =   4095
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducados"
         Height          =   195
         Left            =   2640
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vigentes"
         Height          =   195
         Left            =   1320
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
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
      Width           =   2655
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solo Pendientes de Impresion"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha ultima actualizacion"
      Height          =   615
      Left            =   240
      TabIndex        =   29
      Top             =   7440
      Width           =   4335
      Begin VB.TextBox t_fechaf 
         Height          =   285
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Posterior a"
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
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
         Picture         =   "vta065.frx":0124
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
         Picture         =   "vta065.frx":09A6
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
      Left            =   9240
      TabIndex        =   25
      Top             =   840
      Width           =   2535
      Begin VB.ComboBox c_tipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "vta065.frx":1228
         Left            =   120
         List            =   "vta065.frx":1241
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
      Left            =   9360
      TabIndex        =   23
      Top             =   120
      Width           =   1095
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         Picture         =   "vta065.frx":12BA
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
      Height          =   1455
      Left            =   7680
      TabIndex        =   18
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mixto"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Toma solo algunos caracteres de la descripcion para ordenar y despues el basico  "
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Basico"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Registros"
      Height          =   615
      Left            =   10560
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
      Begin VB.TextBox t_idgrupo 
         Height          =   285
         Left            =   5880
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox t_color 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6120
         MaxLength       =   20
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox t_talle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6120
         MaxLength       =   20
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox t_codprodprov 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6120
         MaxLength       =   20
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Codigo de producto en la lista del proveedor"
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox c_prov 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   2040
         Width           =   4575
      End
      Begin VB.ComboBox c_marca 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.ComboBox c_depto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Combo1"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox c_grupo 
         Appearance      =   0  'Flat
         Height          =   315
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
         Width           =   4575
      End
      Begin VB.TextBox t_codbarra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
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
      Begin VB.Label Label10 
         BackColor       =   &H00800080&
         Caption         =   "Color"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         TabIndex        =   50
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800080&
         Caption         =   "Talle"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00800080&
         Caption         =   "C Prod Prov"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         TabIndex        =   46
         Top             =   1800
         Width           =   975
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
         Width           =   855
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
      Width           =   11880
      _ExtentX        =   20955
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
Attribute VB_Name = "vta_listaprecios_5"
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


Sub carga()
 Call armagrid
 ct = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select [id_producto], [descripcion],  [precio_final], [pu], [stock], [moneda], [emite_etiqueta], [reg_faltante], [pedidos], [talle], [color], [medida]  from a2"
 c = " where "
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
 
 If t_talle <> "" Then
   q = q & c & "[talle] like  '%" & t_talle & "%'"
   c = " and "
   filtro = 1
 End If
 
 If t_color <> "" Then
   q = q & c & "[color] like  '%" & t_color & "%'"
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
     contara = InputBox$("Ingrese Cantidad caracteres para generar Indice [1-20]", "Lista Mixta", "10")
     If Val(contara) < 1 Or Val(conatara) > 20 Then
        MsgBox ("Valor Incorrecto. Se tomará por defecto (10)")
        contara = "10"
     End If
     q = q & " order by left$(a2.[descripcion], " & Val(contara) & "), [id_producto]"
     
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
      p2 = Format$(rs("pu"), "######0.00")
    Else
        p = Format$(rs("pu"), "######0.00")
        p2 = Format$(rs("precio_final"), "######0.00")
      
    End If
    'cl_stock.sacastock (rs("id_producto"))
    'c = Format$(cl_stock.stock_movimientos, "#######0.00")
    c = rs("stock")
    
    If rs("emite_etiqueta") = "N" Then
      ee = ""
    Else
      ee = "E"
    End If
    F = rs("reg_faltante")
    msf1.AddItem b & Chr$(9) & d & Chr$(9) & p & Chr$(9) & c & Chr$(9) & rs("talle") & Chr$(9) & rs("color") & Chr$(9) & rs("medida") & Chr$(9) & ee & Chr$(9) & F & Chr$(9) & rs("pedidos")
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

Private Sub c_grupo_LostFocus()
If c_grupo.ListIndex < 0 Then
  c_grupo.ListIndex = 0
End If
t_idgrupo = c_grupo.ItemData(c_grupo.ListIndex)
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
 If msf1.Rows > 1 Then
  ' msf1.SetFocus
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
  
 
  
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
'If gestilo = 0 Then
'    msf1.FocusRect = flexFocusLight
'Else
    msf1.FocusRect = flexFocusNone
'End If
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 5000
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 800
msf1.ColWidth(4) = 800
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 400
msf1.ColWidth(8) = 600



msf1.TextMatrix(0, 0) = "Basico"
msf1.TextMatrix(0, 1) = "Descripcion"
msf1.TextMatrix(0, 2) = "P.Final"
msf1.TextMatrix(0, 3) = "Stock"
msf1.TextMatrix(0, 4) = "Talle"
msf1.TextMatrix(0, 5) = "Color"
msf1.TextMatrix(0, 6) = "Medida"
msf1.TextMatrix(0, 7) = "R.F."
msf1.TextMatrix(0, 8) = "En O.C"

For i = 0 To 6
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 3
  msf1.ColAlignment(i) = 9 'der
Next i
For i = 7 To 8
  msf1.ColAlignment(i) = 9 'der
Next i


galtofila = msf1.RowHeight(0)
gtamañofuente = 8


End Sub

  




Private Sub Form_Unload(Cancel As Integer)
Unload vta_listaprecios2
Unload vta_listaprecios3
End Sub

Sub muestra2()
 r = msf1.Row
 c = Val(msf1.TextMatrix(r, 0))
 If c > 1 Then
   Set rs31 = New ADODB.Recordset
   q = " select * from a2 where [id_producto] = " & c
   
   rs31.Open q, cn1
   If Not rs31.BOF And Not rs31.EOF Then
     vta_listaprecios2.t_linea = r
     
     vta_listaprecios2.t_basico = rs31("id_producto")
     vta_listaprecios2.t_codbarra = rs31("cod_barra")
     vta_listaprecios2.t_detalle = rs31("descripcion")
     vta_listaprecios2.c_grupo.ListIndex = buscaindice(vta_listaprecios2.c_grupo, rs31("id_grupo"))
     vta_listaprecios2.c_depto.ListIndex = buscaindice(vta_listaprecios2.c_depto, rs31("id_departamento"))
     vta_listaprecios2.c_marca.ListIndex = buscaindice(vta_listaprecios2.c_marca, rs31("id_marca"))
     vta_listaprecios2.c_prov.ListIndex = buscaindice(vta_listaprecios2.c_prov, rs31("id_proveedor"))
     vta_listaprecios2.c_unidad.ListIndex = buscaindice(vta_listaprecios2!c_unidad, rs31("id_unidad"))
     vta_listaprecios2.t_envase = rs31("envase")
     vta_listaprecios2.t_pu = Format$(rs31("pu"), "###0.00")
     vta_listaprecios2.c_iva.ListIndex = buscaindice(vta_listaprecios2!c_iva, rs31("cod_tasaiva"))
     vta_listaprecios2.t_stockminimo = rs31("stock_minimo")
     vta_listaprecios2.t_utilidad = rs31("porc_utilidad")
     vta_listaprecios2.t_costo = rs31("costoreal")
     vta_listaprecios2.t_fletecompra = rs31("flete_compra")
     vta_listaprecios2.t_dtocompra = rs31("dto_compra")
     vta_listaprecios2.t_dtocompra2 = rs31("dto_compra2")
     vta_listaprecios2.t_final = Format$(rs31("precio_final"), "####0.00")
     vta_listaprecios2.t_tasaimpint = rs31("tasa_imp_interno")
     vta_listaprecios2.t_tipo = rs31("tipo_producto")
     vta_listaprecios2.t_moneda = rs31("moneda")
     vta_listaprecios2.t_impuesto = rs31("impuesto")
     vta_listaprecios2.t_observaciones = rs31("observaciones")
     vta_listaprecios2.t_preciocompra = rs31("precio_ult_compra")
     vta_listaprecios2.t_ultvta = rs31("ultima_venta")
     vta_listaprecios2.t_ultimacompra = rs31("ultima_compra")
     vta_listaprecios2.t_fechaactu = rs31("fecha_actu_precio_venta")
     vta_listaprecios2.t_stock = rs31("stock")
     vta_listaprecios2.t_oc = rs31("pedidos")
     vta_listaprecios2.t_pedidos = rs31("requeridos")
     vta_listaprecios2.t_fechaactuc = rs31("fecha_ult_compra")
     vta_listaprecios2.t_textocentral = rs31("texto_central")
     vta_listaprecios2.t_tipocarga = rs31("tipo_carga_tique")
     vta_listaprecios2.c_tasaib.ListIndex = buscaindice(vta_listaprecios2!c_tasaib, rs31("id_tasaib"))
     vta_listaprecios2.t_idprodprov = rs31("id_prod_prov")
     vta_listaprecios2.t_cotizultcom = rs31("dolar_ult_compra")
     vta_listaprecios2.t_talle = rs31("talle")
     vta_listaprecios2.t_color = rs31("color")
     vta_listaprecios2.t_medida = rs31("medida")
     
     
     If rs31("vigente") = True Then
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
   Set rs = Nothing
   
 End If
 
End Sub


Private Sub msf1_DblClick()

If para.id_grupo_modulo_actual > 6 Then
     Call muestra2
  End If
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(1) = "[F1]P.F - [F2]Sel. - [F3]A Faltantes - [F4]Saca - [F5]Grupal - [F6]Op.  - [F7]Imprime - [F10]Imp. Etiq. - [F11]Marca Etiq. - [ENTER] Detalle  - [F12] Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = vbKeySpace Then
  t_detalle.SetFocus
End If

If KeyCode = vbKeyF1 Then
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
  vta_listaprecios3.Show
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
  Call imprimiretiquetas
  
End If




If KeyCode = vbKeyF11 Then
  r = msf1.Row
  ee = msf1.TextMatrix(r, 7)
  If ee = "E" Then
    msf1.TextMatrix(r, 7) = ""
    ee = "N"
  Else
    msf1.TextMatrix(r, 7) = "E"
    ee = "S"
  End If
  Set rs = New ADODB.Recordset
  q = "select [emite_etiqueta] from a2 where [id_producto] = " & Val(msf1.TextMatrix(r, 0))
  rs.MaxRecords = 1
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  
  If Not rs.EOF And Not rs.BOF Then
    rs("emite_etiqueta") = ee
    rs.Update
  End If
  Set rs = Nothing
End If



If KeyCode = vbKeyF12 Then
  Call exportaexcel(msf1)
End If


End Sub
Sub imprimiretiquetas()
Dim Report As New etiquetas0
Dim Report1 As New etiquetas1
Dim Report2 As New etiquetas2
Dim Report3 As New etiquetas3

inp = InputBox$("Ingrese tamaño de Etiqueta [0-3]", " ", "1")
'Select Case Val(inp)
' Case Is = 0
'    archivo = "\REP\E\E0.exe"
' Case Is = 1
'    archivo = "\REP\E\E1.exe"
' Case Is = 2
'    archivo = "\REP\E\E2.exe"
' Case Is = 3
'    archivo = "\REP\E\E3.exe"
' Case Else
'    archivo = "\REP\E\E1.exe"
'End Select
    
If inp <> "" Then
  espere.Show
  espere.Refresh
  
 If abrirconexionrep Then
    Set rs = New ADODB.Recordset
    q = "select * from t1"
    rs.Open q, cnrep, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
      rs.Delete
      rs.MoveNext
    Wend
        
    J = 0
    While J < msf1.Rows
     If Val(msf1.TextMatrix(J, 0)) > 1 And msf1.TextMatrix(J, 7) = "E" Then
         Set rs1 = New ADODB.Recordset
         q = "select * from a2 where [id_producto] = " & Val(msf1.TextMatrix(J, 0))
         rs1.Open q, cn1, adOpenStatic, adLockOptimistic
         If Not rs1.EOF And Not rs1.BOF Then
            rs.AddNew
            rs("basico") = rs1("Id_producto")
            rs("descripcion") = Left$(rs1("descripcion"), 49)
            rs("texto_central") = RTrim$(Left$(rs1("TEXTO_CENTRAL"), 20))
            If Len(rs1("cod_barra")) > 2 Then
              rs("cod_barras") = rs1("cod_barra")
            Else
              rs("cod_barras") = rs1("id_producto")
            End If
            rs("precio") = rs1("precio_final")
            rs.Update
                        
            rs1("emite_etiqueta") = "N"
            rs1.Update

         End If
         
      End If
      J = J + 1
     Wend
     Unload espere
     Load reportes
     Set rs1 = Nothing
     Select Case Val(inp)
       Case Is = 0
        Report.DiscardSavedData
        Report.Database.SetDataSource rs
        Report.ReadRecords
        reportes.CRViewer1.ReportSource = etiquetas0 'report
        reportes!CRViewer1.ViewReport
        reportes.Show
       Case Is = 1
        Report1.DiscardSavedData
        Report1.Database.SetDataSource rs
        Report1.ReadRecords
        reportes.CRViewer1.ReportSource = etiquetas1 'Report1
        reportes!CRViewer1.ViewReport
        reportes.Show
       Case Is = 2
        Report2.DiscardSavedData
        Report2.Database.SetDataSource rs
        Report2.ReadRecords
        reportes.CRViewer1.ReportSource = etiquetas2 'Report2
        reportes!CRViewer1.ViewReport
        reportes.Show
       Case Is = 3
        Report3.DiscardSavedData
        Report3.Database.SetDataSource rs
        Report3.ReadRecords
        reportes.CRViewer1.ReportSource = etiquetas3 'Report3
        reportes!CRViewer1.ViewReport
        reportes.Show
     End Select
     Set rs = Nothing
     'k = Shell(App.Path & archivo, 1)

 End If
 cnrep.Close
End If

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If para.id_grupo_modulo_actual > 6 Then
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
