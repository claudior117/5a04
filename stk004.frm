VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form stk_seguirpedidos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEGUIMIENTO DE PEDIDOS POR PRODUCTO"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6600
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   8520
      TabIndex        =   9
      Top             =   7200
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Mostrar"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      Appearance      =   1
      StartOfWeek     =   120127489
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame2 
      Caption         =   "Producto"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3375
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Saldo Anterior (Composicion)"
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   11655
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4680
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   635
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
            TextSave        =   "24/01/07"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "07:40 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "stk_seguirpedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub carga()
 oc = Space$(10)
 ig = Space$(10)
 pi = Space$(10)
 ttoc = 0
 ttin = 0
 ttpi = 0
 
 tpic = 0
 tpir = 0
 tocc = 0
 tocr = 0
 List1.clear
 Call cabecera
 
 'en oc
 Set rs = New ADODB.Recordset
 q = "select * from a5, a6 where [id_tipocomp] = 65 and a5.[num_int] = a6.[num_int] and [id_producto] = " & c_prod.ItemData(c_prod.ListIndex) & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
 rs.Open q, cn1
 toc = 0
 While Not rs.EOF
   toc = toc + (rs("cantidad"))
   tocc = tocc + (rs("cantidad"))
   tocr = tocr + (rs("cantidad_recibida"))
   rs.MoveNext
 Wend
 Set rs = Nothing
 
 'pedidos int
 Set rs = New ADODB.Recordset
 q = "select * from a5, a6 where [id_tipocomp] = 100 and a5.[num_int] = a6.[num_int] and [id_producto] = " & c_prod.ItemData(c_prod.ListIndex) & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
 rs.Open q, cn1
 tpi = 0
 While Not rs.EOF
   tpi = tpi + (rs("cantidad"))
   tpic = tpic + (rs("cantidad"))
   tpir = tpir + (rs("cantidad_recibida"))
   rs.MoveNext
 Wend
 Set rs = Nothing
 
 'ingresados
 Set rs = New ADODB.Recordset
 q = "select * from a5, a6 where [id_tipocomp] = 101 and a5.[num_int] = a6.[num_int] and [id_producto] = " & c_prod.ItemData(c_prod.ListIndex) & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
 rs.Open q, cn1
 tin = 0
 While Not rs.EOF
   tin = tin + rs("cantidad")
   rs.MoveNext
 Wend
 Set rs = Nothing
 
 f = t_fecha
 c = Space$(20)
 de = Format$("Mov. Anteriores", "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 RSet oc = Format$(toc, "######0.00")
 RSet pi = Format$(tpi, "######0.00")
 RSet ig = Format$(tin, "######0.00")
 
  
  List1.AddItem f & "  " & c & "   " & de & " " & pi & "  " & oc & "  " & ig
 
 ttoc = ttoc + toc
 ttin = ttin + tin
 ttpi = ttpi + tpi

 
 
 Set rs = New ADODB.Recordset
 q = "select * from a5, a6, g1 where ([id_tipocomp] = 101 or [id_tipocomp] = 100 or [id_tipocomp] = 65) and a5.[num_int] = a6.[num_int] and [id_producto] = " & c_prod.ItemData(c_prod.ListIndex) & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')" & " and a5.[id_usuario] = g1.[id_usuario]"
 rs.Open q, cn1
 
 
 While Not rs.EOF
   Select Case rs("id_tipocomp")
   Case Is = 65
        toc = (rs("cantidad"))
        tocc = tocc + (rs("cantidad"))
        tocr = tocr + (rs("cantidad_recibida"))
        tin = 0
        tpi = 0
        comp = "O.C.       "
        prov = 1
   Case Is = 100
        toc = 0
        tin = 0
        tpi = (rs("cantidad"))
        tpic = tpic + (rs("cantidad"))
        tpir = tpir + (rs("cantidad_recibida"))
   
        comp = "Pedido Int."
        prov = 0
   
   Case Is = 101
      toc = 0
      tin = rs("cantidad")
      tpi = 0
      comp = "Ingreso    "
      prov = 1
   End Select
   
   
   
     If prov = 1 Then
       Set rs1 = New ADODB.Recordset
       q = "select * from a1 where [id_proveedor] = " & rs("id_proveedor")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
          de = Format$(Left$(rs1("denominacion"), 20), "@@@@@@@@@@@@@@@@@@@@!")
       Else
          de = "Prov. dado de Baja  "
       End If
       Set rs1 = Nothing
     Else
       Set rs1 = New ADODB.Recordset
       q = "select * from a4 where [id_obra] = " & rs("id_proveedor")
       rs1.Open q, cn1
       If Not rs1.EOF And Not rs1.BOF Then
          de = Format$(Left$(rs1("descripcion"), 20), "@@@@@@@@@@@@@@@@@@@@!")
       Else
          de = "Obra dado de Baja   "
       End If
       Set rs1 = Nothing
     End If
        
     RSet oc = Format$(toc, "######0.00")
     RSet pi = Format$(tpi, "######0.00")
     RSet ig = Format$(tin, "######0.00")
 
     ttoc = ttoc + toc
     ttin = ttin + tin
     ttpi = ttpi + tpi
 
     f = Format$(rs("fecha"), "dd/mm/yyyy")
     c = comp
     d = Format$(rs("num_comprobante"), "00000000")
     u = Left$(Format$(rs("usuario"), "@@@@@@@@@@@@@@!"), 14)
     List1.AddItem f & "  " & c & "  " & d & "  " & de & " " & u & " " & pi & "  " & oc & "  " & ig
     
   rs.MoveNext

Wend
RSet oc = Format$(ttoc, "######0.00")
RSet pi = Format$(ttpi, "######0.00")
RSet ig = Format$(ttin, "######0.00")
List1.AddItem ""
List1.AddItem "Total del producto en Pedidos Emitidos: " & pi
List1.AddItem "Total del producto en O.C. Emitidas   : " & oc
List1.AddItem "Total del Producto Ingresados         : " & ig
List1.AddItem ""
RSet pi = Format$(tpic - tpir, "######0.00")
RSet oc = Format$(tocc - tocr, "######0.00")

List1.AddItem "Total del producto en Pedidos sin Cumplir: " & pi
List1.AddItem "Total del producto en O.C. sin Cumplir   : " & oc




End Sub

Sub cabecera()
lc = "-----------------------------------------------------------------------------------------------------------------------"
List1.AddItem "SEGUIMIENTO DE PEDIDOS POR PRODUCTOS"
List1.AddItem ""
List1.AddItem "Producto:   " & Format$(c_prod.ItemData(c_prod.ListIndex), "00000") & " " & c_prod
List1.AddItem ""
List1.AddItem "Fecha Desde:" & t_fecha
List1.AddItem ""
List1.AddItem lc
List1.AddItem "Fecha        Comprobante           Proveedor /Obra      Usuario             Pedidos     En O.C     Ingresados"
List1.AddItem lc


End Sub
Private Sub c_prod_LostFocus()
If c_prod.ListIndex < 0 Then
  c_prod.ListIndex = 0
End If
End Sub

Private Sub cal1_DblClick()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
t_fecha = cal1
cal1.Visible = False
End Sub

Private Sub Command1_Click()
Call carga

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
End If

End Sub

Private Sub Form_Load()
Call barraesag(Me)
t_fecha = Format$(Now, "dd/mm/yyyy")
Call carga_productos(c_prod)
c_prod.ListIndex = 0
cal1.Visible = False

End Sub

  




Private Sub List1_LostFocus()
List1.ListIndex = -1
End Sub


Private Sub t_fecha_DblClick()
cal1.Visible = True
End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
 t_fecha = Format$(Now, "dd/mm/yyyy")
End If
End Sub
