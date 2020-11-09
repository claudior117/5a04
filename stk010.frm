VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_inventario 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INVENTARIO "
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8715
   ScaleWidth      =   12060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   7200
      TabIndex        =   39
      Top             =   7680
      Width           =   2295
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Muestra productos en 0"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   3240
      TabIndex        =   36
      Top             =   7680
      Width           =   3855
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   1560
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H000000FF&
         Caption         =   "Fecha Corte"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   7680
      Width           =   2895
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000A&
         Caption         =   "Usar Stock Instantaneo"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Valorizado"
      Height          =   735
      Left            =   7680
      TabIndex        =   29
      Top             =   1680
      Width           =   4095
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.Final"
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.U."
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Costo "
         Height          =   255
         Left            =   960
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar"
      Height          =   855
      Left            =   10920
      TabIndex        =   27
      Top             =   0
      Width           =   855
      Begin VB.CommandButton btnacepta 
         Height          =   495
         Left            =   120
         Picture         =   "stk010.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
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
         ItemData        =   "stk010.frx":0882
         Left            =   120
         List            =   "stk010.frx":088F
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   9360
      TabIndex        =   23
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   120
         Picture         =   "stk010.frx":08C1
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   5175
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   4815
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8493
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
      Height          =   855
      Left            =   7680
      TabIndex        =   18
      Top             =   840
      Width           =   1455
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Basico"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Registros"
      Height          =   735
      Left            =   7680
      TabIndex        =   16
      Top             =   0
      Width           =   1575
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   13
         TabIndex        =   17
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
      Width           =   7095
      Begin VB.ComboBox c_prov 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   2040
         Width           =   4575
      End
      Begin VB.ComboBox c_marca 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.ComboBox c_depto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox c_grupo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   9
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
      Begin VB.TextBox t_codbarra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5280
         MaxLength       =   13
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_basico 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000C000&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000C000&
         Caption         =   "Marca"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C000&
         Caption         =   "Departamento"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C000&
         Caption         =   "Grupo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C000&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C000&
         Caption         =   "Cod. Barra"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Caption         =   "Basico"
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
      Top             =   8355
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   635
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
Attribute VB_Name = "stk_inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub carga()
 espere.Show
 espere.Label1 = "Calculando Inventario....."
 espere.Refresh
 Call armagrid
 ct = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select * from a2 where [id_producto] > 1 "
 c = " and "
 
 If t_basico <> "" Then
   q = q & c & "[id_producto] = " & Val(t_basico)
   c = " and "
 End If
 
 If t_codbarra <> "" Then
   If Len(t_codbarra) = 13 Then
     s = " = "
   Else
     s = " >= "
   End If
   q = q & c & "[cod_barra] " & s & Val(t_codbarra)
   c = " and "
 End If
 
 If t_detalle <> "" Then
   q = q & c & "a2.[descripcion] like  '%" & t_detalle & "%'"
   c = " and "
 End If
 
 If c_grupo.ListIndex > 0 Then
   q = q & c & "[id_grupo] = " & c_grupo.ItemData(c_grupo.ListIndex)
   c = " and "
 End If
 
 If c_depto.ListIndex > 0 Then
   q = q & c & "[id_depto] = " & c_depto.ItemData(c_depto.ListIndex)
   c = " and "
 End If
 
  If c_marca.ListIndex > 0 Then
   q = q & c & "[id_marca] = " & c_marca.ItemData(c_marca.ListIndex)
   c = " and "
 End If
 
  If c_prov.ListIndex > 0 Then
   q = q & c & "[id_proveedor] = " & c_prov.ItemData(c_prov.ListIndex)
   c = " and "
 End If
 
  If c_tipo.ListIndex > 0 Then
   q = q & c & "[tipo_producto] = '" & Mid$(c_tipo, 1, 1) & "'"
   c = " and "
  End If
 
 If Option1 = True Then
   q = q & " order by [id_producto]"
 Else
   q = q & " order by a2.[descripcion]"
 End If
 rs.Open q, cn1
 t_encontrados = 0
 Set cl_stock = New STOCK
 tot = 0
 can = 0
 While Not rs.EOF
    b = Format$(rs("id_producto"), "00000")
    'd = Format$(Left$(rs("a2.descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
    d = rs("descripcion")
    If Check1 = 0 Then
     If t_fecha <> "" Then
      Call cl_stock.sacastock(rs("id_producto"), t_fecha)
     Else
      Call cl_stock.sacastock(rs("id_producto"))
     End If
     sp = Format$(cl_stock.stock_movimientos, "#######0.00")
    Else
      sp = Format$(rs("stock"), "#######0.00")
    End If
    If Option3 = False Then
      If Option4 = True Then
        p = rs("costoreal")
      Else
        If Option5 = True Then
          p = rs("pu")
        Else
          p = rs("precio_final")
        End If
      End If
      v = Format$(p * Val(sp), "######0.00")
      tot = tot + Val(v)
    Else
      p = 0
      v = ""
    End If
    can = can + Val(sp)
    If Check2 = 0 Then
     If Val(sp) <> 0 Then
        msf1.AddItem b & Chr$(9) & d & Chr$(9) & sp & Chr$(9) & v
        t_encontrados = Val(t_encontrados) + 1
     End If
    Else
        msf1.AddItem b & Chr$(9) & d & Chr$(9) & sp & Chr$(9) & v
        t_encontrados = Val(t_encontrados) + 1
     End If
    rs.MoveNext
 Wend
 If tot > 0 Then
    msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "----------------------" & Chr$(9) & "----------------------"
    msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & Format$(can, "######0.00") & Chr$(9) & Format$(tot, "#######0.00")
 End If
Set cl_stock = Nothing
msf1.SetFocus
Set rs = Nothing
Unload espere

End Sub


Private Sub btnacepta_Click()
J = MsgBox("Este proceso puede demorar y es recomendable salir del sistema en las terminales. Confirma?", 4)
If J = 6 Then
  Call carga
End If
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
ABM_PROD.Show
End Sub

Private Sub Form_Activate()
para.producto_sel = 0
t_detalle.SetFocus
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
  Option3 = True
  
  c_tipo.ListIndex = 0
  Call armagrid
  
 Check1 = 0
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 5
msf1.ColWidth(0) = 900
msf1.ColWidth(1) = 6400
msf1.ColWidth(2) = 1400
msf1.ColWidth(3) = 1400
msf1.ColWidth(4) = 500



msf1.TextMatrix(0, 0) = "Basico"
msf1.TextMatrix(0, 1) = "Descripcion"
If Check1 = 0 Then
 msf1.TextMatrix(0, 2) = "Stock Mov."
Else
 msf1.TextMatrix(0, 2) = "Stock Inst."
End If
msf1.TextMatrix(0, 3) = "Valor"
msf1.TextMatrix(0, 4) = ""


For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 3 To 4
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

  




Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Detalle - [F4] Saca -  [F7] Imprime - [F11] Excel - [Esc] Cancela"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    msf1.RemoveItem r
    t_encontrados = Val(t_encontrados) - 1
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
    For i = 5 To 14
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
    Call imprimegrid(msf1, c(), "INVENTARIO", "", t, t1, 80, 8, True, False, "V")
  End If
End If


If KeyAscii = vbKeyF11 Then
  Call exportaexcel(msf1)

End If
End Sub

Private Sub msf1_LostFocus()
Call barra(Me)
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

Private Sub T_detalle_GotFocus()
t_detalle = ""
End Sub

Private Sub t_detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
 If Not IsDate(t_fecha) Then
    t_fecha = ""
 End If
End If
End Sub
