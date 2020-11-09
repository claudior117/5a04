VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form exp_productos 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "BUSCADOR DE PRODUCTOS"
   ClientHeight    =   8640
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12480
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8640
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Agregar Observaciones"
      Height          =   735
      Left            =   240
      TabIndex        =   27
      Top             =   7440
      Width           =   9135
      Begin VB.TextBox t_obs 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Filtros Varios"
      Height          =   1095
      Left            =   7680
      TabIndex        =   24
      Top             =   1200
      Width           =   4215
      Begin VB.ComboBox c_cod 
         Height          =   315
         ItemData        =   "exp004.frx":0000
         Left            =   840
         List            =   "exp004.frx":000D
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "exp004.frx":0053
         Left            =   840
         List            =   "exp004.frx":0060
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackColor       =   &H00800080&
         Caption         =   "Tipo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Compra"
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   6135
      Begin VB.TextBox t_fecha2 
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800080&
         Caption         =   "Hasta"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Desde"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar"
      Height          =   855
      Left            =   10080
      TabIndex        =   14
      Top             =   7320
      Width           =   1695
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   960
         Picture         =   "exp004.frx":0090
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   495
         Left            =   120
         Picture         =   "exp004.frx":0912
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   4815
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   4695
         Left            =   0
         TabIndex        =   13
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
      Height          =   1095
      Left            =   9480
      TabIndex        =   9
      Top             =   0
      Width           =   2175
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Compra"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Basico"
         Height          =   255
         Left            =   120
         TabIndex        =   10
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
      TabIndex        =   7
      Top             =   0
      Width           =   1575
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   13
         TabIndex        =   8
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
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   960
         Width           =   5175
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
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
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
Attribute VB_Name = "exp_productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub carga()
espere.Show
espere.Label1 = "Cargando lista de productos disponibles...."
espere.Refresh
 Call armagrid
 ct = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select * from a5, a6, a1 where a5.[num_int] = a6.[num_int] and a5.[id_proveedor] = a1.[id_proveedor] and [id_tipocomp] = 1 "
 c = " and "
 
 If t_basico <> "" Then
   q = q & c & "[id_producto] = " & Val(t_basico)
   c = " and "
 End If
 
 
 If t_detalle <> "" Then
   q = q & c & "[detalle] like  '%" & t_detalle & "%'"
   c = " and "
 End If
 
 If c_prod.ListIndex > 0 Then
   q = q & c & "[id_producto] = " & c_prod.ItemData(c_prod.ListIndex)
   c = " and "
 End If
 
  
 If t_fecha <> "" Then
    q = q & c & " datevalue(a5.[fecha]) >= datevalue('" & t_fecha & "')"
    c = " and "
 End If
 
 If t_fecha2 <> "" Then
    q = q & c & " datevalue(a5.[fecha]) <= datevalue('" & t_fecha2 & "')"
    c = " and "
 End If
 
 If c_tipo.ListIndex = 0 Then
   q = q & c & " [cantidad] > [exportacion] "
 Else
   If c_tipo.ListIndex = 1 Then
     q = q & c & " [exportacion] > 0 "
   End If
 End If
 
 If c_cod.ListIndex = 1 Then
   q = q & c & " [id_producto] > 1 "
 Else
   If c_tipo.ListIndex = 2 Then
     q = q & c & " [id_producto] = 1 "
   End If
 End If
 
 
 If Option1 = True Then
   q = q & " order by a6.[num_int]"
 Else
  If Option2 = True Then
     q = q & " order by [detalle]"
  Else
    q = q & " order by a5.[fecha]"
  End If
 End If

rs.Open q, cn1
t_encontrados = 0
While Not rs.EOF
  
    b = Format$(rs("id_producto"), "00000")
    d = rs("detalle")
    p = Format$(rs("importe"), "######0.00")
    fp = rs("a5.fecha")
    CD = rs("cantidad") - rs("exportacion")
    cu = rs("exportacion")
    p = rs("denominacion")
    ope = "FC. " & rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
    msf1.AddItem "" & Chr$(9) & b & Chr$(9) & d & Chr$(9) & CD & Chr$(9) & cu & Chr$(9) & rs("unidad") & Chr$(9) & rs("a5.fecha") & Chr$(9) & rs("pu") & Chr$(9) & p & Chr$(9) & ope & Chr$(9) & rs("a6.num_int") & Chr$(9) & rs("renglon") & Chr$(9) & rs("a5.id_proveedor")
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



Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
   c_tipo.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()

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
  Call armagrid
  Check1 = 0
  Option6 = True
  Load exp_productos1
  c_tipo.ListIndex = 0
  c_cod.ListIndex = 0
  t_obs = " "
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 13
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 3000
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 1000
msf1.ColWidth(5) = 700
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 1000
msf1.ColWidth(8) = 2000
msf1.ColWidth(9) = 1500
msf1.ColWidth(10) = 0
msf1.ColWidth(11) = 0
msf1.ColWidth(12) = 0


msf1.TextMatrix(0, 0) = "A Imputar"
msf1.TextMatrix(0, 1) = "Basico"
msf1.TextMatrix(0, 2) = "Producto"
msf1.TextMatrix(0, 3) = "Disponible"
msf1.TextMatrix(0, 4) = "Utilizados"
msf1.TextMatrix(0, 5) = "Unidad"
msf1.TextMatrix(0, 6) = "Fecha Compra"
msf1.TextMatrix(0, 7) = "PU s/iva"
msf1.TextMatrix(0, 8) = "Proveedor"
msf1.TextMatrix(0, 9) = "Operacion "
msf1.TextMatrix(0, 10) = "Num.Int "
msf1.TextMatrix(0, 11) = "Renglon "
msf1.TextMatrix(0, 12) = "Id. Prov. "



For i = 0 To 1
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 2 To 6
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

  




Private Sub Form_Unload(Cancel As Integer)
Unload exp_productos1
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(1) = "[Barra] Selecciona -  [F9] Graba Reintegro  - [Esc] Cancela"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)




If KeyCode = vbKeyF9 Then
  J = MsgBox("Confirma Agregar productos a Reintegro", 4)
  If J = 6 Then
     'call graba
     
     q = "select * from exp02 where [num_exp] = " & Val(exp_prodreintegro.t_numop)
     Set rs = New ADODB.Recordset
     rs.Open q, cn1, adOpenDynamic, adLockOptimistic
     If Not rs.EOF And Not rs.BOF Then
        rs.MoveLast
        r = rs("renglon")
     Else
        r = 0
     End If
     Set rs = Nothing
     'busco y grabo
     top2 = Val(exp_prodreintegro.t_numop)
     cn1.BeginTrans
     For i = 1 To msf1.Rows - 1
        If Val(msf1.TextMatrix(i, 0)) > 0 Then
           r = r + 1
          QUERY = "INSERT INTO exp02([num_exp], [renglon], [num_int_c], [renglon_c], [cantidad], [id_producto], [producto], [unidad], [pusiva], [obs], [fecha_compra], [operacion_c], [id_proveedor])"
          QUERY = QUERY & " VALUES (" & top2 & ", " & r & ", " & Val(msf1.TextMatrix(i, 10)) & ", " & Val(msf1.TextMatrix(i, 11)) & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & Left$(RTrim$(msf1.TextMatrix(i, 2)), 50) & "', '" & msf1.TextMatrix(i, 5) & " ', " & Val(msf1.TextMatrix(i, 7)) & ", '" & Left$(t_obs, 50) & " ', '" & msf1.TextMatrix(i, 6) & "', '" & msf1.TextMatrix(i, 9) & "', " & Val(msf1.TextMatrix(i, 12)) & ")"
           cn1.Execute QUERY
        
        
          QUERY = "update a6 set  [exportacion]=[exportacion] + " & Val(msf1.TextMatrix(i, 0))
          QUERY = QUERY & " where [num_int]= " & Val(msf1.TextMatrix(i, 10)) & " and [renglon]= " & Val(msf1.TextMatrix(i, 11))
          cn1.Execute QUERY
        
        
        End If
     Next i
     cn1.CommitTrans
     
    exp_prodreintegro.carga
    Me.Hide
     
  End If
End If








End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
     J = InputBox$("Cantidad a Imputar", "Agrega Producto al Reintegro de Exportacion", msf1.TextMatrix(msf1.Row, 3))
     If Val(J) >= 0 And Val(J) <= Val(msf1.TextMatrix(msf1.Row, 3)) Then
          msf1.TextMatrix(msf1.Row, 0) = J
     Else
        If Val(J) = 0 Then
          msf1.TextMatrix(msf1.Row, 0) = ""
        End If
     End If
Else
  If KeyAscii <> 13 And KeyAscii <> 27 Then
    exp_productos1.t_texto = Chr$(KeyAscii)
    exp_productos1.Show
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

Private Sub Text1_Change()

End Sub

Private Sub t_obs_GotFocus()
t_obs = ""
End Sub
