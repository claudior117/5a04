VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_stockcli 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos por CLIENTE y PRODUCTOS"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8565
   ScaleWidth      =   12135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Producto"
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   720
      Width           =   9375
      Begin VB.ComboBox c_cli 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   240
         Width           =   7455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   6480
      TabIndex        =   18
      Top             =   1320
      Width           =   3135
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "vta057.frx":0000
         Left            =   240
         List            =   "vta057.frx":000D
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tools"
      Height          =   615
      Left            =   5760
      TabIndex        =   16
      Top             =   7440
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Ajuste Stock"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   7440
      Width           =   3255
      Begin VB.TextBox t_stock 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Stock rapido "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   8
      Top             =   7080
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta057.frx":002D
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta057.frx":08AF
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Producto"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   9375
      Begin VB.TextBox t_prod 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox t_id 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
      Begin VB.TextBox t_fecha 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   8205
      Width           =   12135
      _ExtentX        =   21405
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
            TextSave        =   "17/11/2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:47"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8916
      _Version        =   393216
      BackColorBkg    =   14737632
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
Attribute VB_Name = "vta_stockcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub carga()
 Call armagrid
 Set cl_prod = New productos
 cl_prod.cargar (Val(t_id))
 t_stock = cl_prod.STOCK
 If t_fecha <> "" Then
   s = cl_prod.stock_anterior_por_cliente(Val(t_id), t_fecha, c_cli.ItemData(c_cli.ListIndex))
   F = Format$(t_fecha, "dd/mm/yyyy")
 Else
   s = 0
   F = "          "
 End If
 c = "S.I. " & " 0000-00000000"
 de = Format$("Saldo Anterior", "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
 ct = Format$(s, "######0.00")
 If s >= 0 Then
    d = ct
    h = Format$(0, "######0.00")
 Else
    h = ct
    d = Format$(0, "######0.00")
 End If
 msf1.AddItem F & Chr$(9) & c & Chr$(9) & de & Chr$(9) & d & Chr$(9) & h & Chr$(9) & ct
 
 Set rs = New ADODB.Recordset
 q = "select * from stk_01 where [id_producto] = " & Val(t_id) & " and [id_cliente]= " & c_cli.ItemData(c_cli.ListIndex)
 If t_fecha <> "" Then
   q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
 End If
 If c_tipo.ListIndex > 0 Then
   q = q & " and [ubicacion] = '" & Mid$(c_tipo, 1, 1) & "'"
 End If
 q = q & " order by [fecha], [num_mov_stk]"
 rs.Open q, cn1
 saldo = s
 While Not rs.EOF
   s = rs("cantidad")
   F = Format$(rs("fecha"), "dd/mm/yyyy")
   c = Format$(Left$(rs("comprobante"), 20), "@@@@@@@@@@@@@@@@@@@@")
  de = rs("descripcion")
  ct = Format$(rs("CANTIDAD"), "######0.00")
  If rs("ubicacion") = "E" Then
    d = ct
    h = Format$(0, "######0.00")
  Else
    h = ct
    d = Format$(0, "######0.00")
  End If
  saldo = saldo + Val(d) - Val(h)
  ct = Format$(saldo, "######0.00")
  msf1.AddItem F & Chr$(9) & c & Chr$(9) & de & Chr$(9) & d & Chr$(9) & h & Chr$(9) & ct & Chr$(9) & rs("num_mov_int") & Chr$(9) & rs("modulo")
  rs.MoveNext
 Wend



End Sub



Private Sub btnacepta_Click()
If verifica Then
 Call carga
End If
End Sub
Function verifica() As Boolean
v = True
If Val(t_id) <= 0 Then
  MsgBox ("Producto Incorrecto")
  v = False
End If


If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    MsgBox ("Fechga Incorrecta")
    v = False
  End If
End If
verifica = v
End Function
Private Sub btnsale_Click()
Unload Me

End Sub



Private Sub c_cli_LostFocus()
If c_cli.ListIndex < 0 Then
 c_cli.ListIndex = 0
End If
End Sub

Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If

End Sub




Private Sub Command1_Click()
stk_movint.Show
End Sub

Private Sub Form_GotFocus()
If para.producto_sel > 0 Then
  t_id = para.producto_sel
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  Case Is = 27
        Unload Me
End Select

End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8

msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 2200
msf1.ColWidth(2) = 4700
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 800
msf1.ColWidth(7) = 500

msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Comprobante"
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Entrada"
msf1.TextMatrix(0, 4) = "Salida"
msf1.TextMatrix(0, 5) = "Stock"
msf1.TextMatrix(0, 6) = "Num.Int."
msf1.TextMatrix(0, 7) = "Modulo"

For i = 0 To 2
    msf1.ColAlignment(i) = 1 'izq
Next i
For i = 3 To 5
    msf1.ColAlignment(i) = 9 'der
Next i


End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call armagrid
Check1 = 0
Call carga_clientes(c_cli)
c_cli.ListIndex = 0
c_tipo.ListIndex = 0

End Sub

  






Private Sub msf1_GotFocus()
Me.KeyPreview = False
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F11] Excel - [ENTER] Detalle"

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
    
    For i = 8 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), Space$(40) & "STOCK POR CLIENTE", "     Producto............: (" & t_id & ")  " & t_prod, "     Cliente.............: " & c_cli, "     Fecha desde.........: " & t_fecha, 80, 8, True, False, "V")
  End If
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If



End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Select Case msf1.TextMatrix(msf1.Row, 7)
    Case Is = "V"
        Load vta_cc_detalle
        vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 6)
        vta_cc_detalle.Show
    Case Is = "C"
        Load cc_detalle
        cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 6)
        cc_detalle.Show
    Case Is = "S"
        Load stk_cc_detalle
        stk_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 6)
        stk_cc_detalle.Show
    End Select
  
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_id_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F8] Lista Precios - [ENTER] Acepta  - [ESC] Sale "

If para.producto_sel > 0 Then
  t_id = para.producto_sel
End If
End Sub

Private Sub t_id_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If
End Sub

Private Sub t_id_LostFocus()
Call barraesag(Me)
If Val(t_id) > 0 Then
   Set rs = New ADODB.Recordset
   q = "select [descripcion] from a2 where [id_producto] = " & Val(t_id)
   rs.MaxRecords = 1
   rs.Open q, cn1
   If Not rs.EOF And Not rs.BOF Then
      t_prod = rs("descripcion")
   Else
      t_prod = "Producto Inexistente"
   End If
   Set rs = Nothing
End If
End Sub
