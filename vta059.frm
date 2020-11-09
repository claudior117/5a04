VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_stockcli2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Cliente Acumulado por Producto"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Muestra Stock en Cero"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   2775
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Producto"
      Height          =   1695
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   10455
      Begin VB.TextBox t_fecha2 
         Height          =   405
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox t_fecha 
         Height          =   405
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox c_cli 
         Height          =   360
         Left            =   2160
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   240
         Width           =   7455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   4
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta059.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta059.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
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
Attribute VB_Name = "vta_stockcli2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub carga()
 espere.Show
 espere.Refresh
 
 Call armagrid
 
 q = "select * from a2"
 Set rs2 = New ADODB.Recordset
 rs2.Open q, cn1
 While Not rs2.EOF
   espere.Label1 = "Consultando stock producto --> " & rs2("id_producto")
   espere.Label1.Refresh
   sa = 0
   ent = 0
   sal = 0
   s = 0
   If t_fecha <> "" Then
    'calculo saldo anterior
       Set rs = New ADODB.Recordset
       q = "select * from stk_01 where [id_producto] = " & rs2("Id_producto") & " and [id_cliente]= " & c_cli.ItemData(c_cli.ListIndex)
       q = q & " and datevalue([fecha]) < datevalue('" & t_fecha & "')"
       rs.Open q, cn1
        While Not rs.EOF
            If rs("ubicacion") = "E" Then
                sa = sa + rs("cantidad")
            Else
                sa = sa - rs("cantidad")
            End If
            rs.MoveNext
        Wend
        Set rs = Nothing
  End If
  
  
  Set rs = New ADODB.Recordset
  q = "select * from stk_01 where [id_producto] = " & rs2("Id_producto") & " and [id_cliente]= " & c_cli.ItemData(c_cli.ListIndex)
  If t_fecha <> "" Then
    q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If t_fecha2 <> "" Then
    q = q & " and datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
       
  rs.Open q, cn1
  While Not rs.EOF
     If rs("ubicacion") = "E" Then
        ent = ent + rs("cantidad")
     Else
        sal = sal + rs("cantidad")
     End If
     rs.MoveNext
  Wend
  Set rs = Nothing
  s = sa + ent - sal
  If Option1 = True Then
    m = 1
  Else
    If s <> 0 Then
       m = 1
    Else
      m = 0
    End If
  End If
  
  If m = 1 Then
    msf1.AddItem rs2("id_producto") & Chr$(9) & rs2("descripcion") & Chr$(9) & Format$(sa, "#####0.00") & Chr$(9) & Format$(ent, "#####0.00") & Chr$(9) & Format$(sal, "#####0.00") & Chr$(9) & Format$(s, "#####0.00")
  End If
  rs2.MoveNext
 Wend
 Set rs2 = Nothing
 Unload espere


End Sub



Private Sub btnacepta_Click()

 Call carga

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
msf1.Cols = 6

msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 5000
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 1200

msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Producto"
msf1.TextMatrix(0, 2) = "Saldo Ant."
msf1.TextMatrix(0, 3) = "Entradas"
msf1.TextMatrix(0, 4) = "Salida"
msf1.TextMatrix(0, 5) = "Stock"

For i = 0 To 5
    msf1.ColAlignment(i) = 9 'der
Next i
msf1.ColAlignment(1) = 1 'izq


End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call armagrid
Call carga_clientes(c_cli)
c_cli.ListIndex = 0
Option2 = True

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
    
    
    For i = 6 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), Space$(40) & "STOCK CLIENTE ACUMULADO por PRODUCTO", "     Cliente............: " & c_cli, "     Periodo............: " & t_fecha & " al " & t_fecha2, 80, 8, True, False, "V")
  End If
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If



End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load stk_movint
    vta_stockcli.t_id = msf1.TextMatrix(msf1.Row, 0)
    vta_stockcli.c_cli.ListIndex = buscaindice(vta_stockcli.c_cli, c_cli.ItemData(c_cli.ListIndex))
    vta_stockcli.t_fecha = t_fecha
    vta_stockcli.Show
    
    
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





