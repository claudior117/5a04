VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form con_HISTORICOcompras 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HISTORICO DE COMPRAS-VENTA POR PRODUCTO"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8175
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   240
         Width           =   6135
      End
      Begin VB.TextBox t_prod 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   11
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "con011.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "con011.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   0
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
      TabIndex        =   3
      Top             =   8460
      Width           =   12015
      _ExtentX        =   21193
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
            TextSave        =   "11/03/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:54 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5535
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9763
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   2
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
Attribute VB_Name = "con_HISTORICOcompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub carga()
 espere.Show
 espere.Label1 = "Espere...... Buscando Informacion de Compras"
 espere.Refresh
 Call armagrid
 q = "select * from A2 "
 c = " where "
 p = 1
 If c_prod.ListIndex > 0 Then
  q = q & c & " [id_producto] = " & c_prod.ItemData(c_prod.ListIndex)
  c = " and "
 End If
 
 If t_prod <> "" Then
   q = q & c & " [descripcion] like '%" & t_prod & "%'"
  End If
 
 
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  While Not rs.EOF
   
     msf1.AddItem rs("id_producto") & Chr(9) & rs("descripcion")
     rs.MoveNext
  Wend
  Set rs = Nothing
  Unload espere
     
End Sub


Private Sub btnacepta_Click()
  Call nivel_acceso(2) 'compras
  If para.id_grupo_modulo_actual > 5 Then
    Call carga
  Else
    Call sinpermisos
  End If
End Sub

Private Sub btnsale_Click()
Unload Me
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 2
msf1.ColWidth(0) = 1500
msf1.ColWidth(1) = 5000


msf1.TextMatrix(0, 0) = "Id"
msf1.TextMatrix(0, 1) = "Producto"


End Sub

Private Sub Form_Load()
'Call carga_productos(c_prod)
c_prod.AddItem "<Todos>", 0
c_prod.ListIndex = 0
Call armagrid
Option1 = True

End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Historico"

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load vta_listaprecios4
    vta_listaprecios4.t_idprod = msf1.TextMatrix(msf1.Row, 0)
    vta_listaprecios4.t_prod = msf1.TextMatrix(msf1.Row, 1)
    vta_listaprecios4.t_fecha = t_fecha
    vta_listaprecios4.t_fecha2 = t_fecha2
    If Check1 = 1 Then
      vta_listaprecios4.Option1 = True
    Else
      vta_listaprecios4.Option2 = True
    End If
    vta_listaprecios4.Show
  End If
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

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  End If
End If

End Sub


Private Sub t_prod_GotFocus()
t_prod = ""
End Sub
