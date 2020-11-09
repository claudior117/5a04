VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_EGRESO2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SALIDAS DE MERCADERIA DE STOCK"
   ClientHeight    =   2175
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   10335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   10335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10095
      Begin VB.TextBox t_detalle2 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   6600
         MaxLength       =   50
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox t_renglonp 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   6720
         MaxLength       =   8
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
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
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Detalle (F6)"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6600
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   4455
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
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
Attribute VB_Name = "stk_EGRESO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub Form_Activate()
t_basico.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2 where"
If tipo = "I" Then
  q = q & " [id_producto] = " & Val(t_basico)
Else
  q = q & " [cod_barra] = '" & RTrim$(t_basico) & "'"
End If
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  t_ip = rs("id_producto")
  t_detalle.Enabled = False
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
End If
Set rs = Nothing
End Sub

Sub carga()
If IsNumeric(t_basico) Then
    If Len(t_basico) <= 5 Then
       Call busca("I") 'busca por id. producto
    Else
       Call busca("B") 'busca por cod. barra
    End If
Else
  Call busca("B") 'busca por cod. barra
End If
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

End Sub




Private Sub t_basico_GotFocus()
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

Sub cargarenglon(t As String)
  
  ip = t_ip
  d = t_detalle
  cu = Format$(Val(t_cantidad), "######0.00")
  If t = "A" Then
    r = stk_egreso.msf1.Rows
    stk_egreso.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & t_detalle2 & Chr(9) & t_tipo
  Else
    r = t_renglon
    stk_egreso.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & t_detalle2 & Chr(9) & t_tipo, r
    stk_egreso.msf1.RemoveItem r + 1
  End If
   
  
End Sub
 
  
Sub limpia()
t_cantidad = ""
t_detalle = ""
t_detalle2 = ""
t_basico = ""
t_renglon = ""
t_tipo = ""
t_renglonp = ""
End Sub


Private Sub t_detalle2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
  t_detalle2 = stk_egreso.t_detalle
End If
End Sub

Private Sub t_detalle2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If Val(t_cantidad) > 0 Then
      If t_renglon = "" Then
       Call cargarenglon("A")
     Else
       Call cargarenglon("M")
     End If
     Call limpia
     'Me.Hide
     t_basico.SetFocus
    End If
 End If
End Sub

Private Sub t_tipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_tipo = Format$(t_tipo, ">@")
  Select Case t_tipo
  Case Is = "S", Is = "E"
     If Val(t_cantidad) > 0 Then
      If t_renglon = "" Then
       Call cargarenglon("A")
     Else
       Call cargarenglon("M")
     End If
     Call limpia
     Me.Hide
     End If
  End Select
End If

End Sub
