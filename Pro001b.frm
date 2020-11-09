VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form PROD_SOLMAT1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOLICITUD DE MATERIALES"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11160
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   10815
      Begin VB.TextBox t_unidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   8880
         MaxLength       =   5
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox t_fechaesperado 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   9600
         MaxLength       =   10
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox t_obs 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   5520
         MaxLength       =   22
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox t_renglonp 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   120
         MaxLength       =   8
         TabIndex        =   0
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   4335
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
      Begin VB.TextBox t_cantunit 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   7920
         MaxLength       =   8
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8880
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha Esperado"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9600
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Observaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad a Pedir"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7800
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   4455
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11465
            MinWidth        =   11465
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/10/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "9:32"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "PROD_SOLMAT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub Form_Activate()
t_fechaesperado = prod_solmat.t_fechaprob
t_basico.SetFocus

End Sub
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs2 = New ADODB.Recordset
q = "select * from a2 where"
If tipo = "I" Then
  q = q & " [id_producto] = " & Val(t_basico)
Else
  q = q & " [cod_barra] = '" & RTrim$(t_basico) & "'"
End If
rs2.MaxRecords = 1
rs2.Open q, cn1
If Not rs2.BOF And Not rs2.EOF Then
  Set cl_prod = New productos
  cl_prod.cargar (rs2("id_producto"))
  If cl_prod.idproducto > 0 Then
    t_detalle = cl_prod.Detalle
    t_ip = cl_prod.idproducto
    t_unidad = cl_prod.unidad
    t_detalle.Enabled = False
  Else
    MsgBox ("Error al cargar el producto")
    t_basico.SetFocus
  End If
  Set cl_prod = Nothing
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
End If
Set rs2 = Nothing
End Sub

Sub carga()
If t_basico <> "" Then
 If IsNumeric(t_basico) Then
  If Val(t_basico) <= 1 Then
    t_basico = 1
    t_ip = 1
    t_detalle.Enabled = True
    t_detalle.SetFocus
  Else
    If Len(t_basico) <= 5 Then
       Call busca("I") 'busca por id. producto
    Else
       Call busca("B") 'busca por cod. barra
    End If
  End If
 Else
  Call busca("B") 'busca por cod. barra
 End If
Else
  t_basico = 1
  t_ip = 1
  t_detalle.Enabled = True
  t_detalle.SetFocus
 
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub

Sub modifrenglon()
If Val(t_cantunit) >= 0 Then
  ip = Val(t_basico)
  d = t_detalle
  cu = Format$(Val(t_cantunit), "######0.00")
  nr = Format$(t_fechaesperado, "dd/mm/yyyy")
  o = t_obs
  u = t_unidad
  If t_renglon <> "" Then
     r = Val(t_renglon)
     prod_solmat.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & nr & Chr(9) & o & Chr(9) & cu & Chr(9) & u & Chr(9) & pu, Val(t_renglon)
     prod_solmat.msf1.RemoveItem Val(t_renglon) + 1
  Else
     r = prod_solmat.msf1.Rows
     prod_solmat.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & nr & Chr(9) & o & Chr(9) & cu & Chr(9) & u & Chr(9) & pu
  End If
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 5)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)

End Sub

  
Sub limpia()
t_renglon = ""
t_cantunit = ""
t_detalle = ""
t_basico = ""
t_ip = ""
t_nroreq = ""
t_renglonp = ""
t_obs = ""
t_fechaesperado = prod_solmat.t_fechaprob
End Sub

Private Sub t_basico_GotFocus()
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
Me.StatusBar1.Panels.Item(2) = "[F8] Lista de Productos - [1] Productos sin Codificar - [ESC] Regresa "

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

Private Sub t_basico_LostFocus()
Call barra(Me)
End Sub

Private Sub t_cantunit_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)

End Sub

Private Sub t_fechaesperado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If t_fechaesperado = "" Then
   t_fechaesperado = Format$(Now, "dd/mm/yyyy")
 End If
 
 If Not IsDate(t_fechaesperado) Then
    t_fechaesperado = Format$(Now, "dd/mm/yyyy")
 End If
 Call modifrenglon
 Call limpia
 Me.Hide
End If

End Sub

