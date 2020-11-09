VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_oc1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORDENES DE COMPRA"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   11040
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_renglonrf 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   405
      Left            =   9000
      MaxLength       =   8
      TabIndex        =   25
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10815
      Begin VB.ComboBox c_obra 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   3735
      End
      Begin VB.ComboBox c_tasa 
         Height          =   315
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9000
         MaxLength       =   11
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox t_unidad 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6600
         MaxLength       =   5
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox t_pu 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7200
         MaxLength       =   8
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox t_obs 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox t_renglonp 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   9480
         MaxLength       =   8
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   9960
         MaxLength       =   8
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   720
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   9960
         MaxLength       =   8
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_nroreq 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   10080
         MaxLength       =   8
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox t_cantunit 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5760
         MaxLength       =   8
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Obra / destino"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "% Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8040
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8880
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6480
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "P.U."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Observaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Ref."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9960
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad a Pedir"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2880
      Width           =   11040
      _ExtentX        =   19473
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
            TextSave        =   "28/09/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:18"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "abm_oc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub c_obra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If c_obra.ListIndex >= 0 Then
    Call modifrenglon
    Call limpia
    Me.Hide
  End If
End If
End Sub

Private Sub c_obra_LostFocus()
If c_obra.ListIndex < 0 Then
    c_obra.ListIndex = 0
End If

End Sub

Private Sub Form_Activate()
t_basico.SetFocus
End Sub
Sub carga()
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
  nr = Format$(Val(t_nroreq), "00000000")
  o = t_obs
  pu = Format$(Val(t_pu), "######0.00")
  u = RTrim$(t_unidad) & " "
  ti = Format$(Val(c_tasa), "#0.00")
  im = Format$(Val(t_importe), "######0.00")
  
  If t_renglon <> "" Then
     r = Val(t_renglon)
     ABM_OC.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & nr & Chr(9) & o & Chr(9) & cu & Chr(9) & pu & Chr(9) & u & Chr(9) & ti & Chr(9) & im & Chr(9) & c_obra & Chr(9) & c_obra.ItemData(c_obra.ListIndex) & Chr(9) & t_renglonrf, Val(t_renglon)
     ABM_OC.msf1.RemoveItem Val(t_renglon) + 1
  Else
     r = ABM_OC.msf1.Rows
     ABM_OC.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & nr & Chr(9) & o & Chr(9) & cu & Chr(9) & pu & Chr(9) & u & Chr(9) & ti & Chr(9) & im & Chr(9) & c_obra & Chr(9) & c_obra.ItemData(c_obra.ListIndex) & Chr(9) & t_renglonrf
  End If
  Call ABM_OC.sacatotales2
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 8)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
For i = 0 To 9
  c_tasa.AddItem para.tasaiva(i)
Next i
c_tasa.ListIndex = 0

Call carga_obras(c_obra, "E")
c_obra.ListIndex = 0

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
t_pu = ""

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

Private Sub t_cantunit_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)

End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)
End Sub

Private Sub t_pu_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)

End Sub

Private Sub t_pu_LostFocus()
t_importe = Format$(Val(t_pu) * Val(t_cantunit), "#####0.00")
End Sub
