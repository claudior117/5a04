VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form pro_empaque1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO DE ARTICULOS"
   ClientHeight    =   2175
   ClientLeft      =   135
   ClientTop       =   4815
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11910
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   11415
      Begin VB.TextBox t_unidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   9120
         MaxLength       =   5
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   8880
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1200
         MaxLength       =   49
         TabIndex        =   1
         Top             =   720
         Width           =   6735
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   240
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_pu 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   10080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   8160
         MaxLength       =   8
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   7680
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Id.Pieza"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Nro. Bulto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10080
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8040
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle Pieza"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   6975
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   450
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
            TextSave        =   "29/04/2014"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "pro_empaque1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim grecargocc As Single




Private Sub Form_Activate()

't_basico.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 4)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)




End Sub


Private Sub t_basico_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Acepta - [ESC] Sale - [F6] Dto1 - [F7] Dto2 - [F8]Lista Precios  "

t_detalle.Enabled = False
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If

If KeyCode = vbKeyF6 And t_renglon = "" Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  d1 = rs("descuento1")
  Set rs = Nothing
  t_basico = 1
  t_detalle = "Descuento " & Format$(d1, "##0.00") & "%"
  t_pu = "0.00"
  c_tasa.ListIndex = 1
  t_ip = 1
  t_cantidad = "1.00"
  t_pu = Format$(-Val(vta_facturacion.t_subtotal) * d1 / 100, "######0.00")
  t_importe = t_pu
  
End If

If KeyCode = vbKeyF7 And t_renglon = "" Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  d2 = rs("descuento2")
  Set rs = Nothing
  t_basico = 1
  t_detalle = "Descuento " & Format$(d2, "##0.00") & "%"
  t_pu = "0.00"
  c_tasa.ListIndex = 1
  t_ip = 1
  t_cantidad = "1.00"
  t_pu = Format$(-Val(vta_facturacion.t_subtotal) * d2 / 100, "######0.00")
  t_importe = t_pu

End If

End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Sub carga()
If t_basico = "" Then
  t_basico = 1
End If
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
   
  Call busca("B")
End If
End Sub
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2, g5, g12 where a2.[id_unidad] = g5.[id_unidad] and a2.[id_tasaib] = g12.[id_tasaib] "
If tipo = "I" Then
  q = q & " and [id_producto] = " & Val(t_basico)
Else
  q = q & " and [cod_barra] = '" & RTrim$(t_basico) & "'"
End If
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  t_ip = rs("id_producto")
  t_unidad = rs("unidad")
  
  
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
 
End If
Set rs = Nothing
End Sub

Private Sub t_basico_LostFocus()
Call barraesag(Me)
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)
End Sub

Sub cargarenglon(t As String)
  
  ip = Val(t_ip)
  d = t_detalle
  cu = Format$(Val(t_cantidad), "######0.00")
  
  u = RTrim$(t_unidad)
  
  If u = "" Then
    u = " "
  End If
  
  If t = "A" Then
    'nueva linea
    r = pro_empaque.msf1.Rows
    If r <= Val(pro_empaque.t_cantlineas) Then
       pro_empaque.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr$(9) & t_pu
    Else
       MsgBox ("Se ha superado el limite maximo de renglones para este comprobante")
    End If
  
  
  Else
    r = t_renglon
    pro_empaque.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr$(9) & u & Chr$(9) & t_pu, r
    pro_empaque.msf1.RemoveItem r + 1
  End If
   
  
  para.producto_sel = 0
    
  
  
     



End Sub
 
  
Sub limpia()
t_cantidad = ""
t_basico = ""
t_detalle = ""
t_pu = ""

t_ip = ""
t_unidad = ""
End Sub

Private Sub T_detalle_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Acepta - [ESC] Sale - [F3] Descripcion extra   "
End Sub

Private Sub t_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  Form1.Show
End If
End Sub

Private Sub t_detalle_LostFocus()
Me.StatusBar1.Panels.Item(2) = "[ENTER] Acepta - [ESC] Sale "
End Sub



Private Sub t_importe_KeyPress(KeyAscii As Integer)

End Sub

Private Sub t_pu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If t_renglon = "" Then
   Call cargarenglon("A")
   t_basico.SetFocus
   
  Else
   Call cargarenglon("M")
   t_basico.SetFocus
   Me.Hide
  End If
  Call limpia
  
Else
  Call solonum(KeyAscii, 1)
End If
End Sub
