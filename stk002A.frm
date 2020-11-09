VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_pedido2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEDIDO DE MATERIA PRIMA y PRODUCTOS"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   11655
      Begin VB.TextBox t_ref 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox t_obs 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox t_fechae 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   3720
         MaxLength       =   8
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox t_basico 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Observaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   13
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha Esperado"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
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
         Left            =   10560
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   4455
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
            TextSave        =   "11:26 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "stk_pedido2"
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
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
Me.StatusBar1.Panels.Item(2) = "[ENTER] Avanza - [ESC] Sale - [F8] Lista Precios"

End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  vta_listaprecios.Show
End If

End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If

End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  J = MsgBox("Confirma Ingreso", 4)
  If J = 6 Then
    Call graba
    Call limpia
    Me.Hide
  End If
Else
  Call solonum(KeyAscii, 1)
End If
End Sub
Sub graba()
      cn1.BeginTrans
      If Val(t_ref) > 0 Then
        'modifica el item
         QUERY = "update a11 set  [id_producto]=" & Val(t_basico) & " , [detalle]='" & t_detalle & "' , [cantidad]= " & Val(t_cantidad) & " , [observaciones]='" & t_obs & "' , [fecha_pedido]='" & Format$(Now, "dd/mm/yyyy") & "' , [fecha_esperado]='" & t_fechae & "' , [cantidad_ingresada]="
         QUERY = QUERY & " where [num_int_item]= " & Val(t_ref)
         cn1.Execute QUERY
         Call cargarenglon("M")
      Else
        Set rs = New ADODB.Recordset
        q = "select * from g0 where [sucursal] = " & glo.sucursal
        rs.Open q, cn1, adOpenStatic, adLockOptimistic
        If Not rs.EOF And Not rs.BOF Then
          p = rs("ult_num_ref") + 1
          rs("ult_num_ref") = p
          rs.Update

          'agrego item
          QUERY = "INSERT INTO a11([num_int_item], [id_producto], [detalle], [cantidad], [id_obra], [observaciones], [estado], [cantidad_oc], [cantidad_ingresada], [fecha_pedido], [fecha_esperado], [id_usuario], [num_oc])"
          QUERY = QUERY & " VALUES (" & p & ", " & Val(t_basico) & ", '" & t_detalle & "', " & Val(t_cantidad) & ", " & stk_pedidos.c_prov.ItemData(stk_pedidos.c_prov.ListIndex) & ", '" & t_obs & "', 'P', 0, 0, '" & Format$(Now, "dd/mm/yyyy") & "', '" & t_fechae & "', " & para.id_usuario & ", 0)"
          cn1.Execute QUERY
          t_ref = p
          Call cargarenglon("A")
        End If
       End If
      
      cn1.CommitTrans
      Set rs = Nothing
      
     
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub

Sub cargarenglon(t As String)
  ip = t_ip
  d = t_detalle
  cu = Format$(Val(t_cantidad), "######0.00")
  If t = "A" Then
    r = stk_pedidos.msf1.Rows
    stk_pedidos.msf1.AddItem r & Chr(9) & Val(t_ref) & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & t_obs & Chr(9) & Format$(Now, "dd/mm/yyyy") & Chr(9) & t_fechae & Chr(9) & para.id_usuario & Chr(9) & para.usuario & Chr(9) & "P" & Chr(9) & "00000000"
  Else
    r = Val(t_renglon)
    stk_pedidos.msf1.AddItem r & Chr(9) & Val(t_ref) & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & t_obs & Chr(9) & Format$(Now, "dd/mm/yyyy") & Chr(9) & t_fechae & Chr(9) & para.id_usuario & Chr(9) & para.usuario & Chr(9) & "P" & Chr(9) & "00000000", r
    stk_pedidos.msf1.RemoveItem r + 1
  End If
   
  
End Sub
 
  
Sub limpia()
t_cantidad = 0
t_detalle = ""
t_basico = ""
t_renglon = ""
t_ip = ""
t_fechae = ""
t_obs = ""
t_ref = ""
End Sub

Private Sub T_detalle_GotFocus()
Call barraesag(Me)
End Sub


Private Sub t_fechae_LostFocus()
If t_fechae <> "" Then
  If Not IsDate(t_fechae) Then
    t_fechae = Format$(Now, "dd/mm/yyyy")
  Else
     t_fechae = Format$(t_fechae, "dd/mm/yyyy")
  End If
Else
  t_fechae = Format$(Now, "dd/mm/yyyy")
End If
  
End Sub
