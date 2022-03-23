VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_pedidos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORDEN DE PRODUCCION"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   615
      Left            =   360
      TabIndex        =   21
      Top             =   7440
      Width           =   5775
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Obra"
         Height          =   255
         Left            =   3960
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Esperado"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Referencia"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Listar"
      Height          =   615
      Left            =   8040
      Picture         =   "stk002.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1335
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   11295
      Begin VB.TextBox t_oc 
         Height          =   285
         Left            =   9840
         MaxLength       =   10
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox c_estado 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   4455
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   4455
      End
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   4455
      End
      Begin VB.ComboBox c_usuario 
         Height          =   315
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox t_fecha1 
         Height          =   285
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha2 
         Height          =   285
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "O.C."
         Height          =   255
         Left            =   9120
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Estado "
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Producto"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Obra"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario"
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha desde:"
         Height          =   375
         Left            =   6120
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha desde:"
         Height          =   255
         Left            =   6120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9763
      _Version        =   393216
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   2
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "stk002.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "stk002.frx":0B8C
         Style           =   1  'Graphical
         TabIndex        =   3
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
      TabIndex        =   1
      Top             =   8505
      Width           =   12105
      _ExtentX        =   21352
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
            TextSave        =   "11/03/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:51 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "stk_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Dim gnumoc As Long
Sub carga_oc()
    Call busca
     r = 1
     While Not rs.EOF
        i = rs("num_int_item")
        msf1.AddItem r & Chr(9) & i & Chr(9) & Format$(rs("id_producto"), "00000") & Chr(9) & rs("detalle") & Chr(9) & Format$(rs("cantidad"), "######0.00") & Chr(9) & rs("observaciones") & Chr(9) & rs("fecha_pedido") & Chr(9) & rs("fecha_esperado") & Chr(9) & rs("a11.id_usuario") & Chr(9) & rs("usuario") & Chr(9) & rs("a11.estado") & Chr(9) & Format$(rs("num_oc"), "00000000") & Chr(9) & Format$(rs("cantidad_ingresada"), "######0.00")
        rs.MoveNext
        r = r + 1
     Wend
     'Set rs = Nothing
     
     If msf1.Rows >= 2 Then
        Command4.Enabled = True
     End If
  'End If
End Sub
Sub busca()
 Set rs = New ADODB.Recordset
     q = "select * from a11, g1, a4 where  a11.[id_usuario] = g1.[id_usuario] and a11.[id_obra] = a4.[id_obra] "
     If c_prov.ListIndex > 0 Then
         q = q & " and a11.[id_obra] = " & c_prov.ItemData(c_prov.ListIndex)
     End If
     
    If c_estado.ListIndex > 0 Then
         q = q & " and a11.[estado] = '" & Mid$(c_estado, 1, 1) & "'"
    End If

    If c_prod.ListIndex > 0 Then
         q = q & " and [id_producto] = " & c_prod.ItemData(c_prod.ListIndex)
    End If
 
    If c_usuario.ListIndex > 0 Then
         q = q & " and [a11.id_usuario] = " & c_usuario.ListIndex
    End If
     
    If Val(t_oc) > 0 Then
         q = q & " and [num_oc] = " & Val(t_oc)
    End If
     
     
    If Option1 = True Then
       q = q & " order by [num_int_item]"
    Else
      If Option2 = True Then
         q = q & " order by [fecha_esperado]"
      Else
         q = q & " order by A11.[id_obra]"
      End If
    End If

      
     rs.Open q, cn1

End Sub
Private Sub btnacepta_Click()
Call armagrid
Call carga_oc
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 13
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 5000
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 2500
msf1.ColWidth(6) = 1200
msf1.ColWidth(7) = 1200
msf1.ColWidth(8) = 400
msf1.ColWidth(9) = 1000
msf1.ColWidth(10) = 1000
msf1.ColWidth(11) = 1000
msf1.ColWidth(12) = 1000



msf1.TextMatrix(0, 0) = "Renglon"
msf1.TextMatrix(0, 1) = "Referencia"
msf1.TextMatrix(0, 2) = "Id.Prod."
msf1.TextMatrix(0, 3) = "Detalle"
msf1.TextMatrix(0, 4) = "Cantidad"
msf1.TextMatrix(0, 5) = "Obs."
msf1.TextMatrix(0, 6) = "Fecha Ped."
msf1.TextMatrix(0, 7) = "Fecha Sol."
msf1.TextMatrix(0, 8) = "Id."
msf1.TextMatrix(0, 9) = "Usuario"
msf1.TextMatrix(0, 10) = "Estado"
msf1.TextMatrix(0, 11) = "Ult. O.C"
msf1.TextMatrix(0, 12) = "Cant.en O.C"



Command4.Enabled = False
End Sub







Private Sub c_prov_GotFocus()
Call armagrid
End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
      c_prov.ListIndex = 0
End If
 
End Sub

Private Sub Command4_Click()
Call busca

Call ejecutareporte2(rs, stk001)
End Sub

Private Sub Form_Load()
Load cc_detalle

Call carga_obras(c_prov, "E")
c_prov.AddItem "<Todas>", 0
c_prov.ListIndex = 0
t_sucursal = Format$(glo.sucursal, "0000")

'Call carga_productos(c_prod)
c_prod.AddItem "<Todos>", 0
c_prod.ListIndex = 0

Call carga_usuarios(c_usuario)
c_usuario.AddItem "<Todos>", 0
c_usuario.ListIndex = buscaindice(c_usuario, para.id_usuario)

c_estado.AddItem "<Todos>", 0
c_estado.AddItem "Pendientes(O.C sin completar)", 1
c_estado.AddItem "Ordenes de Compra Emitidas", 2
c_estado.ListIndex = 0



Call armagrid
Call barraesag(Me)

Option1 = True
gnumoc = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload cc_detalle
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Ingresa O.C. - [F4] Ver O.C."
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
   If (para.id_usuario = Val(msf1.TextMatrix(msf1.Row, 8))) And msf1.TextMatrix(msf1.Row, 10) = "P" Then
     msf1.RemoveItem (msf1.Row)
   Else
     MsgBox ("Solo el usuario que ingreso el item puede sacarlo")
   End If
 Else
   Call armagrid
 End If
End If

If KeyCode = vbKeyF4 Then
 If msf1.Rows > 2 Then
   If Val(msf1.TextMatrix(msf1.Row, 11)) > 0 Then
       q = "select * from a5 where [id_tipocomp] = 65 and [num_comprobante] = " & Val(msf1.TextMatrix(msf1.Row, 11))
       Set rs = New ADODB.Recordset
       rs.Open q, cn1
       If Not rs.BOF And Not rs.EOF Then
         cc_detalle.t_numint = rs("num_int")
         cc_detalle.Show
       End If
   End If
 End If
End If

If KeyCode = vbKeyF9 Then ' Carga OC
 If msf1.Rows > 2 Then
   If para.id_grupo_modulo_compras >= 7 Then
      'cargar oc
      X = InputBox$("Ingrese Numero O.C.(0 para eliminar referencia a O.C.)", , gnumoc)
      If IsNumeric(X) Then
         gnumoc = Val(X)
         If gnumoc <> 0 Then
           estado = "O"
         Else
           estado = "P"
         End If
         QUERY = "update a11 set  [num_oc]=" & gnumoc & " , [estado]='" & estado & "'"
         QUERY = QUERY & " where [num_int_item]= " & Val(msf1.TextMatrix(msf1.Row, 1))
         cn1.BeginTrans
         cn1.Execute QUERY
         cn1.CommitTrans
      
      End If
   End If
  End If
End If
If KeyCode = vbKeyInsert And c_prov.ListIndex > 0 Then
   stk_pedido2.t_renglon = ""
   stk_pedido2.t_cantidad = ""
   stk_pedido2.Show
   stk_pedido2.Show
End If
End Sub

Sub graba()
      
     For i = 1 To msf1.Rows - 1
      If Val(msf1.TextMatrix(i, 1)) > 0 Then
        'modifica el item
         QUERY = "update a11 set  [Detalle]='" & t_descripcion & "'"
         QUERY = QUERY & " where [id_grupo]= " & Val(t_id)
         cn1.BeginTrans
          cn1.Execute QUERY
          cn1.CommitTrans
      Else
        Set rs = New ADODB.Recordset
        q = "select * from g0 where [sucursal] = " & glo.sucursal
        rs.Open q, cn1, adOpenStatic, adLockOptimistic
        If Not rs.EOF And Not rs.BOF Then
          p = rs("ult_num_ref") + 1
          rs("ult_num_ref") = p
          rs.Update

          'agrego item
          QUERY = "INSERT INTO a11([num_int_item], [id_producto], [detalle], [cantidad], [id_obra], [observaciones], [estado], [cantidad_oc], [cantidad_ingresada], [fecha_pedido], [fecha_esperado], [id_usuario])"
          QUERY = QUERY & " VALUES (" & p & ", " & Val(msf1.TextMatrix(i, 2)) & ", '" & msf1.TextMatrix(i, 3) & "', " & msf1.TextMatrix(i, 4) & ", " & c_prov.ItemData(c_prov.ListIndex) & ", '" & msf1.TextMatrix(i, 5) & "', 'P', 0, 0, '" & Format$(Now, "dd/mm/yyyy") & "', '" & msf1.TextMatrix(i, 7) & "', " & para.id_usuario & ")"
          cn1.Execute QUERY
        End If
       End If
       
      Next i
      
      cn1.CommitTrans
      Set rs = Nothing
      
      Call INICIALIZA2(Me)
      Call armagrid
      c_prov.SetFocus

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
   If para.id_usuario = Val(msf1.TextMatrix(msf1.Row, 8)) And msf1.TextMatrix(msf1.Row, 10) = "P" Then
 
    stk_pedido2.t_renglon = msf1.Row
    stk_pedido2.t_basico = msf1.TextMatrix(msf1.Row, 2)
    stk_pedido2.t_detalle = msf1.TextMatrix(msf1.Row, 3)
    stk_pedido2.t_cantidad = msf1.TextMatrix(msf1.Row, 4)
    stk_pedido2.t_obs = msf1.TextMatrix(msf1.Row, 5)
    stk_pedido2.t_fechae = msf1.TextMatrix(msf1.Row, 7)
    stk_pedido2.t_ref = msf1.TextMatrix(msf1.Row, 1)
  
    stk_pedido2.Show
   Else
     MsgBox ("Solo el Usuario que ingrese el Item puede Modificarlo")
   End If
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
End Sub


Private Sub t_oc_GotFocus()
t_oc = ""
End Sub
