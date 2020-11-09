VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form stk_recepcion 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEPCION DE MERCADERIA"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9340
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   8295
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   3
         Top             =   1440
         Width           =   5895
      End
      Begin VB.TextBox t_fechaprob 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox t_numoc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_sucursal 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Detalle:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Ingreso:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Parte Recepcion Nro.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   7
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "stk003.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "stk003.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   8
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
      TabIndex        =   6
      Top             =   6555
      Width           =   8235
      _ExtentX        =   14526
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
            TextSave        =   "23/01/07"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:36 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "stk_recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Sub carga_oc()
  Set cl_comp = New COMPROBANTES
  Call cl_comp.cargar(101, "X", Val(t_sucursal), Val(t_numoc), 0)
  If cl_comp.numint = 0 Then
     MsgBox ("Parte de Ingreso Inexistente. Si deseas Ingresar un nuevo pedido  deje en blanco el campo numero")
     t_numoc = ""
  Else
     Set rs = New ADODB.Recordset
     q = "select * from a6 where [num_int] = " & cl_comp.numint
     t_fecha = cl_comp.fecha
     t_fechaprob = cl_comp.fechaprobentrega
     c_prov.ListIndex = buscaindice(c_prov, cl_comp.idproveedor)
     
     rs.Open q, cn1
     While Not rs.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs("id_producto"), "00000") & Chr(9) & rs("detalle") & Chr(9) & rs("cantidad")
       rs.MoveNext
     Wend
     Set rs = Nothing
  End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 5000
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 800
msf1.ColWidth(5) = 800

msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Cantidad"


End Sub







Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 4)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()

Call carga_proveedores(c_prov)
c_prov.ListIndex = 0
t_sucursal = Format$(glo.sucursal, "0000")
Call armagrid
Call barraesag(Me)


End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If

If KeyCode = vbKeyF9 Then
 If msf1.Rows > 1 Then
  j = MsgBox("Confirma Grabar Parte de Recepcion", 4)
  If j = 6 Then
   Call graba
  End If
 End If
End If

If KeyCode = vbKeyInsert Then
   stk_recepcion2.t_renglon = ""
   stk_recepcion2.t_cantidad = ""
   stk_recepcion2.Show
   stk_recepcion2.Show
End If
End Sub

Sub graba()
If EXISTE = "N" Then
   'oc nueva
      'On Error GoTo ERRORGRABA
      numint = saca_ultnumero_int_comp("C")
      t_numoc = Format$(saca_ultnumero_comp(101), "00000000")
      
      Set cl_comp = New COMPROBANTES
      cl_comp.ACTUAL (101)
      STOCK = cl_comp.STOCK
      ctacte = cl_comp.ctacte
      
      
      cn1.BeginTrans
      QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], [iva], [no_grabado], [percep_ret], [total], [fecha_prob_entrega], [fecha_recepcion], [estado], [ID_CODRETGAN], [ID_CUENTA], [STOCK], [CTACTE], [grabado], [estado_pago], [num_op])"
      QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numoc) & ", 'X', 101, " & c_prov.ItemData(c_prov.ListIndex) & ", '" & Format$(Now, "dd/mm/yyyy") & "', " & para.id_usuario & ", 0, 0, 0, 0, 0,'" & t_fechaprob & "', '" & t_fechaprob & "', 'P', 0, 0, '" & STOCK & "', '" & ctacte & "', '" & cl_comp.grabado & "', 'X', '0000-00000000'" & ")"
      cn1.Execute QUERY
      
      
      
      For i = 1 To msf1.Rows - 1
        QUERY = "INSERT INTO a6([num_int], [RENGLON], [id_producto], [detalle], [cantidad], [pu], [importe], [envase], [bultos],[id_requisicion],[estado], [tasa_iva], [cantidad_recibida])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 3)) & ", 0, 0, 0, 0, 0,'P',0, 0)"
        cn1.Execute QUERY
      
        cantidadp = Val(msf1.TextMatrix(i, 3))
        Set cl_prod = New productos
        Call cl_prod.actualizar(Val(msf1.TextMatrix(i, 1)), cantidadp, 0, -cantidadp)
        Set cl_prod = Nothing
          
        QUERY = "update a6 set  [CANTIDAD_RECIBIDA]=[CANTIDAD_RECIBIDA] + " & Val(msf1.TextMatrix(i, 3))
        QUERY = QUERY & " where [NUM_INT]= " & Val(msf1.TextMatrix(i, 4)) & " AND [RENGLON]=" & Val(msf1.TextMatrix(i, 5))
        cn1.Execute QUERY
        
        
      Next i
      
      cn1.CommitTrans
      Set rs = Nothing
      
      j = MsgBox("Imprime Parte", 4)
      If j = 6 Then
         Set rs = New ADODB.Recordset
         q = "select * from a5, a6, a1, g1 where a5.[num_int] = " & numint & " and a5.[num_int] = a6.[num_int] and a5.[id_proveedor] = a1.[id_proveedor] and a5.[id_usuario] = g1.[id_usuario]"
         rs.Open q, cn1
         Call ejecutareporte2(rs, com_recepcion)
         Set rs = Nothing
      End If

      
      Call INICIALIZA2(Me)
      Call armagrid
      t_numoc.SetFocus
Else
   MsgBox ("No se puede modificar Nota de Pedido")
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    stk_recepcion2.t_renglon = msf1.Row
    stk_recepcion2.t_basico = msf1.TextMatrix(msf1.Row, 1)
    stk_recepcion2.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    stk_recepcion2.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
    stk_recepcion2.t_nroreq = msf1.TextMatrix(msf1.Row, 4)
    stk_recepcion2.t_renglonp = msf1.TextMatrix(msf1.Row, 5)
    stk_recepcion2.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
End Sub


Private Sub t_fechaprob_LostFocus()
If Not IsDate(t_fechaprob) Then
  t_fechaprob = Format$(Now, "dd/mm/yyyy")
Else
  t_fechaprob = Format$(t_fechaprob, "dd/mm/yyyy")
End If
End Sub

Private Sub t_numoc_GotFocus()
Call INICIALIZA2(Me)
End Sub

Private Sub t_numoc_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_numoc_LostFocus()
If t_numoc = "" Then
    EXISTE = "N"
Else
   Call carga_oc
End If
End Sub
