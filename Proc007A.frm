VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ABM_cotizacion 
   BackColor       =   &H00E0E0E0&
   Caption         =   "EMITIR SOLICITUD DE COTIZACION"
   ClientHeight    =   8775
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambiar"
      Height          =   855
      Left            =   240
      TabIndex        =   23
      Top             =   7320
      Width           =   1095
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   120
         Picture         =   "Proc007A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   240
      TabIndex        =   18
      Top             =   6360
      Width           =   9135
      Begin VB.TextBox t_obs 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   70
         TabIndex        =   7
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox t_condiciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   70
         TabIndex        =   6
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Condiciones de Compra:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4335
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7646
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   240
      TabIndex        =   12
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   7200
         Picture         =   "Proc007A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox t_tecontacto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7560
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox t_contacto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
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
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Text            =   "c_prov"
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Te:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   21
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Contacto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Solicitud de Cotizacion :"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   9
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Proc007A.frx":040F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Proc007A.frx":0C91
         Style           =   1  'Graphical
         TabIndex        =   10
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
      TabIndex        =   8
      Top             =   8520
      Width           =   11955
      _ExtentX        =   21087
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
            TextSave        =   "31/03/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:47 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "ABM_cotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2 where "
If tipo = "I" Then
  q = q & "  [id_producto] = " & Val(t_basico)
Else
  q = q & "  [cod_barra] = " & Val(t_basico)
End If
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  t_ip = rs("id_producto")
  t_pu = rs("PRECIO_ULT_COMPRA")
  c_tasa.ListIndex = rs("cod_tasaiva")
  t_detalle.Enabled = False
  t_precioultcompra = rs("PRECIO_ULT_COMPRA")
  t_fechaultcompra = rs("fecha_ULT_COMPRA")
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
End If
Set rs = Nothing
End Sub


Sub carga_oc()
  Set cl_comp = New COMPROBANTES
  Call cl_comp.cargar(70, "X", Val(t_sucursal), Val(t_numoc), 0)
  If cl_comp.numint = 0 Then
    EXISTE = "N"
    
  Else
     EXISTE = "S"
     MsgBox ("La Solicitud de Cotizacion ya existe en el Sistema")
     Set rs = New ADODB.Recordset
     q = "select * from a6 where [num_int] = " & cl_comp.numint
     t_fecha = cl_comp.fecha
     c_prov.ListIndex = buscaindice(c_prov, cl_comp.idproveedor)
     rs.Open q, cn1
     While Not rs.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs("id_producto"), "00000") & Chr(9) & rs("detalle") & Chr(9) & "0.00" & Chr(9) & rs("cantidad") & Chr(9) & Format$(rs("id_requisicion"), "00000") & Chr(9) & rs("ESTADO")
       rs.MoveNext
     Wend
     Set rs = Nothing
  End If
End Sub

Private Sub btnacepta_Click()
 If msf1.Rows > 1 Then
  J = MsgBox("Confirma Grabar Solicitud Coptizacion", 4)
  If J = 6 Then
   If verificaperiodog(t_fecha) = "A" Then
     Call graba
   Else
     MsgBox ("El periodo para el cual desea ingresar el comprobante esta CERRADO!!!!")
  End If
 End If
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 4000
msf1.ColWidth(3) = 2500
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 800
msf1.ColWidth(6) = 2000
msf1.ColWidth(7) = 600



msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Observaciones"
msf1.TextMatrix(0, 4) = "Cantidad"
msf1.TextMatrix(0, 5) = "Unid."
msf1.TextMatrix(0, 6) = "Obra/Destino"
msf1.TextMatrix(0, 7) = "Id.Obra"


End Sub





Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
  End If
End If
Set cl_prov = New proveedores
cl_prov.carga (c_prov.ItemData(c_prov.ListIndex))
If cl_prov.idprov > 0 Then
   t_contacto = cl_prov.contacto
   t_tecontacto = cl_prov.tecontacto
End If
Set cl_prov = Nothing
End Sub



Private Sub Command1_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Command2_Click()
ABM_PROv.Show
End Sub

Private Sub Command2_LostFocus()
c_prov.clear
Call carga_proveedores(c_prov)
c_prov.ListIndex = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
    gen_tools.Show
    
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 7)
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
Call numera
End Sub
Sub numera()
q = "select * from g2 where [id_tipo_comp] = 70 "
Set rs = New ADODB.Recordset
rs.MaxRecords = 1
rs.Open q, cn1

If Not rs.EOF And Not rs.BOF Then
  t_numoc = rs("ult_num") + 1
Else
  MsgBox ("Error al inicializar comprobante")
  Exit Sub
End If
Set rs = Nothing

End Sub
Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Continua - [F3] Pendientes"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  abm_solmat.Show
End If


If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If

If KeyCode = vbKeyF9 Then
  Frame2.Enabled = True
  t_condiciones.SetFocus
End If



If KeyCode = vbKeyInsert Then
  If msf1.Row < 40 Then
   abm_cotizacion1.t_renglon = ""
   abm_cotizacion1.t_basico = ""
   abm_cotizacion1.t_detalle = ""
   abm_cotizacion1.t_renglonp = ""
   abm_cotizacion1.t_cantunit = ""
   abm_cotizacion1.t_unidad = ""
   abm_cotizacion1.Show
 End If
End If
End Sub

Sub graba()
If EXISTE = "S" Then
  Set cl_comp = New COMPROBANTES
  Call cl_comp.cargar(70, "X", Val(t_sucursal), Val(t_numoc), 0)
  If cl_comp.numint <> 0 Then
    J = MsgBox("Comprobante existente, desea modificar", 4)
    If J = 6 Then
      cl_comp.borrar
      EXISTE = "N"
    Else
      EXISTE = "S"
    End If
  End If
  Set cl_comp = Nothing
End If

If EXISTE = "N" Then
   'oc nueva
      On Error GoTo ERRORGRABA
      numint = saca_ultnumero_int_comp("C")
      t_numoc = Format$(saca_ultnumero_comp(70), "00000000")
      
      Set cl_comp = New COMPROBANTES
      cl_comp.actual (70)
      STOCK = cl_comp.STOCK
      ctacte = cl_comp.ctacte
      moneda = "P"
      
      infocontacto = Left$(RTrim$(t_contacto) & "  " & RTrim$(t_tecontacto), 80)
      cn1.BeginTrans
      QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], [iva], [no_grabado], [percep_ret], [total], [fecha_prob_entrega], [fecha_recepcion], [estado], [ID_CODRETGAN], [ID_CUENTA], [STOCK], [CTACTE], [grabado], [estado_pago], [num_op], [obs], [condiciones], [info_contacto], [moneda], [cotiz_dolar], [contado])"
      QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numoc) & ", 'X', 70, " & c_prov.ItemData(c_prov.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & ", " & 0 & ", " & 0 & ", " & 0 & ", 0, " & 0 & ",'" & t_fecha & "', '" & t_fecha & "', 'P', 0, 0, '" & STOCK & "', '" & ctacte & "', '" & cl_comp.grabado & "', 'X', '0000-00000000'" & ", '" & t_obs & "', '" & t_condiciones & "', '" & infocontacto & "', '" & moneda & "', " & 1 & ", 'S')"
      'MsgBox (QUERY)
      cn1.Execute QUERY
      
      
      
      For i = 1 To msf1.Rows - 1
        nr = 0
        QUERY = "INSERT INTO a6([num_int], [RENGLON], [id_producto], [detalle], [cantidad], [pu], [importe], [envase], [bultos],[id_requisicion],[estado], [tasa_iva], [renglon_requisicion], [observaciones], [num_int_item], [unidad], [id_obra])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 4)) & ", " & 0 & ", " & 0 & ", 0, 0, 0," & " 'P', " & 0 & ", 0,'" & Left$(msf1.TextMatrix(i, 3) & " ", 30) & "', " & nr & ", '" & msf1.TextMatrix(i, 5) & "', " & Val(msf1.TextMatrix(i, 7)) & ")"
       
        cn1.Execute QUERY
      
      Next i
      
      cn1.CommitTrans
      Set rs = Nothing
      
      J = MsgBox("Imprime Cotizacion", 4)
      If J = 6 Then
         Set cl_comp = New COMPROBANTES
         cl_comp.cargar2 (numint)
         If cl_comp.numint > 0 Then
           cl_comp.imprimir
         End If
      End If
 
      
      Call INICIALIZA2(Me)
      Call armagrid
      t_sucursal = Format$(glo.sucursal, "0000")
      Call numera
      t_numoc.SetFocus
Else
   MsgBox ("No se puede modificar Cotizacion")
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    abm_cotizacion1.limpia
    abm_cotizacion1.t_renglon = msf1.Row
    abm_cotizacion1.t_basico = msf1.TextMatrix(msf1.Row, 1)
    abm_cotizacion1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    abm_cotizacion1.t_cantunit = msf1.TextMatrix(msf1.Row, 4)
    abm_cotizacion1.t_obs = msf1.TextMatrix(msf1.Row, 3)
    abm_cotizacion1.t_unidad = msf1.TextMatrix(msf1.Row, 5)
    abm_cotizacion1.c_obra.ListIndex = buscaindice(abm_cotizacion1.c_obra, Val(msf1.TextMatrix(msf1.Row, 7)))

    abm_cotizacion1.Show
  End If

End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
End Sub

Private Sub t_condiciones_LostFocus()
t_condiciones = RTrim$(t_condiciones) & " "
End Sub

Private Sub t_contacto_LostFocus()
t_contacto = RTrim$(t_contacto) & " "
End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If

If verificaperiodo(t_fecha) = "C" Then
   MsgBox ("El periodo para el cual se deseas ingresar el comprobante esta CERRADO!!!!!")
   t_fecha.SetFocus
   t_fecha = ""
Else
   Call verifica_fechacorte(t_fecha)
End If


End Sub

Private Sub t_numoc_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_numoc_LostFocus()
 Call carga_oc
End Sub

Private Sub t_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If

End Sub

Private Sub t_obs_LostFocus()
t_obs = RTrim$(t_obs) & " "
End Sub

Private Sub t_tecontacto_LostFocus()
t_tecontacto = RTrim$(t_tecontacto) & " "
End Sub

