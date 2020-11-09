VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_movbanco 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debitos y Creditos Bancarios"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   29
      Top             =   3960
      Width           =   6255
      Begin VB.Label Label12 
         Caption         =   "Si se ingresa importe en IVA se generará un movimiento en iva compras como contado al proveedor asignado en el Banco"
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Movimiento"
      Height          =   3855
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   6255
      Begin VB.TextBox t_numcomp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_suc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox t_subtotal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   6
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox t_iva 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3120
         MaxLength       =   12
         TabIndex        =   7
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   5640
         Picture         =   "cyb014.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3240
         Width           =   375
      End
      Begin VB.ComboBox c_op 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Text            =   "c_cuenta"
         Top             =   1800
         Width           =   4095
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Text            =   "c_cuenta"
         Top             =   3360
         Width           =   4095
      End
      Begin VB.ComboBox c_caja 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox t_op 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   9
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox t_importe 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         MaxLength       =   12
         TabIndex        =   8
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox t_destino 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   5
         Top             =   2280
         Width           =   4455
      End
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox t_numint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
         Caption         =   "Num. Comprob."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "Subtotal:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Iva:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Tipo Operacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Concepto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Banco:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Operacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Total:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Detalle(F6):"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800080&
         Caption         =   "Num. Interno:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1560
      TabIndex        =   11
      Top             =   5040
      Width           =   3135
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Emitir"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   5835
      Width           =   6600
      _ExtentX        =   11642
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:43"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_movbanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim contab As String
Dim abreviatura As String
Dim numintb As Long
Dim ctabanco As Long
Dim nmic As Long

Sub graba()
   
   If Val(t_iva) > 0 Then
     J = MsgBox("El importe de Iva es mayor que cero, registra movimiento en libro iva compras", 4)
     If J = 6 Then
       Call compprov
     Else
       nmic = 0
     End If
   End If
   
   
   Set rs1 = New ADODB.Recordset
   q = "select * from cyb_01 where [id_forma_pago] = " & c_caja.ItemData(c_caja.ListIndex)
   rs1.Open q, cn1
   If Not rs1.BOF And Not rs1.EOF Then
      ctabanco = rs1("id_cuenta_cont")
   Else
      ctabanco = 0
   End If
   Set rs1 = Nothing
   
   If t_op = "D" Then
     tipoo = 20
   Else
     tipoo = 30
   End If
   
   Set rs2 = New ADODB.Recordset
   q = "select * from cyb_06 where [id_tipomov] = " & tipoo
   rs2.Open q, cn1
   If Not rs2.BOF And Not rs2.EOF Then
      u = rs2("ubicacion")
      contab = rs2("contabilidad")
      abreviatura = rs2("abreviatura")
   Else
      u = tipoop
      contab = "N"
      abreviatura = "Mov.Bco."
   End If
   Set rs2 = Nothing
   
   Set rs = New ADODB.Recordset
   If t_numint <> "" Then
     q = "SELECT * FROM CYB_04 where [num_mov_banco] = " & Val(t_numint)
     rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   Else
     q = "SELECT * FROM CYB_04 "
     rs.Open q, cn1, adOpenDynamic, adLockOptimistic
     rs.AddNew
   End If
     
   rs("id_banco") = c_caja.ItemData(c_caja.ListIndex)
   rs("fecha") = t_fecha
   rs("Detalle") = t_destino
   rs("Importe") = Val(t_importe)
   rs("ubicacion") = u
   rs("id_tipomov") = tipoo
   rs("fecha_dif") = t_fecha
   rs("entro") = "S"
   rs("fecha_acreed") = t_fecha
   rs("num_comp") = 0
   rs("entro") = "S"
   rs("modulo") = "B"
   rs("num_mov_int") = 0
   rs("id_tipodbcr") = c_op.ItemData(c_op.ListIndex)
    rs("num_mov_int_compras") = nmic
    rs("id_cuenta04") = c_cuenta.ItemData(c_cuenta.ListIndex)
   
   
   rs.Update
   numintb = rs("num_mov_banco")
   Set rs = Nothing
   
   
   If contab <> "N" And c_cuenta.ListIndex > 0 Then
     'grabo contabilidad
     Call grabacontab
     
     
   End If
   
   
   
End Sub

Sub compprov()
      numint2 = saca_ultnumero_int_comp("C")
      
      Set cl_comp = New COMPROBANTES
      
      If t_op = "D" Then
         cl_comp.actual (20)
         cl_comp.idtipocomp = 20
      Else
          cl_comp.actual (30)
           cl_comp.idtipocomp = 30
      End If
      
      If c_cuenta.ListIndex = 0 Then
            c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_compras_varias)
      End If
      
      'contado
      ep = "S"
      cp = "ctdo"
      cc = "N"
      ssi = 0
      
      
      t_destino = RTrim$(t_destino) & " "
      
       moneda = "P"
       tom = 0

      Set rs = New ADODB.Recordset
      q = "select * from cyb_01 where [id_forma_pago] = " & c_caja.ItemData(c_caja.ListIndex)
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        idprov = rs("id_proveedor01")
      Else
        idprov = 1
      End If
      Set rs = Nothing
     cn1.BeginTrans
      
QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], " & _
" [no_grabado], [percep_ret], [iva], [total], [fecha_prob_entrega], [fecha_recepcion], [estado], [id_codretgan], [id_cuenta], [stock], [ctacte], [grabado], " & _
" [estado_pago], [num_op], [id_codretib], [obs], [condiciones], [info_contacto], [moneda], [cotiz_dolar], [contado], [TOTAL_D], [monto_suj_ret], " & _
"[alicuota_ret], [ret_mes], [pagos_realizados], [pago_actual], [minimo_no_imp], [fecha_vto], [COMPRA], [saldo_impago])"
      
 QUERY = QUERY & " VALUES (" & numint2 & ", " & Val(t_suc) & ", " & Val(t_numcomp) & ", 'A', " & cl_comp.idtipocomp & _
 ", " & idprov & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotal) & ", " & Val(t_nograbado) & ", " & Val(t_perc) & ", " & Val(t_iva) & _
 ", " & Val(t_importe) & ", '" & Format$(Now, "dd/mm/yyyy") & "', '" & t_fecha & "', 'A', 1, " & c_cuenta.ItemData(c_cuenta.ListIndex) & _
 ", 'N', '" & t_op & "', '" & cl_comp.grabado & "', '" & ep & "', '" & cp & "', 1, '" & t_destino & "', ' ', ' ', '" & moneda & "', " & _
 1 & ", '" & ep & "', " & tom & ", 0, 0, 0, 0, 0, 0, '" & t_fecha & "', 'N', " & ssi & ")"
   
      
      cn1.Execute QUERY
      
    nmic = numint2
      
      
      

      
      
      
     If cl_comp.contabilidad <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")

         
         u1 = cl_comp.contabilidad
          
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         
         
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[ComprasxBco] " & cl_comp.abreviatura & " " & Format$(Val(t_suc), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'C', " & numint2 & ", " & Val(t_importe) & ", " & Val(t_importe) & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_caja), 50) & "')"
         cn1.Execute QUERY
      
         ic = 1
        
         'sacocuenta banco
         Set rs = New ADODB.Recordset
         q = "select * from cyb_01 where [id_forma_pago] = " & c_caja.ItemData(c_caja.ListIndex)
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           cta = rs("id_cuenta_cont")
         Else
           cta = 1
         End If
         Set rs = Nothing
         
         'contado
         Set rs = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & cta
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
             dcta = rs("descripcion")
          Else
             dcta = "Cuenta Inexistente"
          End If
          Set rs = Nothing
              
          im = Format(Val(t_importe), "######0.00")
          QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
          QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & im & ", '" & dcta & "')"
            'MsgBox (QUERY)
          cn1.Execute QUERY
          
          ic = ic + 1
         
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_compras & ", '" & u2 & "', " & Val(t_iva) & ", 'IVA Banco')"
           cn1.Execute QUERY
           ic = ic + 1
         
         'contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & u2 & "', " & Val(t_subtotal) & ", '" & c_cuenta & "')"
         cn1.Execute QUERY
      
      End If
          
      
      
      cn1.CommitTrans

End Sub
Sub grabacontab()
If Generaasientosauto Then
         numintcgr = saca_ultnumero_int_comp("G")
         u1 = contab
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
          
         'grabo asiento
         cn1.BeginTrans
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[" & abreviatura & "] N.I." & Format$(numintb, "00000000") & "', 'B', " & numintb & ", " & Val(t_importe) & ", " & Val(t_importe) & ", " & para.id_usuario & ", '" & Left$(Detalle & " ", 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre bancos
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & ctabanco & ", '" & u1 & "', " & Val(t_importe) & ", '" & abreviatura & ".N.I" & Format$(numintb, "00000000") & "')"
         cn1.Execute QUERY
      
         'cta = rs("id_cuenta_cont")
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & 2 & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & u2 & "', " & Val(t_importe) & ", '" & t_descripcion & " ')"
         cn1.Execute QUERY
      
     
      cn1.CommitTrans
  End If
End Sub

Private Sub c_cuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If c_cuenta.ListIndex < 0 Then
   c_cuenta.ListIndex = 0
 End If
 J = MsgBox("Confirma Grabar Movimiento", 4)
 If J = 6 Then
   If verifica Then
    If verificaperiodog(t_fecha) = "A" Then
     Call graba
     Call limpia
     Frame2.Enabled = True
     Frame3.Enabled = False
    Else
     MsgBox ("Periodo cerrado. Imposible grabar operacion")
   End If
   End If
 End If
End If
End Sub
Function verifica() As Boolean
  v = True
  If t_fecha = "" Then
    v = False
    MsgBox ("Ingrese fecha")
  End If
  
  If Val(t_importe) <= 0 Then
    MsgBox ("Importe Incorrecto")
    v = False
  End If
  
  If t_op = "" Then
    MsgBox ("Ingrese Tipo de Operacion")
    v = False
  End If
  verifica = v

End Function

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  If Val(c_cuenta) > 0 Then
    c_cuenta.ListIndex = buscaindice(c_cuenta, Val(c_cuenta))
  Else
    c_cuenta.ListIndex = 0
  End If
End If
End Sub

Private Sub c_op_LostFocus()
If c_op.ListIndex < 0 Then
  c_op.ListIndex = 0
End If

Set rs = New ADODB.Recordset
q = "select * from cyb_07 where [id_tipomov] = " & c_op.ItemData(c_op.ListIndex)
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, rs("id_cuenta"))
Else
  c_cuenta.ListIndex = 0
End If
Set rs = Nothing
End Sub

Private Sub Command1_Click()
If t_numint = "" Then
   Call limpia
End If
   
Frame3.Enabled = True
c_caja.SetFocus
End Sub
Sub limpia()
t_fecha = ""
t_fechadif = ""
t_destino = ""
t_importe = ""
t_op = ""
t_suc = ""
t_numcomp = ""
t_subtotal = ""
t_iva = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
 J = MsgBox("Confirma Grabar Movimiento", 4)
 If J = 6 Then
   If verifica Then
    Call graba
    Call limpia
    Frame2.Enabled = True
    Frame3.Enabled = False
   End If
 End If

End If
End Sub

Private Sub Form_Activate()
Call barraesag(Me)



If para.cuenta_sel > 0 Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_sel)
  para.cuenta_sel = 0
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then

 Call tabup(Me)
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Frame2.Enabled = True
   Frame3.Enabled = False
End If

If KeyAscii = 13 Then
  Call TabEnter2(Me, 10)
End If
End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_formas_pago(c_caja, "B")
Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "<No Imputa>", 0
c_cuenta.ListIndex = 0
Call carga_dbcrbanco(c_op)
c_op.ListIndex = 0
Frame2.Enabled = True
Frame3.Enabled = False
End Sub

Private Sub t_destino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
  t_destino = c_op
End If
End Sub

Private Sub t_destino_LostFocus()
If t_destino = "" Then
  t_destino = "*"
End If
End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Or t_fecha = "" Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub


Private Sub t_importe_GotFocus()
t_importe = Format$(Val(t_subtotal) + Val(t_iva), "#####0.00")
End Sub

Private Sub t_numcomp_LostFocus()
t_numcomp = Format$(Val(t_numcomp), "00000000")

End Sub

Private Sub t_op_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[D] Debe/Debitos - [H] Haber/Creditos "
End Sub

Private Sub t_op_LostFocus()
 Call barraesag(Me)
 
 t_op = UCase$(t_op)
 If t_op <> "D" And t_op <> "H" Then
   t_op = "H"
 End If
End Sub

Private Sub t_suc_LostFocus()
t_suc = Format$(Val(t_suc), "0000")

End Sub
