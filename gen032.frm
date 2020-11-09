VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_depuraperiodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ELIMINA MOVIMIENTOS DE UN PERIODO"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7170
   ScaleWidth      =   12495
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Modulos"
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   10575
      Begin VB.CheckBox Check3 
         Caption         =   "Caja"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Stock por Movimientos"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cuentas Corrientes"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10575
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox t_f2 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Año:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10800
      TabIndex        =   2
      Top             =   5760
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "gen032.frx":0000
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
         Picture         =   "gen032.frx":0882
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
      Top             =   6915
      Width           =   12495
      _ExtentX        =   22040
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:39"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Todas las terminales tienen que estar fuera del sistema"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   6240
      Width           =   8895
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Realice siempre un BACKUP  antes de ejecutarlo."
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5880
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Este proceso es Irreversible, NO lo ejecute si no esta seguro"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5520
      Width           =   8895
   End
End
Attribute VB_Name = "gen_depuraperiodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Dim fechacorte As String
Dim sa As Double
Dim sao As Double
Dim idcliente As Long
Dim dcliente As String
Dim idcuentacontra As Long


Private Sub btnacepta_Click()
J = MsgBox("Este proceso elimina datos en forma definitiva, desdea continuar?", 4)
If J = 6 Then
 If Val(t_f2) > 1999 Then
   fechacorte = "31/12/" & t_f2
   J = MsgBox("Este proceso es importante e irreversible. Vuelva a Confirmar?", 4)
   If J = 6 Then
     Call graba
   End If
 End If
End If
End Sub


Sub graba()
If Check1 = 1 Then
 Call ctactes
End If


If Check2 = 1 Then
 Call STOCK
End If

If Check3 = 1 Then
 Call caja
End If


End Sub











Private Sub btnsale_Click()
Unload Me
End Sub


Sub ctactes()
espere.Show
Set rs7 = New ADODB.Recordset
q = "select * from vta_01  "
rs7.Open q, cn1
While Not rs7.EOF
  'para cada cliente saco saldo anterior
   espere.Label1 = rs7("id_cliente")
   espere.Refresh
   q = "select * from vta_02 where [id_cliente] = " & rs7("id_cliente")
   q = q & " and datevalue([fecha]) <= datevalue('" & fechacorte & "')"
   Set rs2 = New ADODB.Recordset
   rs2.Open q, cn1, adOpenStatic, adLockOptimistic
   sa = 0
   sao = 0
   da = 0
   dao = 0
   ha = 0
   hao = 0
   
   While Not rs2.EOF
    If rs7("id_cliente") > 1 And rs2("cta_cte") <> "N" And rs2("contado") = "N" Then
       If rs2("moneda") = "P" Then
        t = rs2("total")
        T2 = rs2("total_otra_moneda")
      Else
        t = rs2("total_otra_moneda")
        T2 = rs2("total")
      End If
      
     If rs2("cta_cte") = "D" Then
        da = da + t
        dao = dao + T2
     Else
        ha = ha + t
        hao = hao + T2
     End If
   
     sa = da - ha
     sao = dao - hao
   End If
   
   ni = rs2("num_int")
  
   rs2.MoveNext
   
  cn1.BeginTrans
      
   q = "delete from vta_09 where [num_int] = " & ni
   cn1.Execute q
   
  q = "delete from vta_011 where [num_int] = " & ni
  cn1.Execute q
      
  q = "delete from vta_012 where [num_int] = " & ni
   cn1.Execute q
   
  q = "delete from vta_015 where [num_int] = " & ni
  cn1.Execute q
      
  q = "delete from vta_03 where [num_int] = " & ni
  cn1.Execute q
      
  q = "delete from vta_02 where [num_int] = " & ni
  cn1.Execute q
  
  cn1.CommitTrans
 Wend
 Set rs2 = Nothing
 'termino el cliente
  If rs7("id_cliente") > 1 And (sa > 0 Or sao > 0) Then
    'genero movimiento en cuenta corriente
     idcliente = rs7("id_cliente")
     Call generoajuste
    
    
  End If
  
  rs7.MoveNext
 
Wend
Set rs = Nothing
Unload espere
MsgBox ("Proceso Cuentas Corrientes Terminado")

End Sub

Sub caja()
espere.Show
Set rs7 = New ADODB.Recordset
q = "select * from cyb_01 where [caja] = 'S'"
rs7.Open q, cn1
While Not rs7.EOF
  'para cada cliente saco saldo anterior
   espere.Label1 = rs7("id_forma_pago")
   espere.Refresh
   Set rs8 = New ADODB.Recordset
   q = "select * from c_01 where [tipo]= 'C'"
   rs8.Open q, cn1
   While Not rs8.EOF
    q = "select * from cyb_05 where [id_forma_pago] = " & rs7("id_forma_pago") & " and [id_cuenta_contra] = " & rs8("id_cuenta")
    q = q & " and datevalue([fecha]) <= datevalue('" & fechacorte & "')"
    Set rs2 = New ADODB.Recordset
    rs2.Open q, cn1, adOpenStatic, adLockOptimistic
    sa = 0
    da = 0
    ha = 0
   
    While Not rs2.EOF
   
     t = rs2("importe")
      
     If rs2("ubicacion") = "D" Then
        da = da + t
        
     Else
        ha = ha + t
        
     End If
   
     sa = da - ha
    
     ni = rs2("num_mov_caja")
  
     rs2.MoveNext
   
    cn1.BeginTrans
    q = "delete from cyb_04 where [num_mov_int] = " & ni & " and [modulo]= 'J' "
    cn1.Execute q
         
    q = "delete from cyb_05 where [num_mov_caja] = " & ni
    cn1.Execute q
  
    cn1.CommitTrans
  Wend
  Set rs2 = Nothing
 'termino el cliente
  If sa <> 0 Then
    'genero movimiento en caja
     idcliente = rs7("id_forma_pago")
     idcuentacontra = rs8("id_cuenta")
     Call generoajustecaja
  End If
  rs8.MoveNext
 Wend
 Set rs8 = Nothing
 rs7.MoveNext
 
Wend
Set rs = Nothing
Unload espere
MsgBox ("Proceso Caja Terminado")

End Sub
Sub STOCK()
espere.Show
espere.Label1 = "Espere Borrando Movimientos de Stock"
espere.Refresh
'borro stk_o2 y stock _03
Set rs4 = New ADODB.Recordset
q = "select * from stk_02 where datevalue([fecha]) <= datevalue('" & fechacorte & "') "
rs4.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs4.EOF
  cn1.BeginTrans
  q = "delete from stk_03 where [num_int] = " & rs4("num_int")
  cn1.Execute q
  cn1.CommitTrans

 rs4.Delete
 rs4.MoveNext
Wend
Set rs4 = Nothing
  
  

Set rs7 = New ADODB.Recordset
q = "select * from a2 "
rs7.Open q, cn1
While Not rs7.EOF
  'para cada producto saco saldo anterior
   espere.Label1 = "Espere Calculando stock producto " & rs7("id_producto")
   espere.Refresh
   q = "select * from stk_01 where [id_producto] = " & rs7("id_producto")
   q = q & " and datevalue([fecha]) <= datevalue('" & fechacorte & "')"
   Set rs2 = New ADODB.Recordset
   rs2.Open q, cn1, adOpenStatic, adLockOptimistic
   sa = 0
   sao = 0
   da = 0
   dao = 0
   ha = 0
   hao = 0
   
   While Not rs2.EOF
    If rs7("id_producto") > 1 Then
       If rs2("ubicacion") = "E" Then
        da = da + rs2("cantidad")
       Else
        ha = ha + rs2("cantidad")
       End If
        sa = da - ha
    End If
   
   ni = rs2("num_mov_stk")
  
   rs2.MoveNext
   
  cn1.BeginTrans
     
  q = "delete from stk_01 where [num_mov_stk] = " & ni
  cn1.Execute q
  
  cn1.CommitTrans
 Wend
 Set rs2 = Nothing
 'termino el producto
  If rs7("id_producto") > 1 And sa <> 0 Then
    'genero movimiento en cuenta corriente
     idcliente = rs7("id_producto")
     dcliente = rs7("descripcion")
     Call generoajustestock
    
    
  End If
  
  rs7.MoveNext
 
Wend
Set rs = Nothing




Unload espere
MsgBox ("Proceso Stock Terminado")

End Sub



Sub generoajuste()
  numint = saca_ultnumero_int_comp("V")
    
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(1)
  If sa > 0 Or sao > 0 Then
    tipoc = 80
  Else
    tipoc = 85
  End If
  cl_compvta.actual (tipoc)
  abreviatura = cl_compvta.abreviatura
  ubicacionctacte = cl_compvta.ctacte
  numc = cl_compvta.numcomp
  ep = "S"
  cp = "0000-00000000"
  contado = "N"
  ssi = 0
      
  moneda = "P"
      
  
  codact = 0
  alicuotaib = 0
  cuentaact = para.cuenta_ventas
  T2 = sao
 
      
        
  Set cl_cli = New Clientes
  cl_cli.carga (idcliente)
              
  tiporespiva = cl_cli.idtipoiva
  idcli = idcliente
  letrac = cl_cli.letra
      
  cl_compvta.letra = letrac
  cl_compvta.SACANUMCOMP
  
  
  
 
    cn1.BeginTrans
       
       
       
    QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
    "[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
    " [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
    " [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], [dominio_acoplado02], [SALDO_IMPAGO02], [num_z])"



    QUERY = QUERY & " VALUES (" & numint & ", " & Val(1) & ", " & cl_compvta.numcomp & ", '" & letrac & "', " & tipoc & _
    ", " & idcli & ", '" & fechacorte & "', " & para.id_usuario & ", " & Val(0) & ", " & Val(0) & ", " & Val(0) & ", " & Val(sa) & _
    ", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & "Cierre" & _
    " ', " & Val(1) & ", " & T2 & ", '" & moneda & "', 1, '" & cl_compvta.venta & "', '" & contado & "', " & Val(0) & _
    ", 0, " & Val(0) & ", " & codact & ", " & Val(0) & ", " & Val(0) & ", 0, '" & fechacorte & "', 0, " & Val(0) & ", ' ', ' ', ' ', 0, " & Val(1) & _
    ", '" & Left$(cl_cli.razonsocial, 50) & "', '" & Left$(cl_cli.direccion, 50) & "', '" & Left$(cl_cli.CUIT, 20) & "', '" & Left$(cl_cli.localidad, 50) & "', " & tiporespiva & ", ' ', ' ', ' ', " & ssi & ", " & para.z_actual & ")"

   ' MsgBox (QUERY)
    
    cn1.Execute QUERY
      
      
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Emitir Cierre NI:" & numint & "', " & para.id_usuario & ", 'V', " & numint & ", '" & Now & "', '[" & tipoc & "] " & "', 12, " & idcli & ")"
  
     cn1.Execute QUERY

      cn1.CommitTrans
      
     cl_compvta.ACTUALIZA_NUMERADOR
     
 

Set cl_compvta = Nothing
Set cl_cli = Nothing

End Sub



Sub generoajustecaja()


   Set rs1 = New ADODB.Recordset
   q = "select * from cyb_01 where [id_forma_pago] = " & idcliente
   rs1.Open q, cn1
   If Not rs1.BOF And Not rs1.EOF Then
      cta = rs1("id_cuenta_cont")
   Else
      cta = 0
   End If
   Set rs1 = Nothing
   
   Set rs = New ADODB.Recordset
   
    q = "SELECT * FROM CYB_05 "
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    rs.AddNew
                 
   If sa > 0 Then
      u = "D"
      impo = sa
   Else
      u = "H"
      impo = -sa
   End If
   rs("ID_forma_pago") = idcliente
   rs("id_cuenta_caja") = cta
   rs("id_cuenta_contra") = idcuentacontra
   rs("Descripcion") = "Cierre"
   rs("Importe") = impo
   rs("ubicacion") = u
   rs("fecha") = fechacorte
   rs("num_mov_int") = rs("num_mov_caja")
   rs("modulo") = "J"
   rs("Operacion") = "Mov.Caja " & Format$(rs("num_mov_caja"), "00000000")
   rs("id_usuario") = para.id_usuario
   numint = rs("num_mov_caja")
   t_numint = numint
   rs.Update
   Set rs = Nothing
   
   
End Sub
Sub generoajustestock()
cn1.BeginTrans
      QUERY = "INSERT INTO stk_02([fecha], [letra], [num_comprobante], [id_usuario], [detalle], [sucursal], [tipo_comprobante], [id_proveedor], [id_obra])"
      QUERY = QUERY & " VALUES ('" & fechacorte & "', 'X', 0, " & para.id_usuario & ", 'Cierre ', 0, 1, 1,1)"
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs = cn1.Execute(qr)
      numint = rs.Fields("NewID").Value

       If sa > 0 Then
         u = "E"
       Else
         u = "S"
       End If
      
    
        QUERY = "INSERT INTO stk_03([num_int], [RENGLON], [id_producto], [descripcion], [unidad], [detalle], [cantidad], [ubicacion])"
        QUERY = QUERY & " VALUES (" & numint & ", 1, " & idcliente & ", '" & Left$(dcliente, 50) & "', ' ', 'cierre', " & sa & " , '" & u & "')"
        cn1.Execute QUERY
      
        QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo])"
        QUERY = QUERY & " VALUES ('" & fechacorte & "', " & idcliente & ", " & sa & ", '" & u & "', 'Mov.Int.Stk " & Format$(numint, "00000000") & "', 'Cierre', " & numint & ", 'S')"
        cn1.Execute QUERY
  cn1.CommitTrans
  
End Sub
Private Sub Form_Load()
t_f2 = "2010"



End Sub







Private Sub UpDown2_DownClick()
If Val(t_f2) > 1999 Then
  t_f2 = Val(t_f2) - 1
End If
End Sub

Private Sub UpDown2_UpClick()
If Val(t_f2) < 2050 Then
  t_f2 = Val(t_f2) + 1
End If

End Sub
