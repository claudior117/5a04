VERSION 5.00
Begin VB.Form gen_sincronizar 
   Caption         =   "Sincronizacion con la Nube"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_dia 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox T_mes 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox t_año 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sincronizar con la Nube"
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label4"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   $"gen_sincronizar.frx":0000
      Height          =   855
      Left            =   720
      TabIndex        =   7
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Fecha Sincronizacion"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "Última Fecha de sincronización"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "gen_sincronizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 J = MsgBox("Desea sincronizar movimientos con la nube, asegurese de tener conexión a internet", 4)
 If J = 6 Then
   Call sincroniza
 End If
End Sub

Sub sincroniza()
'conexion remota a digital ocean
'usuario mysql = aero
'password mysql = s7VXRGBZ27E=
'ip = http://159.65.106.144

 
 
 Load espere
 espere.Label1 = "Conectando a Base de Datos Remota"
 espere.Show
 'On Error GoTo err1
 Set cnn_nube = New ADODB.Connection
  
 
 ' ****OJO ***** cambiar en load  Label4 = "Modo LOCALHOST ACTIVADO" 0 Label4 = "Modo SERVER ACTIVADO CUANDO SE CMABI ESTA LINEA"
 'cnn_nube.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=5a04;connection="
 cnn_nube.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=159.65.106.144;uid=aero;pwd=s7VXRGBZ27E=;database=5a04;connection="
 
 
 
 cnn_nube.Open
    
 'On Error GoTo err2
 m = "Paso 1/8 --> Borrando elementos Base de Datos remota-Clientes:Clientes"
 espere.Label1 = m
 espere.Refresh
 
 Set rs = New ADODB.Recordset
 q = "delete from clientes where id_sistema = " & para.idsistema
 rs.Open q, cnn_nube, adOpenDynamic, adLockOptimistic
 Set rs = Nothing
 
 
 espere.Label1 = "Paso 2/8 --> Borrando elementos Base de Datos remota-Ventas:Movimientos"
 espere.Refresh
 Set rs = New ADODB.Recordset
 q = "delete from ventas where id_sistema = " & para.idsistema
 f2 = ""
 f2i = ""
 f2 = t_dia & "-" & T_mes & "-" & t_año
 f2i = t_año & "-" & T_mes & "-" & t_dia
 If IsDate(f2i) Then
       q = q & " and fecha >= '" & f2i & "'"
 End If
 rs.Open q, cnn_nube, adOpenDynamic, adLockOptimistic
 Set rs = Nothing

 
 espere.Label1 = "Paso 3/8 --> Borrando elementos Base de Datos remota-Compras:Proveedores"
 espere.Refresh
  Set rs = New ADODB.Recordset
 q = "delete from proveedores where id_sistema = " & para.idsistema
 rs.Open q, cnn_nube, adOpenDynamic, adLockOptimistic
 Set rs = Nothing
 
 
 espere.Label1 = "Paso 4/8 --> Borrando elementos Base de Datos remota-Compras:Movimientos"
 espere.Refresh
 Set rs = New ADODB.Recordset
 q = "delete from compras where id_sistema = " & para.idsistema
 f2 = ""
 f2i = ""
 f2 = t_dia & "-" & T_mes & "-" & t_año
 f2i = t_año & "-" & T_mes & "-" & t_dia
 If IsDate(f2i) Then
       q = q & " and fecha >= '" & f2i & "'"
 End If
 rs.Open q, cnn_nube, adOpenDynamic, adLockOptimistic
 Set rs = Nothing
 
 
 'clientes
 m = "Paso 5/8 --> Subiendo informacion remota...CLIENTES"
 espere.Label1 = m
 espere.Refresh
 Set rs2 = New ADODB.Recordset
 q = "select * from vta_01"
 rs2.Open q, cn1
 cnn_nube.BeginTrans
 While Not rs2.EOF
       q = ""
       q = "INSERT INTO clientes(id_sistema, id_cliente , cliente)"
       q = q & " VALUES (" & para.idsistema & ", " & rs2("id_cliente") & ", '" & rs2("denominacion") & "')"
       cnn_nube.Execute q
 
 
       espere.Label1 = m & "   " & rs2("id_cliente")
       espere.Refresh
 
 
       rs2.MoveNext
 Wend
 cnn_nube.CommitTrans

Set rs2 = Nothing
 
 
 
'Proveedores
 m = "Paso 6/8 --> Subiendo informacion remota...PROVEEDORES"
 espere.Label1 = m
 espere.Refresh
 Set rs2 = New ADODB.Recordset
 q = "select * from a1"
 rs2.Open q, cn1
 cnn_nube.BeginTrans
 While Not rs2.EOF
       q = ""
       q = "INSERT INTO proveedores(id_sistema, id_prov , prov)"
       q = q & " VALUES (" & para.idsistema & ", " & rs2("id_proveedor") & ", '" & rs2("denominacion") & "')"
       cnn_nube.Execute q
 
 
       espere.Label1 = m & "   " & rs2("id_proveedor")
       espere.Refresh
 
 
       rs2.MoveNext
 Wend
 cnn_nube.CommitTrans

Set rs2 = Nothing
 
  
 
 m = "Paso 7/8 --> Subiendo informacion...VENTAS.Movimientos"
 espere.Label1 = m
 espere.Refresh
 Set rs2 = New ADODB.Recordset
 q = "select * from vta_02"
 If f2 <> "" Then
   If IsDate(f2) Then
       q = q & " where datevalue([fecha]) > datevalue('" & f2 & "')"
   End If
 End If

 
 rs2.Open q, cn1
 cnn_nube.BeginTrans
 While Not rs2.EOF
       
       F = Year(rs2("fecha")) & "-" & Month(rs2("fecha")) & "-" & Day(rs2("fecha"))
       
       If rs2("Contado") = "S" Then
           u = "N"
       Else
           u = rs2("cta_cte")
       End If
           
           
       Select Case rs2("id_tipocomp")
       Case Is = 1
          a = "Fc"
       Case Is = 2
           a = "Nd"
       Case Is = 3
           a = "Nd"
       Case Is = 50
           a = "Rb"
       Case Else
           a = "Xx"
       End Select
           
           
       c = a & " " & rs2("letra") & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comp"), "00000000")
              
              
       If rs2("moneda") = "P" Then
            ip = rs2("total")
            id = rs2("total_otra_moneda")
       Else
            ip = rs2("total_otra_moneda")
            id = rs2("total")
       End If
       q = ""
       q = "INSERT INTO ventas(id_sistema, fecha, importe, ubicacion, comprobante, id_cliente , cliente, num_int, total_dolares, moneda, cotizacion)"
       q = q & " VALUES (" & para.idsistema & ",'" & F & "', " & ip & ", '" & u & "', '" & c & "', " & rs2("id_cliente") & ", '" & rs2("cliente02") & "', " & rs2("num_int") & ", " & id & ", '" & rs2("moneda") & "', " & rs2("cotizacion_dolar") & ")"
       cnn_nube.Execute q
 
 
       espere.Label1 = m & "   " & rs2("num_int")
       espere.Refresh
 
 
       rs2.MoveNext
 Wend
 cnn_nube.CommitTrans
 
 
 
 
 m = "Paso 8/8 --> Subiendo informacion...COMPRAS.Movimientos"
 espere.Label1 = m
 espere.Refresh
 Set rs2 = New ADODB.Recordset
 q = "select * from A5"
 If f2 <> "" Then
   If IsDate(f2) Then
       q = q & " where datevalue([fecha]) > datevalue('" & f2 & "')"
   End If
 End If

 
 rs2.Open q, cn1
 cnn_nube.BeginTrans
 While Not rs2.EOF
       
       F = Year(rs2("fecha")) & "-" & Month(rs2("fecha")) & "-" & Day(rs2("fecha"))
       
       If rs2("Contado") = "S" Then
           u = "N"
       Else
           u = rs2("ctacte")
       End If
           
           
       Select Case rs2("id_tipocomp")
       Case 1 To 9
          a = "Fc"
       Case 20 To 29
           a = "Nd"
       Case 30 To 39
           a = "Nd"
       Case Is = 50
           a = "op"
       Case 90 To 99
           a = "Ret"
       Case Else
           a = "Xx"
       End Select
           
           
       c = a & " " & rs2("letra") & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
              
       q = ""
       q = "INSERT INTO compras(id_sistema, fecha, importe, ubicacion, comprobante, id_proveedor , proveedor, num_int)"
       q = q & " VALUES (" & para.idsistema & ",'" & F & "', " & rs2("total") & ", '" & u & "', '" & c & "', " & rs2("id_proveedor") & ", '" & rs2("proveedor05") & "', " & rs2("num_int") & ")"
       cnn_nube.Execute q
 
 
       espere.Label1 = m & "   " & rs2("num_int")
       espere.Refresh
 
 
       rs2.MoveNext
 Wend
 
 cnn_nube.CommitTrans
 
 cn1.BeginTrans
 q = "update g0 set ult_sinc_nube ='" & Now & "' where sucursal= 0"
 cn1.Execute q
 cn1.CommitTrans
 
 
 MsgBox ("Proceso realizado exitosamente!!!")
 Unload espere
 
 
 Exit Sub
 
err1:
 MsgBox ("Error al abrir Base de Datos remota. Intente mas tarde")
 Unload espere
 Exit Sub
 
err2:
 MsgBox ("Error durante el precesamiento, los datos pueden tener inconsistencias intentelo nuevamente")
 Unload espere
 Exit Sub
 
 

End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Command3_Click()
Set cnn_nube = New ADODB.Connection
  
 
 ' ****OJO ***** cambiar en load  Label4 = "Modo LOCALHOST ACTIVADO" 0 Label4 = "Modo SERVER ACTIVADO CUANDO SE CMABI ESTA LINEA"
 cnn_nube.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=5a04;connection="
 'cnn_nube.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=159.65.106.144;uid=aero;pwd=s7VXRGBZ27E=;database=5a04;connection="
 
 
 
 cnn_nube.Open
    
 'On Error GoTo err2
 
 
 espere.Label1 = "Paso 2/8 --> Borrando elementos Base de Datos remota-Ventas:Movimientos"
 espere.Refresh
 Set rs = New ADODB.Recordset
 q = "delete from ventas where id_sistema = " & para.idsistema
 f2 = ""
 f2i = ""
 f2 = t_dia & "-" & T_mes & "-" & t_año
 f2i = t_año & "-" & T_mes & "-" & t_dia
 MsgBox (f2i)
 If IsDate(f2i) Then
       q = q & " and fecha >= '" & f2i & "'"
 End If
 rs.Open q, cnn_nube, adOpenDynamic, adLockOptimistic
 Set rs = Nothing

 
 
End Sub

Private Sub Form_Load()

Label4 = "Modo SERVER Activado"
'Label4 = "Modo LOCALHOST Activado"


Set rs = New ADODB.Recordset
q = "select ult_sinc_nube from g0 where sucursal = 0"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  Label1 = "Última Fecha de sincronización " & rs("ult_sinc_nube")
Else
  Label1 = "Error en archivo de parametros"
  End
End If
Set rs = Nothing
End Sub
