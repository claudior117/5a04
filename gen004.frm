VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_migrardatos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIGRAR DATOS EXMA"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5580
   ScaleWidth      =   9405
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command11 
      Caption         =   "Access de terceros"
      Height          =   495
      Left            =   240
      TabIndex        =   32
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton Command10 
         Caption         =   "wemec2"
         Height          =   495
         Left            =   5640
         TabIndex        =   31
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Genera clientes"
         Height          =   495
         Left            =   2520
         TabIndex        =   30
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Desde Access"
         Height          =   495
         Left            =   5160
         TabIndex        =   29
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Desde Excel"
         Height          =   495
         Left            =   5160
         TabIndex        =   28
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Iva"
         Height          =   495
         Left            =   5160
         TabIndex        =   27
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cuit"
         Height          =   495
         Left            =   5160
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Depurar Productos"
         Height          =   495
         Left            =   3480
         TabIndex        =   25
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Genera cod. prod"
         Height          =   495
         Left            =   1560
         TabIndex        =   24
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Corregir Varios"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Corregir Precios"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Corregir Ventas"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Plan de Cuentas"
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Empleados"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Ch.terc."
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Bancos"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Corregir Compras"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Stock"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Version del sistema a Migrar"
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   2760
         Width           =   2895
         Begin VB.OptionButton Option2 
            Caption         =   "Exmapvf"
            Height          =   195
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Empresas"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Mov. Compras"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Mov. Ventas"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Migrar"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Proveedores"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Productos/Grupos/Dto"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Clientes/Vendedores"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5325
      Width           =   9405
      _ExtentX        =   16589
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
            TextSave        =   "10/04/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:28 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6720
      TabIndex        =   15
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"gen004.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   6480
      TabIndex        =   10
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"gen004.frx":00E7
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   6480
      TabIndex        =   9
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "gen_migrardatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim cncac As ADODB.Connection
Dim cncap As ADODB.Connection
Dim cnstk As ADODB.Connection
Dim cnf As ADODB.Connection
Dim cnf2 As ADODB.Connection
Dim cnfb As ADODB.Connection
Dim cne As ADODB.Connection
Dim cncgr As ADODB.Connection

Sub corrigevarios()
'Set rs = New ADODB.Recordset
'q = "select * from vta_03 where [id_tasaiva] = 3"
'rs.Open q, cn1, adOpenDynamic, adLockOptimistic
'While Not rs.EOF
'  rs("id_tasaiva") = 1
'  rs("tasaiva") = 21
'  rs.Update
'  rs.MoveNext
'Wend


Set rs = New ADODB.Recordset
q = "select * from a2 "
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
 If IsNull(rs("costoreal")) Then
  rs("costoreal") = rs("precio_ult_compra")
  rs.Update
 End If
  rs.MoveNext
Wend
Set rs = Nothing

MsgBox ("proceso terminado")
End Sub

Private Sub Command1_Click()

J = InputBox$("Ingrese Clave de Administrador General")
If J = "0969" Then
  'On Error GoTo errtemp
  Set cnf = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\cac.mdb;User id=claudio" & ";password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
  cnf.Open gconexion
  
  Set cnf2 = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\cap.mdb;User id=claudio" & ";password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
  cnf2.Open gconexion
   
  Set cnstk = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\stk.mdb;User id=claudio" & ";password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
  cnstk.Open gconexion
  
  
  u = 0
     
   If Check1 = 1 Then
     
      
      'clientes
     Set rs = New ADODB.Recordset
     q = "select * from a1" 'exma
     rs.Open q, cnf
     
     
     'los copdigos de vendedor y proveedor contienen los viejos
     'despues cuando se migre vendedores y proveedores hay que actualizarlos
     u = 0
     While Not rs.EOF
       Label3 = u
       Label3.Refresh
       Set rs1 = New ADODB.Recordset
       q = "select * from  vta_01 where [id_cliente] = " & rs("cod-cliente") '5a
       rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs1.EOF And Not rs1.BOF Then
          
       Else
         rs1.AddNew
       End If
         rs1("Denominacion") = rs("denominacion") & " "
         rs1("Direccion") = rs("direccion") & " "
         rs1("cp") = rs("cod-postal") & " "
         rs1("provincia") = "Buenos Aires"
         rs1("localidad") = rs("localidad") & " "
         rs1("te") = rs("te1") & " / " & rs("te2")
        If Not IsNull(rs("cuit")) Then
         If Len(rs("cuit")) <> 13 Then
          If Len(rs("cuit")) = 11 Then
           cc = Mid$(rs("cuit"), 3, 8) & "-" & Mid$(rs("cuit"), 11, 1)
          Else
             cc = Mid$(rs("cuit"), 3, 7) & "-" & Mid$(rs("cuit"), 10, 1)
          End If
          nc = Mid$(rs("cuit"), 1, 2) & "-" & cc
         Else
          nc = rs("cuit")
         End If
        Else
          nc = "0"
        End If
        rs1("cuit") = nc
        rs1("email") = " "
        rs1("id_proveedor") = 1
        rs1("limite_credito") = 99999999.99
        rs1("exportacion") = "N"
        rs1("id_cliente_anterior") = rs("cod-cliente")
        rs1("id_vendedor") = 1
        Select Case rs("cod-tipoiva")
        Case Is = 1, Is = 2, Is = 3
          ti = rs("cod-tipoiva")
        Case Is = 4
          ti = 5
        Case Is = 6
          ti = 4
        Case Is = 7
          ti = 6
        End Select
             
        rs1("id_tipoiva") = ti
        rs1("Observaciones") = " "
        rs1("inscripto_operador_granos") = "N"
        rs1("percive_ib") = "N"
        rs1("saldo_incobrable") = "N"
        rs1("id_prov") = 2 'provincia ba
        rs1("direccion_local") = rs("direccion")
        
        
        rs1.Update
        Set rs1 = Nothing
     
      rs.MoveNext
      u = u + 1
     Wend
 End If
   
   
   
   
 If Check3 = 1 Then 'proveedores
       
     Set rs = New ADODB.Recordset
     q = "select * from a19" 'exma
     rs.Open q, cnf2
     
     
     'los copdigos de vendedor y proveedor contienen los viejos
     'despues cuando se migre vendedores y proveedores hay que actualizarlos
     u = 0
     While Not rs.EOF
      Label3 = u
      Label3.Refresh
      Set rs1 = New ADODB.Recordset
      q = "select * from  a1 where [id_proveedor] = " & rs("cod-proveedor") '5a
      rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rs1.EOF And Not rs.BOF Then
        'rs1.Edit
      Else
       rs1.AddNew
      End If
         rs1("Denominacion") = rs("denominacion") & " "
         rs1("Direccion") = rs("direccion") & " "
         rs1("cp") = rs("cod-postal") & " "
         rs1("provincia") = "BA"
         rs1("localidad") = rs("localidad") & " "
         rs1("te") = rs("te1") & " / " & rs("te2")
         rs1("cuit") = Left$(rs("cuit"), 13)
         rs1("email") = rs("mail") & " "
         rs1("id_codretgan") = 0
         rs1("inscripto_gan") = "S"
         rs1("num_ib") = rs("cuit")
         rs1("fecha_vto_exepcion_ib") = Format$(Now, "dd/mm/yyyy")
         rs1("id_tipoib") = 0
         rs1("id_prov_anterior") = rs("cod-proveedor")
         rs1("id_codretib") = 0
         rs1("contacto") = "*"
         rs1("te_contacto") = "*"
         rs1("email_contacto") = "*"
         Select Case rs("cod-tipoiva")
         Case Is = 1, Is = 2, Is = 3
           rs1("cod_tipoiva") = rs("cod-tipoiva")
         Case Is = 7
           rs1("cod_tipoiva") = 4
         Case Else
           rs1("cod_tipoiva") = 3
         End Select
         rs1("transporte") = "N"
         rs1("id_provincia") = 2
         rs1("id_cuenta_a1") = 210101
         
        
        rs1.Update
        Set rs1 = Nothing
 
     
      rs.MoveNext
      u = u + 1
     Wend
   
     Set rs = Nothing
          
   End If
 
 


If Check2 = 1 Then
  'Call productos   'solo marcas / deptos y grupos
  Call productos2  'importa productos mantenienfo codifos que ya deben estar generados con genera dodigos(antes ejecutado productos)
  'Call productos3 'productos con codigos nuevos (antes ejecutado productos)
  'Call productos4 'productos con codigos nuevos (antes ejecutado productos) comparando codigos de barra

End If

 If Check7 = 1 Then 'movimietos de venta
    Call movventas
 
 End If

If Check4 = 1 Then 'stock
    Call STOCK
End If
 
 
If Check8 = 1 Then
  Call compras
End If

If Check6 = 1 Then
  Call bancos
End If

If Check9 = 1 Then
  Call bancos2
End If

If Check10 = 1 Then
  Call empleados
End If

If Check11 = 1 Then
  Call plan
End If

If Check5 = 1 Then
  Call corrigecompras
End If

If Check12 = 1 Then
  Call corrigeventas
End If


If Check13 = 1 Then
  Call corrigeprecios
End If


If Check14 = 1 Then
  Call corrigevarios
End If

End If



Exit Sub


errtemp:
   MsgBox ("Error al Abrir Base de Datos a Importar")
   Exit Sub
End Sub

Sub productos2()
  'migro productos para mantener los codigos
  'los codigos ya deben estar generados y las marcas, grupos, etc importados
  'usar precio-unitario porc-utilidad y precio-final para tomar precios publico
  'usar precio-unitario2 porc-utilidad2 y precio-final2 para tomar precios mayorista
  
   'productos
  Set rs = New ADODB.Recordset
  q = "select * from a2" 'exma
  rs.Open q, cnf
     
  While Not rs.EOF
    Set rs1 = New ADODB.Recordset
    q = "select * from  a2 where [id_producto] = " & rs("basico")
    rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF And Not rs1.BOF Then
       
         rs1("Descripcion") = rs("descripcion") & " "
         
         Set rs2 = New ADODB.Recordset
         q = "select * from a8 where [id_anterior] = " & rs("tipo-producto")
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cg = rs2("id_grupo")
         Else
           cg = 1
         End If
         Set rs2 = Nothing
         
         rs1("id_grupo") = cg
         
         q = "select * from a9 where [id_anterior] = " & rs("cod-depto")
         Set rs2 = New ADODB.Recordset
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           CDep = rs2("id_departamento")
         Else
           CDep = 1
         End If
         Set rs2 = Nothing
         rs1("id_departamento") = CDep
         
         q = "select * from a10 where [id_anterior] = " & rs("cod-marca")
         Set rs2 = New ADODB.Recordset
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cmar = rs2("id_marca")
         Else
           cmar = 1
         End If
         Set rs2 = Nothing
         rs1("id_marca") = cmar
         
         
         q = "select * from a1 where [id_prov_anterior] = " & rs("cod-proveedor")
         Set rs2 = New ADODB.Recordset
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cprov = rs2("id_proveedor")
         Else
           cprov = 1
         End If
         Set rs2 = Nothing
         rs1("id_proveedor") = cprov
         rs1("precio_ult_compra") = rs("precio-compra")
         rs1("fecha_ult_compra") = Format$(rs("fecha-compra"), "dd/mm/yyyy")
         rs1("id_proveedor_ult_compra") = cprov
         rs1("pu") = rs("precio-unitario")
         Select Case rs("cod-tasaiva")
          Case Is = 1
            ti = 1
          Case Is = 2
            ti = 2
          Case Is = 0
            ti = 3
          Case Else
            ti = 1
         End Select
         rs1("cod_tasaiva") = ti
         rs1("id_unidad") = 7
         rs1("envase") = 1
         rs1("stock") = rs("stock-venta")
         rs1("requeridos") = 0
         rs1("pedidos") = 0
         rs1("stock_minimo") = rs("stock-minimo")
         rs1("porc_utilidad") = rs("porc-utilidad")
         rs1("costoreal") = rs("costoreal")
         rs1("flete_compra") = rs("flete-compra")
         rs1("dto_compra") = rs("dto-compra")
         rs1("cod_barra") = rs("cod-producto")
         rs1("precio_final") = Format(rs("precio-final"), "#####0.00")
         rs1("tasa_imp_interno") = rs("tasa-imp-interno")
         rs1("tipo_producto") = "P"
         rs1("moneda") = rs("moneda")
         rs1("impuesto") = rs("impuesto")
         rs1("observaciones") = "*"
         rs1("ultima_compra") = rs("fecha-compra") & "  " & rs("precio-compra")
         rs1("ultima_venta") = " "
         rs1("fecha_actu_precio_venta") = Format$(rs("fecha"), "dd/mm/yyyy")
         rs1("id_anterior") = rs("basico")
         rs1("emite_etiqueta") = "N"
         rs1("texto_central") = rs("texto-central-etiqueta") & " "
         If rs("vigente") = "S" Then
           rs1("vigente") = True
         Else
           rs1("vigente") = False
         End If
         rs1("reg_faltante") = 0
         rs1("tipo_carga_tique") = rs("tipo-carga-ticket")
         rs1("abreviatura") = rs("abreviatura")
         rs1("id_tasaib") = 1
         rs1("id_prod_prov") = 0
         rs1("dto_compra2") = 0
       rs1.Update
    Else
      MsgBox ("Codigo no encontrado " & rs("basico"))
    End If
    Set rs12 = Nothing
    rs.MoveNext
  Wend

End Sub

Sub productos()
  'migro productos
  'a20 marcas
  'A22 departamentos
  'a11 Grupos migro
   
   
  Set rs = New ADODB.Recordset
  q = "select * from a22" 'exma
  rs.Open q, cnf
     
  Set rs1 = New ADODB.Recordset
  q = "select * from  a9" 'departamentos
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
       rs1.AddNew
         rs1("Descripcion") = rs("depto") & " "
         rs1("id_anterior") = rs("cod-depto")
       rs1.Update
       rs.MoveNext
  Wend
  Set rs = Nothing
  
 Set rs = New ADODB.Recordset
 q = "select * from a20" 'marcas exma
 rs.Open q, cnf
     
 Set rs1 = New ADODB.Recordset
 q = "select * from  a10" '5a
 rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
 While Not rs.EOF
       rs1.AddNew
        rs1("Descripcion") = rs("marca") & " "
        rs1("id_anterior") = rs("cod-marca")
       rs1("cod_barra") = 0
       rs1.Update
       rs.MoveNext
  Wend
  Set rs = Nothing
  
  
   Set rs = New ADODB.Recordset
   q = "select * from a11" 'grupos
   rs.Open q, cnf
  While Not rs.EOF
   Set rs1 = New ADODB.Recordset
   q = "select * from  a8 where [id_grupo] = " & rs("tipo-producto") 'grupos
   rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
   If Not rs1.EOF And Not rs1.BOF Then
       'edito
        rs1("Descripcion") = rs("descripcion") & " "
        rs1("id_anterior") = rs("tipo-producto")
   
   Else
       'nuevo
       rs1.AddNew
         rs1("Descripcion") = rs("descripcion") & " "
        rs1("id_anterior") = rs("tipo-producto")
   End If
   rs1.Update
   Set rs1 = Nothing
   rs.MoveNext
  Wend
  Set rs = Nothing
  
  
   Set rs = New ADODB.Recordset
   q = "select * from a21" 'grupos
   rs.Open q, cnf
     
  Set rs1 = New ADODB.Recordset
  q = "select * from  a18" 'paises
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
       rs1.AddNew
        rs1("id_pais") = rs("cod-pais")
        rs1("pais") = rs("pais") & " "
        rs1("id_anterior") = rs("cod-pais")
       rs1.Update
       rs.MoveNext
  Wend
  Set rs = Nothing
  
  
  

End Sub
Sub productos4()
  'migro productos generando nuevos codigos, previo hay que ejecutar productos para crear marcas deptos, etc
   'productos
   'en esta verison verifico codigo de barra si no existe en la lista
 ca = 0
 cr = 0
  
  Set rs = New ADODB.Recordset
  q = "select * from a2" 'exma
  rs.Open q, cnf
     
  Set rs1 = New ADODB.Recordset
  q = "select * from A2 "
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
    'primero busco en la lista si no existe el codigo de barra
   'cbexma = rs("cod-producto")
    
   
       rs1.AddNew
         ca = ca + 1
         rs1("Descripcion") = rs("descripcion") & " "
         
         Set rs2 = New ADODB.Recordset
         q = "select * from a8 where [id_anterior] = " & rs("tipo-producto")
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cg = rs2("id_grupo")
         Else
           cg = 1
         End If
         Set rs2 = Nothing
         
         rs1("id_grupo") = cg
         
         q = "select * from a9 where [id_anterior] = " & rs("cod-depto")
         Set rs2 = New ADODB.Recordset
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           CDep = rs2("id_departamento")
         Else
           CDep = 1
         End If
         Set rs2 = Nothing
         rs1("id_departamento") = CDep
         
         q = "select * from a10 where [id_anterior] = " & rs("cod-marca")
         Set rs2 = New ADODB.Recordset
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cmar = rs2("id_marca")
         Else
           cmar = 1
         End If
         Set rs2 = Nothing
         rs1("id_marca") = cmar
         
         
        ' q = "select * from a1 where [id_prov_anterior] = " & rs("cod-proveedor")
        ' Set rs2 = New ADODB.Recordset
        ' rs2.Open q, cn1
        ' If Not rs2.EOF And Not rs2.BOF Then
        '   cprov = rs2("id_proveedor")
        ' Else
           cprov = 1
        ' End If
        ' Set rs2 = Nothing
         rs1("id_proveedor") = cprov
         rs1("precio_ult_compra") = rs("precio-compra")
         rs1("fecha_ult_compra") = Format$(rs("fecha-compra"), "dd/mm/yyyy")
         rs1("id_proveedor_ult_compra") = cprov
         rs1("pu") = rs("precio-unitario")
         Select Case rs("cod-tasaiva")
          Case Is = 1
            ti = 1
          Case Is = 2
            ti = 2
          Case Is = 0
            ti = 3
          Case Else
            ti = 1
         End Select
         rs1("cod_tasaiva") = ti
         rs1("id_unidad") = 7
         rs1("envase") = 1
         rs1("stock") = rs("stock-venta")
         rs1("requeridos") = 0
         rs1("pedidos") = 0
         rs1("stock_minimo") = rs("stock-minimo")
         rs1("porc_utilidad") = rs("porc-utilidad")
         rs1("costoreal") = rs("costoreal")
         rs1("flete_compra") = rs("flete-compra")
         rs1("dto_compra") = rs("dto-compra")
         rs1("cod_barra") = rs("cod-producto")
         rs1("precio_final") = Format(rs("precio-final"), "#####0.00")
         rs1("tasa_imp_interno") = rs("tasa-imp-interno")
         rs1("tipo_producto") = "P"
         rs1("moneda") = rs("moneda")
         rs1("impuesto") = rs("impuesto")
         rs1("observaciones") = "*"
         rs1("ultima_compra") = rs("fecha-compra") & "  " & rs("precio-compra")
         rs1("ultima_venta") = " "
         rs1("fecha_actu_precio_venta") = Format$(rs("fecha"), "dd/mm/yyyy")
         rs1("id_anterior") = rs("basico")
         rs1("emite_etiqueta") = "N"
         rs1("texto_central") = rs("texto-central-etiqueta") & " "
         If rs("vigente") = "S" Then
           rs1("vigente") = True
         Else
           rs1("vigente") = False
         End If
         rs1("reg_faltante") = 0
         rs1("tipo_carga_tique") = rs("tipo-carga-ticket")
         rs1("abreviatura") = rs("abreviatura")
         rs1("id_tasaib") = 1
         rs1("id_prod_prov") = 1
         rs1("dto_compra2") = 0
       rs1.Update
      
     
     rs.MoveNext
  Wend
  MsgBox ("Terminado --> Agregados: " & ca & "  Repetidas: " & cr)
End Sub
Sub corrigeprecios()
  'productos
  Set rs = New ADODB.Recordset
  q = "select * from a2 where [descuento] <> 0 " 'exma
  rs.Open q, cnf
  While Not rs.EOF
    Set rs1 = New ADODB.Recordset
    q = "select * from  a2 where [id_anterior] = " & rs("cod-producto")
    rs1.MaxRecords = 1
    rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF And Not rs1.BOF Then
         pu = Format(rs("precio-unitario") - (rs("precio-unitario") * (rs("DESCUENTO") / 100)), "#####0.00")
         pf = Format(pu * 1.21, "#####0.00")
         rs1("pu") = pu
         rs1("precio_final") = pf
         rs1.Update
    End If
    Set rs1 = Nothing
    rs.MoveNext
  Wend
   
End Sub

Sub productos3()
  'migro productos generando nuevos codigos, previo hay que ejecutar productos para crear marcas deptos, etc
   'productos
  Set rs = New ADODB.Recordset
  q = "select * from a2" 'exma
  rs.Open q, cnf
     
  Set rs1 = New ADODB.Recordset
  q = "select * from  a2" '5a
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
       rs1.AddNew
         rs1("Descripcion") = rs("descripcion") & " "
         
         Set rs2 = New ADODB.Recordset
         q = "select * from a8 where [id_anterior] = " & rs("tipo-producto")
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cg = rs2("id_grupo")
         Else
           cg = 1
         End If
         Set rs2 = Nothing
         
         rs1("id_grupo") = cg
         
         q = "select * from a9 where [id_anterior] = " & rs("cod-depto")
         Set rs2 = New ADODB.Recordset
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           CDep = rs2("id_departamento")
         Else
           CDep = 1
         End If
         Set rs2 = Nothing
         rs1("id_departamento") = CDep
         
         q = "select * from a10 where [id_anterior] = " & rs("cod-marca")
         Set rs2 = New ADODB.Recordset
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cmar = rs2("id_marca")
         Else
           cmar = 1
         End If
         Set rs2 = Nothing
         rs1("id_marca") = cmar
         
         
         q = "select * from a1 where [id_prov_anterior] = " & rs("cod-proveedor")
         Set rs2 = New ADODB.Recordset
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cprov = rs2("id_proveedor")
         Else
           cprov = 1
         End If
         Set rs2 = Nothing
         rs1("id_proveedor") = cprov
         rs1("precio_ult_compra") = rs("precio-compra")
         rs1("fecha_ult_compra") = Format$(rs("fecha-compra"), "dd/mm/yyyy")
         rs1("id_proveedor_ult_compra") = cprov
         rs1("pu") = rs("precio-unitario")
         Select Case rs("cod-tasaiva")
          Case Is = 1
            ti = 1
          Case Is = 2
            ti = 2
          Case Is = 0
            ti = 3
          Case Else
            ti = 1
         End Select
         rs1("cod_tasaiva") = ti
         rs1("id_unidad") = 7
         rs1("envase") = 1
         rs1("stock") = rs("stock-venta")
         rs1("requeridos") = 0
         rs1("pedidos") = 0
         rs1("stock_minimo") = rs("stock-minimo")
         rs1("porc_utilidad") = rs("porc-utilidad")
         rs1("costoreal") = rs("costoreal")
         rs1("flete_compra") = rs("flete-compra")
         rs1("dto_compra") = rs("dto-compra")
         rs1("cod_barra") = rs("cod-producto")
         rs1("precio_final") = Format(rs("precio-final"), "#####0.00")
         rs1("tasa_imp_interno") = rs("tasa-imp-interno")
         rs1("tipo_producto") = "P"
         rs1("moneda") = rs("moneda")
         rs1("impuesto") = rs("impuesto")
         rs1("observaciones") = "*"
         rs1("ultima_compra") = rs("fecha-compra") & "  " & rs("precio-compra")
         rs1("ultima_venta") = " "
         rs1("fecha_actu_precio_venta") = Format$(rs("fecha"), "dd/mm/yyyy")
         rs1("id_anterior") = rs("basico")
         rs1("emite_etiqueta") = "N"
         rs1("texto_central") = rs("texto-central-etiqueta") & " "
         If rs("vigente") = "S" Then
           rs1("vigente") = True
         Else
           rs1("vigente") = False
         End If
         rs1("reg_faltante") = 0
         rs1("tipo_carga_tique") = rs("tipo-carga-ticket")
         rs1("abreviatura") = rs("abreviatura")
         
       rs1.Update
       rs.MoveNext
  Wend

End Sub

Sub corrigeventas()
'Set rs = New ADODB.Recordset
'q = "select * from vta_02 where [estado] = 'B' "
'rs.Open q, cn1, adOpenDynamic, adLockOptimistic
'While Not rs.EOF
'  Set rs2 = New ADODB.Recordset
'  q = "select * from vta_03 where [num_int] = " & rs("num_int")
'  rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
'  While Not rs2.EOF
'     rs2.Delete
'     rs2.MoveNext
'  Wend
'  Set rs2 = Nothing
'  rs.Delete
'  rs.MoveNext
'Wend
'Set rs = Nothing

'Set rs = New ADODB.Recordset
'q = "select * from vta_02 where [sucursal_ingreso] = 0 "
'rs.Open q, cn1, adOpenDynamic, adLockOptimistic
'While Not rs.EOF
'    rs("sucursal_ingreso") = 1
'    rs("sucursal") = 1
'    rs.Update
'    rs.MoveNext
'Wend
'Set rs = Nothing
    
    
 Set rs = New ADODB.Recordset
q = "select * from vta_03  "
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    rs("Unidad") = "U."
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
    
End Sub
Sub plan()
' On Error GoTo err2


ns = 0
u = 0
Set rs1 = New ADODB.Recordset
q = "select * from c_01"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
'borro plan
While Not rs1.EOF
 Label3 = "1 - " & u
 Label3.Refresh
 rs1.Delete
 rs1.MoveNext
  u = u + 1
Wend

Set rs = New ADODB.Recordset
q = "select * from a3"
rs.Open q, cncgr

ns = 0
u = 0
'migro pos1
While Not rs.EOF
   Label3 = "2 - " & u
   Label3.Refresh
   rs1.AddNew
   rs1("id_cuenta") = Val(Format$(rs("pos1"), "0") & "00000")
   rs1("pos1") = rs("pos1")
   rs1("pos2") = 0
   rs1("pos3") = 0
   rs1("pos4") = 0
   rs1("pos5") = 0
   rs1("descripcion") = rs("descripcion")
   rs1("tipo") = "T"
   rs1.Update
   rs.MoveNext
  u = u + 1
Wend
Set rs = Nothing


Set rs = New ADODB.Recordset
q = "select * from a4"
rs.Open q, cncgr

u = 0
While Not rs.EOF
   Label3 = "3 - " & u
   Label3.Refresh
   rs1.AddNew
   rs1("id_cuenta") = Val(Format$(rs("pos1"), "0") & Format$(rs("pos2"), "0") & "0000")
   rs1("pos1") = rs("pos1")
   rs1("pos2") = rs("pos2")
   rs1("pos3") = 0
   rs1("pos4") = 0
   rs1("pos5") = 0
   rs1("descripcion") = rs("descripcion")
   rs1("tipo") = "T"
   rs1.Update
  
  rs.MoveNext
  u = u + 1
Wend
Set rs = Nothing


Set rs = New ADODB.Recordset
q = "select * from a5"
rs.Open q, cncgr

u = 0
While Not rs.EOF
   Label3 = "4 - " & u
   Label3.Refresh
   rs1.AddNew
   rs1("id_cuenta") = Val(Format$(rs("pos1"), "0") & Format$(rs("pos2"), "0") & Format$(rs("pos3"), "00") & "00")
   rs1("pos1") = rs("pos1")
   rs1("pos2") = rs("pos2")
   rs1("pos3") = rs("pos3")
   rs1("pos4") = 0
   rs1("pos5") = 0
   rs1("descripcion") = rs("descripcion")
   rs1("tipo") = "T"
   rs1.Update
  
  rs.MoveNext
  u = u + 1
Wend
Set rs = Nothing

Set rs = New ADODB.Recordset
q = "select * from a2"
rs.Open q, cncgr
u = 0
While Not rs.EOF
   Label3 = "5 - " & u
   Label3.Refresh
   rs1.AddNew
   rs1("id_cuenta") = rs("cod-cuenta")
   rs1("pos1") = rs("pos1")
   rs1("pos2") = rs("pos2")
   rs1("pos3") = rs("pos3")
   rs1("pos4") = rs("pos4")
   rs1("pos5") = 0
   rs1("descripcion") = rs("descripcion")
   rs1("tipo") = "C"
   rs1.Update
  
  rs.MoveNext
  u = u + 1
Wend
Set rs = Nothing

  

End Sub
Sub corrigecompras()
 'On Error GoTo err3
     
    q = "select * from a5 "
    Set rs1 = New ADODB.Recordset
    rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
    While Not rs1.EOF
       rs1("fecha_vto") = rs1("fecha")
       rs1.Update
       rs1.MoveNext
    Wend
    Set rs1 = Nothing
    MsgBox ("proceso terminado")
    
End Sub
Sub empleados()
' On Error GoTo err2

Set rs = New ADODB.Recordset
q = "select * from a2"
rs.Open q, cne

Set rs1 = New ADODB.Recordset
q = "select * from emp_02"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
ns = 0
u = 0
While Not rs.EOF
  If rs("importe") > 0 Then
   Label3 = "2 - " & u
   Label3.Refresh
   rs1.AddNew
   rs1("num_mov_int") = rs("num-movimiento")
   rs1("id_legajo") = rs("legajo")
   rs1("importe") = rs("importe")
   rs1("fecha") = rs("fecha")
   rs1("tipo_movimiento") = rs("tipo-movimiento")
   If rs("tipo-movimiento") = 1 Then
      ub = "D"
   Else
      ub = "H"
   End If
   rs1("ubicacion") = ub
   rs1("observaciones") = rs("descripcion") & " "
   rs1.Update
  End If
  rs.MoveNext
  u = u + 1
Wend
Set rs = Nothing


   
   

End Sub

Sub compras()
 'On Error GoTo err3
 Set rs = New ADODB.Recordset
 q = "select * from a20"
 rs.Open q, cnf2
  
 q = "select * from a5 "
 Set rs1 = New ADODB.Recordset
 rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1


ns = 0
u = 0
While Not rs.EOF
   Label3 = "1 - " & u
   Label3.Refresh
 
 
   If rs("num-mov-stk") > ns Then
     ns = rs("num-mov-stk")
   End If
 
 
 rs1.AddNew
 'If rs("cotiz-dolar") > 1 Then
 '     cot = rs("cotizacion")
 '   Else
 '     cot = 1
 'End If
 'If rs("moneda") = "P" Then
 '       rs1("total_d") = Format(rs("importe") / cot, "#######0.00")
 '       rs1("cotiz_dolar") = cot
 'Else
 '       rs1("total_d") = rs("total") * cot
 '       rs1("cotiz_dolar") = cot
 'End If
     
 Set rs2 = New ADODB.Recordset
 q = "select * from a1 where [id_prov_anterior] = " & rs("cod-proveedor")
 rs2.Open q, cn1
 If Not rs2.EOF And Not rs2.BOF Then
        cv = rs2("id_proveedor")
        ctdo = "N"
 Else
        cv = 1
        ctdo = "S"
 End If
 Set rs2 = Nothing
 
 Select Case rs("tipo-comprobante")
 Case Is = 1, Is = 2, Is = 3
   tc = 1
   ps = "E"
   pc = "H"
   po = "E"
   pg = "S"
 Case Is = 20 ', Is = 21
   tc = 20
   ps = "E"
   pc = "H"
   po = "E"
   pg = "S"
 
 Case Is = 21
   'solo para center clean
   tc = 30
   ps = "S"
   pc = "D"
   po = "S"
   pg = "R"
 
 Case Is = 24
   'tc = 24
   'ps = "N"
   'pc = "H"
   'po = "N"
   'pg = "N"
    
    'solo para center clean

    tc = 34
    ps = "N"
    pc = "D"
    po = "N"
    pg = "N"
       
   
 Case Is = 30 ', Is = 31
   tc = 30
   ps = "S"
   pc = "D"
   po = "S"
   pg = "R"
 
 Case Is = 31
  'solo para center clean
   tc = 20
   ps = "E"
   pc = "H"
   po = "E"
   pg = "S"
 
 
 Case Is = 34
   'tc = 34
   'ps = "N"
   'pc = "D"
   'po = "N"
   'pg = "N"
    'solo para center clean
   tc = 24
   ps = "N"
   pc = "H"
   po = "N"
   pg = "N"
  
 Case Is = 40
   tc = 45
   ps = "E"
   pc = "N"
   po = "N"
   pg = "N"
 Case Is = 41
   tc = 46
   ps = "S"
   pc = "N"
   po = "N"
   pg = "N"
 
 Case Is = 51
   tc = 50
   ps = "N"
   pc = "D"
   po = "N"
   pg = "N"

Case Is = 65
   tc = 65
   ps = "N"
   pc = "N"
   po = "N"
   pg = "N"


 End Select
 
 
 rs1("num_int") = rs("num-mov-stk")
 rs1("id_proveedor") = cv
 rs1("sucursal") = rs("sucursal")
 rs1("num_comprobante") = rs("num-comprobante")
 rs1("letra") = rs("cod-comprobante")
 rs1("id_tipocomp") = tc
 rs1("fecha") = rs("fecha")
 rs1("id_usuario") = 1
 rs1("subtotal") = rs("subtotal")
 rs1("iva") = rs("iva")
 rs1("no_grabado") = rs("nograbado")
 rs1("percep_ret") = rs("percepcion") + rs("ib")
 rs1("total") = rs("importe")
 rs1("fecha_prob_entrega") = rs("fecha")
 rs1("fecha_recepcion") = rs("fecha")
 rs1("estado") = rs("estado")
 rs1("id_codretgan") = rs("cod-tipoiva")
 rs1("id_cuenta") = 0
 rs1("stock") = ps
 rs1("ctacte") = pc
 rs1("grabado") = pg
 rs1("estado_pago") = "P" 'rs("estado-pago")
 rs1("num_op") = "0000-00000000" 'rs("num-comp-pago")
 rs1("saldo_impago") = 0
 rs1("id_codretib") = 0
 rs1("ret_gan") = rs("tasa-iva")
 rs1("ret_ib") = 0
 rs1("compra") = po
 If IsNull(rs("detalle")) Then
   rs1("obs") = "*"
 Else
   rs1("obs") = Left$(rs("detalle"), 49) & " "
 End If
 rs1("id_obra") = 0
 rs1("condiciones") = " "
 rs1("info_contacto") = " "
 rs1("contado") = ctdo
 rs1("monto_suj_ret") = 0
 rs1("alicuota_ret") = 0
 rs1("ret_mes") = 0
 rs1("pagos_realizados") = 0
 rs1("pago_actual") = 0
 rs1("minimo_no_imp") = 0
 rs1("total") = rs("importe")
 rs1("total_d") = rs("total-dolar")
 rs1("moneda") = rs("moneda")
 rs1("cotiz_dolar") = rs("cotiz-dolar")
 rs1("fecha_vto") = rs1("fecha")
 rs1.Update
 rs.MoveNext
 u = u + 1
Wend
Set rs = Nothing
Set rs1 = Nothing


 q = "select * from g0"
 Set rs = New ADODB.Recordset
 rs.Open q, cn1, adOpenDynamic, adLockOptimistic
 If Not rs.EOF And Not rs.BOF Then
    rs("ult_num_int_comp") = ns + 100
    rs.Update
 End If
 Set rs = Nothing



'productos en comprobantes
 Set rs = New ADODB.Recordset
 q = "select * from a5 "
 rs.Open q, cn1, adOpenDynamic, adLockOptimistic

Set rs1 = New ADODB.Recordset
q = "select * from a6"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic

u = 0
While Not rs.EOF
 q = "select * from a21 where [num-mov-stk] = " & rs("num_int")
 Set rs2 = New ADODB.Recordset
 rs2.Open q, cnf2
 r = 1
 While Not rs2.EOF
   Label3 = "2 - " & u
   Label3.Refresh
   
   rs1.AddNew
     rs1("num_int") = rs2("num-mov-stk")
     rs1("renglon") = r
     If rs2("cod-producto") > 0 Then
       Set rs3 = New ADODB.Recordset
       q = "select * FROM a2 where [id_anterior] = " & rs2("cod-producto")
       rs3.Open q, cn1
       ip = 1
       If Not rs3.EOF And Not rs3.BOF Then
          rs1("id_producto") = rs3("ID_producto")
          ip = rs3("ID_producto")
          prod = rs3("descripcion")
        Else
         rs1("id_producto") = 1
         prod = "Ingreso manual"
       End If
       Set rs3 = Nothing
     Else
        rs1("id_producto") = 1
        prod = "Ingreso manual"
     End If
     rs1("cantidad") = rs2("cantidad")
     rs1("pu") = rs2("precio-unitario")
     rs1("unidad") = 1
     rs1("detalle") = Left$(prod & " ", 50)
     rs1("importe") = Format(rs2("cantidad") * rs2("precio-unitario"), "#####0.00")
     rs1("ENVASE") = 1
     rs1("tasa_iva") = rs("ret_gan") * 100
     'rs1("impuesto") = 0
     'rs1("costo") = rs("costo")
     rs1("unidad06") = "U."
     rs1("bultos") = 1
     rs1("cantidad_recibida") = 0
     rs1("renglon_requisicion") = 0
     rs1("observaciones") = " "
     rs1("id_obra") = 0
     rs1("id_prod_pedido") = ip
     rs1("fecha") = Format$(Now, "dd/mm/yyyy")
     rs1("id_usuario") = 0
     rs1("num_int_item") = 0
     rs1("estado") = "A"
     rs1("descuento") = 0
     rs1("pusindto") = rs2("precio-unitario")
     rs1("exportacion") = 0
   rs1.Update
   rs2.MoveNext
   u = u + 1
   r = r + 1
  Wend
  rs("ret_gan") = 0
  rs("id_codretgan") = 0
  rs.Update
  rs.MoveNext
Wend
Set rs = Nothing
Set rs1 = Nothing
 


'conceptos de recibos
'Set rs2 = New ADODB.Recordset
'q = "select * from a7"
'rs2.Open q, cn1, adOpenDynamic, adLockOptimistic, 1


'Set rs1 = New ADODB.Recordset
'q = "select * from a5 where [id_tipocomp] = 50"
'rs1.Open q, cn1
'u = 0
'While Not rs1.EOF
  'para cada recibo busco sus ch. terc.
'   Set rs = New ADODB.Recordset
'   q = "select * from caja01 where [num-int-mov-sal] = " & rs1("num_int")
'   rs.Open q, cnf
'   s = 1
'   Label3 = "3 - " & u
'   Label3.Refresh
'   While Not rs.EOF
'      rs2.AddNew
'      rs2("num_int") = rs("num-int-mov-sal")
'      rs2("secuencia") = s
'      If rs("tipo") = "T" Then
'        rs2("id_formapago") = 3
'        tfp = "Ch.Terc."
'      Else
'        rs2("id_formapago") = 50
'        tfp = "Ch.Propio"
'      End If
'      rs2("formapago") = tfp
'      rs2("num_ch") = rs("num-cheque")
'      rs2("fecha_dif") = rs("fecha-dif")
'      rs2("detalle_banco") = rs("banco-suc")
'      rs2("sucursal") = " "
'      rs2("titular") = Left$(rs("titular"), 25)
'      rs2("importe") = rs("importe")
'      rs2("num_int_fp") = 0
'      rs2.Update
'      rs.MoveNext
'      s = s + 1
'   Wend
'   Set rs = Nothing
  
   'otras formas de pago
'   Set rs = New ADODB.Recordset
'   q = "select * from caja02 where [num-mov-int-sal] = " & rs1("num_int")
 '  rs.Open q, cnf
'   While Not rs.EOF
'      rs2.AddNew
'      rs2("num_int") = rs("num-mov-int-sal")
'      rs2("secuencia") = s
'      rs2("id_formapago") = 7
'      rs2("formapago") = "Otras"
'      rs2("num_ch") = 0
'      rs2("fecha_dif") = rs1("fecha")
'      rs2("detalle_banco") = Left$(rs("detalle") & " ", 50)
'      rs2("sucursal") = " "
'      rs2("titular") = " "
 '     rs2("importe") = rs("importe")
 '     rs2("num_int_fp") = 0
'      rs2.Update
'      rs.MoveNext
'      s = s + 1
'   Wend
'   Set rs = Nothing
  
 ' rs1.MoveNext
 ' u = u + 1
'Wend


Exit Sub
err3:
 
End Sub

Sub bancos()
 'On Error GoTo err3
 
 Set rs = New ADODB.Recordset
 q = "select * from a09"
 rs.Open q, cnfb
  
 q = "select * from cyb_07 "
 Set rs1 = New ADODB.Recordset
 rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1


'debitos y Creditos
 
ns = 0
u = 0
While Not rs.EOF
 Label3 = "1 - " & u
 Label3.Refresh
 
 rs1.AddNew
 rs1("descripcion") = Left$(rs("descripcion"), 30)
 rs.MoveNext
 u = u + 1
Wend
Set rs = Nothing


Set rs = New ADODB.Recordset
 q = "select * from a10"
 rs.Open q, cnfb
 
While Not rs.EOF
   Label3 = "1 - " & u
   Label3.Refresh
 
 rs1.AddNew
 rs1("descripcion") = rs("descripcion")
 rs.MoveNext
 u = u + 1
Wend
Set rs = Nothing
Set rs1 = Nothing
 
 
'mov- bancos
 Set rs = New ADODB.Recordset
 q = "select * from A05"
 rs.Open q, cnfb

Set rs1 = New ADODB.Recordset
q = "select * from cyb_04"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1

u = 0
While Not rs.EOF
   Label3 = "2 - " & u
   Label3.Refresh
   
   If rs("cod-banco") = 1 Then
     cb = 52
   Else
     cb = 51
   End If
   rs1.AddNew
     rs1("id_banco") = cb
     rs1("fecha") = rs("fecha")
     rs1("importe") = rs("importe")
     Select Case rs("tipo-movimiento")
     Case Is = 50
       tm = 1
     Case Is = 51
       tm = 50
     Case Is = 80
       tm = 20
     Case Is = 90
       tm = 30
     End Select
     rs1("id_tipomov") = tm
     rs1("fecha_dif") = rs("fecha-dif")
     If rs("ubicacion") = "D" Then
       rs1("ubicacion") = rs("ubicacion")
     Else
       rs1("ubicacion") = "H"
     End If
     rs1("entro") = rs("entro")
     rs1("fecha_acreed") = rs("fecha-acred")
     rs1("num_comp") = rs("num-comprobante")
     rs1("detalle") = rs("detalle")
     rs1("modulo") = "B"
     rs1("num_mov_int") = 0
     rs1("id_tipodbcr") = 1
     rs1("num_mov_banco_ant") = rs("num-mov-banco")
   rs1.Update
   rs.MoveNext
   u = u + 1
Wend
Set rs = Nothing
Set rs1 = Nothing
 

'Ch. Propios
 Set rs = New ADODB.Recordset
 q = "select * from A03"
 rs.Open q, cnf

Set rs1 = New ADODB.Recordset
q = "select * from cyb_02"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1

u = 0
While Not rs.EOF
   Label3 = "2 - " & u
   Label3.Refresh
   
   If rs("cod-banco") = 1 Then
     cb = 52
   Else
     cb = 51
   End If
   
   Set rs2 = New ADODB.Recordset
   q = "select * from cyb_04 where [num_mov_banco_ant] = " & rs("num-mov-banco")
   rs2.Open q, cn1
   If Not rs2.EOF And Not rs2.BOF Then
     nmb = rs2("num_mov_banco")
   Else
     nmb = 0
   End If
   Set rs2 = Nothing
      
   rs1.AddNew
     rs1("id_banco") = cb
     rs1("num_cheque") = rs("num-cheque")
     rs1("fecha_emision") = rs("fecha-emision")
     rs1("fecha_dif") = rs("fecha-acreditacion")
     rs1("estado") = rs("estado")
     rs1("destino") = rs("nombre-salida") & " "
     rs1("importe") = rs("importe")
     rs1("num_mov_banco") = 0
     rs1("id_chequera") = 1
     rs1("num_int_op") = 0
   rs1.Update
   rs.MoveNext
   u = u + 1
Wend
Set rs = Nothing
Set rs1 = Nothing
 



Exit Sub
err3:
 

End Sub
Sub bancos2()
'Ch. Terc
 Set rs = New ADODB.Recordset
 q = "select * from A04"
 rs.Open q, cnfb

Set rs1 = New ADODB.Recordset
q = "select * from cyb_03"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1

u = 0
While Not rs.EOF
   Label3 = "1 - " & u
   Label3.Refresh
   
   If rs("num-mov-banco-i") > 0 Then
     Set rs2 = New ADODB.Recordset
     q = "select * from cyb_04 where [num_mov_banco_ant] = " & rs("num-mov-banco-i")
     rs2.Open q, cn1
     If Not rs2.EOF And Not rs2.BOF Then
      nmbi = rs2("num_mov_banco")
     Else
      nmbi = 0
     End If
     Set rs2 = Nothing
   Else
     nmbi = 0
   End If
   
      
   If rs("num-mov-banco-e") > 0 Then
     Set rs2 = New ADODB.Recordset
     q = "select * from cyb_04 where [num_mov_banco_ant] = " & rs("num-mov-banco-e")
     rs2.Open q, cn1
     If Not rs2.EOF And Not rs2.BOF Then
      nmbe = rs2("num_mov_banco")
     Else
      nmbe = 0
     End If
     Set rs2 = Nothing
   Else
     nmbe = 0
   End If
      
   rs1.AddNew
     rs1("fecha_emision") = rs("fecha-emision")
     rs1("num_cheque") = rs("num-cheque")
     rs1("banco") = rs("banco")
     rs1("sucursal") = rs("sucursal")
     rs1("titular") = rs("titular")
     rs1("importe") = rs("importe")
     If rs("estado") = "E" Then
       e = "J"
     Else
       e = rs("estado")
     End If
     rs1("estado") = e
     rs1("fecha_dif") = rs("fecha-acreditacion")
     rs1("origen") = rs("nombre-entrada")
     rs1("destino") = rs("nombre-salida")
     rs1("num_mov_banco_i") = nmbi
     rs1("num_mov_banco_e") = nmbe
     rs1("num_int_op") = 0
     rs1("num_int_rbo") = 0
     rs1("fecha_salida") = rs("fecha-acreditacion")
     rs1("tipo_salida") = "M"
   rs1.Update
  
   rs.MoveNext
   u = u + 1
Wend
Set rs = Nothing
Set rs1 = Nothing
 

End Sub
Sub STOCK()
'migrar movimientos internos stock
'Set rs = New ADODB.Recordset
'q = "select * from A3 "
'rs.Open q, cnstk

'Set rs1 = New ADODB.Recordset
'q = "select * from stk_02"
'rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
'u = 0
'While Not rs.EOF
'   Label3 = u
'   Label3.Refresh
'   u = u + 1
'   rs1.AddNew
'       rs1("fecha") = rs("fecha")
'       rs1("letra") = "X"
'       rs1("num_comprobante") = rs("num-comprobante")
'       rs1("id_usuario") = 1
'       rs1("detalle") = Left$(rs("linea1"), 50)
'       rs1("sucursal") = 1
'       rs1("tipo_comprobante") = 1
'       rs1("id_proveedor") = 1
'       rs1("id_obra") = 1
'       rs1("num_int_anterior") = rs("num-mov-stk")
'   rs1.Update
   
'   rs.MoveNext
'Wend
'Set rs = Nothing
'Set rs1 = Nothing


'agrego renglones
'Set rs = New ADODB.Recordset
'q = "select * from A4 "
'rs.Open q, cnstk

'Set rs1 = New ADODB.Recordset
'q = "select * from stk_03"
'rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
'u = 0
'While Not rs.EOF
'   Label3 = u
'   Label3.Refresh
'   u = u + 1
'   Set rs2 = New ADODB.Recordset
'   q = "select * from a2 where [id_anterior] = " & rs("cod-producto")
'   rs2.Open q, cn1
'   If Not rs2.EOF And Not rs2.BOF Then
'       cv = rs2("id_producto")
'   Else
'       cv = 1
'   End If
'   Set rs2 = Nothing
   
'   Set rs2 = New ADODB.Recordset
'   q = "select * from stk_02 where [num_int_anterior] = " & rs("num-mov-stk")
'   rs2.Open q, cn1
'   If Not rs2.EOF And Not rs2.BOF Then
'       ni = rs2("num_int")
'   Else
'       ni = 1
'   End If
'   Set rs2 = Nothing
   
   
   
'   rs1.AddNew
'       rs1("num_int") = ni
'       rs1("renglon") = rs("renglon")
'       rs1("id_producto") = cv
'       rs1("descripcion") = rs("descripcion")
'       rs1("ubicacion") = rs("tipo")
'       rs1("cantidad") = rs("cantidad")
'       rs1("detalle") = rs("detalle")
'       rs1("unidad") = "*"
'   rs1.Update
   
'   rs.MoveNext
'Wend
'Set rs = Nothing
'Set rs1 = Nothing





'On Error GoTo err2
Set rs = New ADODB.Recordset
q = "select * from A1 "
rs.Open q, cnstk

Set rs1 = New ADODB.Recordset
q = "select * from stk_01"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
u = 0
While Not rs.EOF
   Label3 = u
   Label3.Refresh
   u = u + 1
   rs1.AddNew
       rs1("fecha") = rs("fecha")
       If Len(rs("cod-producto")) > 5 Then
         Set rs2 = New ADODB.Recordset
         q = "select * from a2 where [cod_barra] = " & rs("cod-producto")
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
             cv = rs2("id_producto")
          Else
              cv = 1
         End If
         Set rs2 = Nothing
       Else
         Set rs2 = New ADODB.Recordset
         q = "select * from a2 where [id_anterior] = " & rs("cod-producto")
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cv = rs2("id_producto")
         Else
           cv = 1
         End If
         Set rs2 = Nothing
       End If
       rs1("id_producto") = cv
       rs1("cantidad") = rs("cantidad")
       rs1("ubicacion") = rs("ubicacion")
       Set rs2 = New ADODB.Recordset
       Select Case rs("origen")
        Case Is = "C"
          'compras
          q = "select * from a20 where [num-mov-stk] = " & rs("num-mov-stk")
          rs2.Open q, cnf2
          If Not rs2.EOF And Not rs2.BOF Then
             comp = rs2("cod-comprobante") & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num-comprobante"), "00000000")
             Desc = rs2("denominacion")
          Else
             comp = "C0000-00000000"
             Desc = "Error"
          End If
          ni = rs("num-mov-stk")
          Set rs2 = Nothing
       Case Is = "V"
          q = "select * from a6 where [num-mov-stk] = " & rs("num-mov-stk")
          rs2.Open q, cnf
          If Not rs2.EOF And Not rs2.BOF Then
             comp = rs2("cod-comprobante") & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num-comprobante"), "00000000")
             Desc = rs2("denominacion")
          Else
             comp = "V0000-00000000"
             Desc = "Error"
          End If
          ni = rs("num-mov-stk")
          Set rs2 = Nothing
       Case Is = "S"
          q = "select * from stk_02 where [num_int_anterior] = " & rs("num-mov-stk")
          rs2.Open q, cn1
          If Not rs2.EOF And Not rs2.BOF Then
             comp = "S" & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
             Desc = "Mov.Int. Stock"
             ni = rs2("num_int")
          Else
             comp = "S0000-00000000"
             Desc = "Error"
             ni = rs("num-mov-stk")
          End If
          Set rs2 = Nothing

       End Select
       rs1("comprobante") = comp
       rs1("Descripcion") = Desc
       rs1("num_mov_int") = ni
       rs1("modulo") = rs("origen")
     rs1.Update
   rs.MoveNext
Wend
Set rs = Nothing
Set rs1 = Nothing
Label3 = "Fin"
Label3.Refresh


End Sub

Sub stock2()
'ventas
q = "select * from vta_02, vta_03, vta_06 where vta_02.[num_int] = vta_03.[num_int] and vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and  [id_producto] > 1 and vta_06.[stock] <> 'N'"
Set rs = New ADODB.Recordset
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
 comp = Left$(rs("abreviatura"), 5) & " " & rs("letra") & Format$(rs("vta_02.sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
 Desc = rs("cliente02")
 QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo])"
 QUERY = QUERY & " VALUES ('" & rs("fecha") & "', " & rs("id_producto") & ", " & rs("cantidad") & ", '" & rs("vta_06.stock") & "', '" & comp & "', '" & Desc & " ', " & rs("vta_02.num_int") & ", 'V')"
 cn1.BeginTrans
  cn1.Execute QUERY
 cn1.CommitTrans
 
 
  rs("vta_02.stock") = rs("vta_06.stock")
 rs.Update
 
 rs.MoveNext
Wend


'COMPRAS
q = "select * from A5, A6, G2, a1 where A5.[num_int] = A6.[num_int] and [id_tipocomp] = [id_tipo_comp] and  [id_producto] > 1 and g2.[stock] <> 'N' and a5.[id_proveedor] = a1.[id_proveedor] "
Set rs = New ADODB.Recordset
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
 comp = Left$(rs("abreviatura"), 5) & " " & rs("letra") & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
 Desc = rs("denominacion")
 QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo])"
 QUERY = QUERY & " VALUES ('" & rs("a5.fecha") & "', " & rs("id_producto") & ", " & rs("cantidad") & ", '" & rs("g2.stock") & "', '" & comp & "', '" & Desc & " ', " & rs("a5.num_int") & ", 'C')"
 cn1.BeginTrans
  cn1.Execute QUERY
 cn1.CommitTrans
 
 
 rs("a5.stock") = rs("g2.stock")
 rs.Update
 
 rs.MoveNext
Wend


End Sub

Private Sub Command10_Click()
  Set rs1 = New ADODB.Recordset
  q = "select * from vta_01 where id_tipoiva = 3"
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
  While Not rs1.EOF
     rs1("cuit") = "0"
     rs1.Update
     rs1.MoveNext
  Wend
End Sub

Private Sub Command11_Click()
J = InputBox$("Ingrese Clave de Administrador General")
If J = "1975" Then
   'Call clientes_terceros_access
   'Call grupos_terceros_access
   ' Call productos_terceros_access
   'Call proveedores_terceros_access
End If
       
End Sub
Sub clientes_terceros_access()
'On Error GoTo errtemp
  Set cnf = New ADODB.Connection
  'bd tercero access
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\stock.mdb;"
  cnf.Open gconexion

  Set rs = New ADODB.Recordset
  q = "select * from clientes" 'tercero
  rs.Open q, cnf
  l = 0
  While Not rs.EOF
     Label3 = u
       Label3.Refresh
       Set rs1 = New ADODB.Recordset
       q = "select * from  vta_01 where [id_cliente] = " & rs("codcli")
       rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs1.EOF And Not rs1.BOF Then
          
       Else
         rs1.AddNew
       End If
         rs1("Denominacion") = rs("nombre") & " "
         rs1("Direccion") = rs("domicilio") & " "
         rs1("cp") = rs("cpostal") & " "
         rs1("provincia") = "*"
         rs1("localidad") = "*"
         rs1("te") = rs("telefonos") & " "
        If Not IsNull(rs("cuit")) Then
         If Len(rs("cuit")) <> 13 Then
          If Len(rs("cuit")) = 11 Then
           cc = Mid$(rs("cuit"), 3, 8) & "-" & Mid$(rs("cuit"), 11, 1)
          Else
             cc = Mid$(rs("cuit"), 3, 7) & "-" & Mid$(rs("cuit"), 10, 1)
          End If
          nc = Mid$(rs("cuit"), 1, 2) & "-" & cc
         Else
          nc = rs("cuit")
         End If
        Else
          nc = "0"
        End If
        rs1("cuit") = nc
        rs1("email") = rs("email") & " "
        rs1("id_proveedor") = 1
        rs1("limite_credito") = 99999999.99
        rs1("exportacion") = "N"
        rs1("id_cliente_anterior") = rs("codcli")
        rs1("id_vendedor") = 1
        Select Case rs("civa")
        Case Is = 1, Is = 2, Is = 3
          ti = rs("civa")
        Case Is = 4
          ti = 5
        Case Is = 6
          ti = 4
        Case Is = 7
          ti = 6
        End Select
             
        rs1("id_tipoiva") = ti
        rs1("Observaciones") = Left$(rs("observacion"), 49) & " "
        rs1("inscripto_operador_granos") = "N"
        rs1("percive_ib") = "N"
        rs1("saldo_incobrable") = "N"
        rs1("id_prov") = 2 'provincia ba
        rs1("direccion_local") = rs("domicilio")
        
        
        rs1.Update
        Set rs1 = Nothing
     
      rs.MoveNext
      u = u + 1
     Wend
  
End Sub

Sub proveedores_terceros_access()
'On Error GoTo errtemp
  Set cnf = New ADODB.Connection
  'bd tercero access
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\stock.mdb;"
  cnf.Open gconexion

  Set rs = New ADODB.Recordset
  q = "select * from proveedores" 'tercero
  rs.Open q, cnf
  l = 0
  While Not rs.EOF
     Label3 = u
       Label3.Refresh
       Set rs1 = New ADODB.Recordset
       q = "select * from  a1 where [id_proveedor] = " & rs("codprov")
       rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs1.EOF And Not rs1.BOF Then
          
       Else
         rs1.AddNew
       End If
         rs1("Denominacion") = rs("nombre") & " "
         rs1("Direccion") = rs("domicilio") & " "
         rs1("cp") = "*"
         rs1("provincia") = "BA"
         rs1("localidad") = "*"
         rs1("te") = rs("telefonos") & " "
        If Not IsNull(rs("cuit")) Then
         If Len(rs("cuit")) <> 13 Then
          If Len(rs("cuit")) = 11 Then
           nc = rs("cuit")
          Else
           nc = "0"
          End If
         Else
          nc = Mid$(rs("cuit"), 1, 2) & Mid$(rs("cuit"), 4, 8) & Mid$(rs("cuit"), 13, 1)

         End If
        Else
          nc = "0"
        End If
        rs1("cuit") = nc
        rs1("email") = rs("email") & " "
        rs1("cod_tipoiva") = 1
        rs1("id_codretgan") = 0
        rs1("inscripto_gan") = "N"
        rs1("id_tipoib") = 0
        rs1("num_ib") = nc
        rs1("fecha_vto_exepcion_ib") = "01/01/2000"
        rs1("id_prov_anterior") = rs("codprov")
        rs1("id_codretib") = 0
        rs1("contacto") = "*"
        rs1("te_contacto") = "*"
        rs1("email_contacto") = "*"
        rs1("alicuota_retib") = 0
        rs1("transporte") = "N"
        rs1("id_provincia") = 2
        rs1("id_cuenta_a1") = 210503
        
        
        
        
        
        rs1.Update
        Set rs1 = Nothing
     
      rs.MoveNext
      u = u + 1
     Wend
  
End Sub



Sub productos_terceros_access()
 Set cnf = New ADODB.Connection
  'bd tercero access
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\stock.mdb;"
  cnf.Open gconexion

  Set rs = New ADODB.Recordset
  q = "select * from productos" 'tercero
  rs.Open q, cnf
  l = 0
  While Not rs.EOF
     Label3 = u
       Label3.Refresh
       Set rs1 = New ADODB.Recordset
       q = "select * from  a2 where [id_producto] = 0"
       rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs1.EOF And Not rs1.BOF Then
          
       Else
         rs1.AddNew
       End If
         
           rs1("Descripcion") = rs("descrip") & " "
         
         Set rs2 = New ADODB.Recordset
         q = "select * from a8 where [id_anterior] = " & rs("rubro")
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           cg = rs2("id_grupo")
         Else
           cg = 1
         End If
         Set rs2 = Nothing
         
         rs1("id_grupo") = cg
         
         CDep = 1
         rs1("id_departamento") = CDep
         
         cmar = 1
         rs1("id_marca") = cmar
         
         
           cprov = 1
         rs1("id_proveedor") = cprov
         rs1("precio_ult_compra") = rs("pucompra")
         rs1("fecha_ult_compra") = Format$("01/01/2022", "dd/mm/yyyy")
         rs1("id_proveedor_ult_compra") = cprov
         rs1("pu") = rs("pventa")
         Select Case 1
          Case Is = 1
            ti = 1
          Case Is = 2
            ti = 2
          Case Is = 0
            ti = 3
          Case Else
            ti = 1
         End Select
         rs1("cod_tasaiva") = ti
         rs1("id_unidad") = 7
         rs1("envase") = 1
         rs1("stock") = 0
         rs1("requeridos") = 0
         rs1("pedidos") = 0
         rs1("stock_minimo") = 0
         rs1("porc_utilidad") = rs("pmargen")
         rs1("costoreal") = rs("pcosto")
         rs1("flete_compra") = 0
         rs1("dto_compra") = 0
         rs1("cod_barra") = rs("codigo")
         rs1("precio_final") = Format(rs("pventafinal"), "#####0.00")
         rs1("tasa_imp_interno") = 0
         rs1("tipo_producto") = "P"
         If rs("divisa") = 0 Then
           rs1("moneda") = "P"
         Else
           rs1("moneda") = "D"
         End If
         rs1("impuesto") = 0
         rs1("observaciones") = "*"
         rs1("ultima_compra") = "*"
         rs1("ultima_venta") = " "
         rs1("fecha_actu_precio_venta") = Format$("01/01/2022", "dd/mm/yyyy")
         rs1("id_anterior") = rs("id_producto")
         rs1("emite_etiqueta") = "N"
         rs1("texto_central") = "*"
           rs1("vigente") = True
         rs1("reg_faltante") = 0
         rs1("tipo_carga_tique") = "A"
         rs1("abreviatura") = "*"
         rs1("id_tasaib") = 1
         rs1("id_prod_prov") = 0
         rs1("dto_compra2") = 0

         
         
         
         
        rs1.Update
        Set rs1 = Nothing
     
      rs.MoveNext
      u = u + 1
     Wend
 
End Sub

Sub grupos_terceros_access()
'On Error GoTo errtemp
  Set cnf = New ADODB.Connection
  'bd tercero access
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\stock.mdb;"
  cnf.Open gconexion

  Set rs = New ADODB.Recordset
  q = "select * from rubros" 'tercero
  rs.Open q, cnf
  l = 0
       
  While Not rs.EOF
     Label3 = u
       Label3.Refresh
       
       Set rs1 = New ADODB.Recordset
       q = "select * from  a8 where [id_grupo] = 0"
       rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs1.EOF And Not rs1.BOF Then
          
       Else
         rs1.AddNew
       End If
         rs1("descripcion") = rs("descrip") & " "
         rs1("id_anterior") = rs("id") & " "
         rs1.Update
        Set rs1 = Nothing
     
      rs.MoveNext
      u = u + 1
     Wend

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Sub movventas()
' On Error GoTo err2
 Set rs = New ADODB.Recordset
 q = "select * from A6 where [estado] <> 'B'"
 rs.Open q, cnf

Set rs1 = New ADODB.Recordset
q = "select * from vta_02"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1

ns = 0
u = 0
nz = 2000
While Not rs.EOF
   If rs("num-mov-stk") > 0 Then
     ns = rs("num-mov-stk")
   Else
     ns = nz + 1
     nz = nz + 1
   End If
   Label3 = "1 - " & u
   Label3.Refresh
       
   rs1.AddNew
     rs1("sucursal") = rs("sucursal")
      rs1("sucursal_ingreso") = rs("sucursal")
     Select Case rs("tipo-comprobante")
     Case Is = 1, Is = 5, Is = 6
        cc = 1
        c = "D"
        g = "S"
        v = "S"
        s = "S"
     Case Is = 8
        cc = 3
        c = "H"
        g = "R"
        v = "R"
        s = "E"

     Case Is = 10
        cc = 300
        c = "D"
        g = "N"
        v = "N"
        s = "N"
     
     Case Is = 20
        cc = 2
        c = "D"
        g = "S"
        v = "N"
        s = "S"
     
     Case Is = 24
        cc = 80
        c = "D"
        g = "N"
        v = "N"
        s = "N"
     Case Is = 34
        cc = 85
        c = "H"
        g = "N"
        v = "N"
        s = "N"
     Case Is = 40
        cc = 40
        c = "D"
        g = "N"
        v = "S"
        s = "S"
     Case Is = 45
        cc = 45
        c = "N"
        g = "N"
        v = "N"
        s = "N"
     Case Is = 48
        cc = 46
        c = "N"
        g = "N"
        v = "N"
        s = "N"
     Case Is = 50
        cc = 50
        c = "H"
        g = "N"
        v = "N"
        s = "N"
   Case Is = 95
        cc = 101
        c = "H"
        g = "R"
        v = "N"
    Case Is = 96
        cc = 100
        c = "H"
        g = "N"
        v = "N"
    Case Is = 97
        cc = 102
        c = "H"
        g = "N"
        v = "N"
     Case Is = 98
        cc = 103
        c = "H"
        g = "N"
        v = "N"
    Case Else
        cc = 0
        c = "N"
        g = "N"
        v = "N"
        s = "N"
    End Select

    Select Case rs("tipoiva")
        Case Is = 1, Is = 2, Is = 3
          ti = rs("tipoiva")
        Case Is = 4
          ti = 5
        Case Is = 6
          ti = 4
        Case Is = 7
          ti = 6
        End Select


     rs1("id_tipocomp") = cc
     rs1("letra") = rs("cod-comprobante")
     rs1("num_comp") = rs("num-comprobante")
     rs1("fecha") = rs("fecha")
     rs1("estado") = rs("estado")
     rs1("num_int") = ns
     rs1("total") = rs("total")
     rs1("iva") = rs("iva1")
     rs1("impuestos") = rs("impuesto")
     rs1("subtotal") = rs("subtotal")
     rs1("descuento") = rs("descuento")
     rs1("id_cuenta") = 0
     If IsNull(rs("linea1")) Then
       rs1("OBServaciones") = " "
     Else
       rs1("observaciones") = rs("linea1") & " "
     End If
     rs1("id_usuario") = 1
     rs1("servicio") = "N"
     rs1("cotizacion_dolar") = rs("dolar")
     rs1("total_otra_moneda") = rs("total-dolar")
     rs1("moneda") = rs("MONEDA")
     rs1("id_vendedor") = 1
     If rs("ctacte") <> "N" Then
       rs1("cta_cte") = c
       rs1("contado") = "N"
     Else
       rs1("cta_cte") = "N"
       rs1("contado") = "S"
     End If
     rs1("stock") = s
     rs1("grabado") = g
     rs1("recibo_pago") = "0000-00000000"
     rs1("fecha_pago") = rs("fecha")
     rs1("id_transporte") = 1
     rs1("alicuota_perc_iva") = 0
     rs1("canje_cereal") = False
     rs1("fecha_vto") = rs("fecha")
     rs1("total_bultos") = 0
     rs1("valor_declarado") = 0
     rs1("transporte") = " "
     rs1("direccion_transp") = " "
     rs1("cuit_transp") = " "
     rs1("perc_iva") = 0
     rs1("perc_ib") = 0
     rs1("perc_gan") = 0
     rs1("perc_ss") = 0
     rs1("venta") = v
     rs1("estado_pago") = rs("estado-pago")
     rs1("id_actividad") = 1
     rs1("alicuota_ib") = 3
     rs1("fecha_vto") = rs("fecha")
     'rs1("remitos") = "0000-00000000"
     rs1("total_bultos") = 0
     rs1("valor_declarado") = 0
     rs1("transporte") = " "
     rs1("direccion_transp") = " "
     rs1("cuit_transp") = " "
     rs1("perc_ss") = 0
     rs1("sucursal_ingreso") = 1
     rs1("cliente02") = rs("denominacion")
     rs1("direccion02") = rs("direccion")
     rs1("cuit02") = rs("cuit")
     rs1("id_tipo_iva02") = ti
     rs1("localidad02") = rs("localidad")
     rs1("chofer02") = "*"
     rs1("dominio02") = "*"
     rs1("dominio_acoplado02") = "*"
     If rs("estado-pago") = "N" Then
       rs1("saldo_impago02") = rs("total")
     Else
       rs1("saldo_impago02") = 0
     End If
     rs1("id_camion02") = 1
     rs1("dni_chofer02") = 0
     rs1("num_z") = 0
     
     
     
     If rs("cod-cliente") <> 99999 Then
       Set rs2 = New ADODB.Recordset
       q = "select * FROM vta_01 where [id_cliente_ANTERIOR] = " & rs("cod-cliente")
       rs2.Open q, cn1
       If Not rs2.EOF And Not rs2.BOF Then
         rs1("id_cliente") = CLng(rs2("ID_cliente"))
       Else
         rs1("id_cliente") = 1
       End If
       Set rs2 = Nothing
     Else
       rs1("id_cliente") = 1
     End If
     rs1.Update
  
  
  
  
  
  
  
  rs.MoveNext
  ni = ni + 1
  u = u + 1
 Wend
 Set rs = Nothing
 Set rs1 = Nothing
 
 q = "select * from g0"
 Set rs = New ADODB.Recordset
 rs.Open q, cn1, adOpenDynamic, adLockOptimistic
 If Not rs.EOF And Not rs.BOF Then
    rs("ult_num_int_vta") = nz + 100
    rs.Update
 End If
 Set rs = Nothing

 
'productos


Set rs1 = New ADODB.Recordset
q = "select * from vta_03"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1

 
 
Set rs2 = New ADODB.Recordset
q = "select * from vta_02"
 
rs2.Open q, cn1

While Not rs2.EOF

 Set rs = New ADODB.Recordset
 q = "select * from A7 where [num-mov-stk] = " & rs2("num_int")
 rs.Open q, cnf
 r = 1
 While Not rs.EOF
   
   rs1.AddNew
     rs1("num_int") = rs2("num_int")
     rs1("renglon") = r
     If rs("basico") > 0 Then
       Set rs3 = New ADODB.Recordset
       q = "select * FROM a2 where [id_anterior] = " & rs("basico")
       rs3.Open q, cn1
       If Not rs3.EOF And Not rs3.BOF Then
          rs1("id_producto") = rs3("id_producto")
          Desc = rs3("descripcion")
       Else
         rs1("id_producto") = 1
         Desc = rs("descripcion")
       End If
     Else
        rs1("id_producto") = 1
        Desc = rs("descripcion")
     End If
     Set rs3 = Nothing
     rs1("cantidad") = rs("cantidad")
     rs1("pu") = rs("precio-unitario")
     rs1("unidad") = 7
     rs1("descripcion") = Left$(rs("descripcion") & " ", 50)
     rs1("importe") = rs("cantidad") * rs("precio-unitario")
     Select Case rs("tasaiva")
     Case Is = 21
       rs1("id_tasaiva") = 1
     Case Is = 10.5
       rs1("id_tasaiva") = 2
     Case Is = 0
       rs1("id_tasaiva") = 3
     Case Else
       rs1("id_tasaiva") = 1
     End Select
     rs1("tasaiva") = rs("tasaiva")
     rs1("impuesto") = 0
     rs1("costo") = rs("costo")
     rs1("tunidad") = rs("unidad") & " "
     rs1("bultos") = 1
     rs1("cantidad_original") = rs("cantidad")
     rs1("pu_final") = Format(rs("precio-unitario") * (1 + (rs("tasaiva") / 100)), "#####0.00")
   rs1.Update
   rs.MoveNext
   u = u + 1
   r = r + 1
 Wend
 Set rs = Nothing
 
 rs2.MoveNext

Wend
Set rs2 = Nothing


'conceptos de recibos
'Set rs2 = New ADODB.Recordset
'q = "select * from vta_04"
'rs2.Open q, cn1, adOpenDynamic, adLockOptimistic

'Set rs1 = New ADODB.Recordset
'q = "select * from vta_02 where [id_tipocomp] = 50"
'rs1.Open q, cn1
'u = 0
'While Not rs1.EOF
  'para cada recibo busco sus ch. terc.
'   Set rs = New ADODB.Recordset
'   q = "select * from a13 where [num-mov-stk] = " & rs1("num_int")
'   rs.Open q, cnf
'   s = 1
'   Label3 = "3 - " & u
'   Label3.Refresh
'   While Not rs.EOF
'     If rs("secuencia") > 0 Then
'      rs2.AddNew
'      rs2("num_int") = rs("num-mov-stk")
'      rs2("secuencia") = s
'      rs2("id_formapago") = 3
'      rs2("formapago") = "Ch.Terc."
'      rs2("num_ch") = rs("num-ch")
'      rs2("fecha_dif") = rs("fecha")
'      rs2("detalle_banco") = rs("banco")
'      rs2("sucursal") = " "
'      rs2("titular") = rs("titular")
'      rs2("importe") = rs("importe")
'      rs2("num_int_fp") = 0
'      rs2.Update
'     Else
'      rs2.AddNew
'      rs2("num_int") = rs("num-mov-stk")
'      rs2("secuencia") = s
'      rs2("id_formapago") = 1
'      rs2("formapago") = "Efectivo"
'      rs2("num_ch") = 0
'      rs2("fecha_dif") = rs("fecha")
'      rs2("detalle_banco") = "Efectivo"
'      rs2("sucursal") = " "
'      rs2("titular") = " "
'      rs2("importe") = rs("importe")
'      rs2("num_int_fp") = 0
'      rs2.Update
'    End If
      
'      rs.MoveNext
'      s = s + 1
'   Wend
'   Set rs = Nothing
  
 ' rs1.MoveNext
 ' u = u + 1
'Wend
 
Exit Sub
err2:
Resume Next

End Sub

Private Sub Command3_Click()
i = InputBox("Ingrese cantidad de codigos consecutivos a generar", , 0)
If Val(i) > 0 Then
    Set rs = New ADODB.Recordset
    q = "select * from a2"
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    For i = 0 To Val(i)
      rs.AddNew
      rs.Update
    Next i
    Set rs = Nothing
    MsgBox ("Proceso terminado")
End If
End Sub

Private Sub Command4_Click()
J = MsgBox("Confirma depurar productos", 4)
If J = 6 Then
  Set rs = New ADODB.Recordset
  q = "select * from a2 "
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
   If IsNull(rs("costoreal")) Then
      rs("costoreal") = 0
   End If
   
   If IsNull(rs("dolar_ult_compra")) Then
      rs("dolar_ult_compra") = 1
   End If
   
   If IsNull(rs("talle")) Then
      rs("talle") = "*"
   End If
   
   If IsNull(rs("color")) Then
      rs("color") = "*"
   End If
   
   If IsNull(rs("medida")) Then
      rs("medida") = "*"
   End If
   
   
   rs.Update

   
    rs.MoveNext
  Wend
  Set rs = Nothing
End If
End Sub

Private Sub Command5_Click()
 'clientes
     Set cnf = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\cac.mdb;User id=claudio" & ";password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
  cnf.Open gconexion
  
  
     Set rs = New ADODB.Recordset
     q = "select * from a1" 'exma
     rs.Open q, cnf
     
     Set rs1 = New ADODB.Recordset
     q = "select * from  vta_01" '5a
     rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
     
     'los copdigos de vendedor y proveedor contienen los viejos
     'despues cuando se migre vendedores y proveedores hay que actualizarlos
     u = 0
     While Not rs.EOF
       Set rs1 = New ADODB.Recordset
       q = "select * from  vta_01 where [id_cliente_anterior] = " & rs("cod-cliente")
       rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
      
       Label3 = u
       Label3.Refresh
       If Not rs1.EOF And Not rs1.BOF Then
         rs1("cuit") = rs("cuit")
         rs1.Update
       End If
       Set rs1 = Nothing
       rs.MoveNext
      u = u + 1
     Wend

End Sub

Private Sub Command6_Click()
J = MsgBox("confirma actualizar tablka iva", 4)
If J = 6 Then
  q = "select * from vta_02 where datevalue([fecha]) >= datevalue('01/10/2010') and datevalue([fecha]) <= datevalue('31/10/2010')"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  
  While Not rs.EOF
       Set rs1 = New ADODB.Recordset
       q = "select * from vta_09 where [num_int] = " & rs("num_int")
       rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs1.EOF And Not rs1.BOF Then
          
       Else
          rs1.AddNew
          rs1("num_int") = rs("num_int")
          rs1("tasa_iva") = 21
          rs1("iva") = rs("iva")
          rs1("neto") = rs("subtotal")
          rs1("tipo_iva") = rs("id_tipo_iva02")
          rs1.Update
       End If
       Set rs1 = Nothing
   rs.MoveNext
  Wend
  End If
MsgBox ("Operacion Terminada")

End Sub

Private Sub Command7_Click()

'se debe vincular la planilla como una tabla a la que se le debe dar un nombre
J = InputBox$("Ingrese Clave de Administrador General")
If J = "0969" Then
  'Call productos_excel
  'Call clientes_excel
  Call productos_excel2
End If


End Sub


Sub productos_excel2()
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim F        As Long
    Dim c     As Long
    Dim CODIGO As String
   Dim nombrehoja As String
   
    nombrehoja = "Hoja 1"
    Path = "E:\a\andreoli\corilux.xls"
     
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    'On Error GoTo e1
    Set o_Libro = o_Excel.Workbooks.Open(Path, True, True, , "")
    'On Error GoTo e2
    Set o_Hoja = o_Libro.Worksheets(nombrehoja)

 
 fin = 500
 
 
 
 a = 0
 cerr = 0
 espere.Show
  
 'On Error GoTo ERRORGRABA
Set rs1 = New ADODB.Recordset
q = "select * from a2"
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic


 For F = 1 To fin
    espere.Label1 = F
    espere.Label1.Refresh
    actu = 0
    CODIGO = o_Hoja.Cells(F, "A").Value
    Desc = o_Hoja.Cells(F, "B").Value
    preciocompra = o_Hoja.Cells(F, "C").Value
     
    If Mid$(CODIGO, 1, 3) = "001" Then
       'actualizo grupo
       
       
    Else
      'agrego producto
      If (Left$(Desc, 150)) <> "" Then
            
            rs1.AddNew
         rs1("Descripcion") = Left$(Desc, 150)
         rs1("id_grupo") = 1
         rs1("id_departamento") = 1
         rs1("id_marca") = 1
         rs1("id_proveedor") = 2
         rs1("precio_ult_compra") = preciocompra
         rs1("fecha_ult_compra") = "01/01/2018"
         rs1("id_proveedor_ult_compra") = 0
         p = Format(preciocompra * 1.5, "######0.00")
         rs1("pu") = Format(p / 1.21, "#####0.00")
         ti = 1
         rs1("cod_tasaiva") = ti
         rs1("id_unidad") = 7
         rs1("envase") = 1
         rs1("stock") = 0
         rs1("requeridos") = 0
         rs1("pedidos") = 0
         rs1("stock_minimo") = 0
         rs1("porc_utilidad") = 50
         rs1("costoreal") = preciocompra
         rs1("flete_compra") = 0
         rs1("dto_compra") = 0
         rs1("cod_barra") = 0
         rs1("precio_final") = p
         rs1("tasa_imp_interno") = 0
         rs1("tipo_producto") = "P"
         rs1("moneda") = "P"
         rs1("impuesto") = 0
         rs1("observaciones") = "*"
         rs1("ultima_compra") = "*"
         rs1("ultima_venta") = "*"
         rs1("fecha_actu_precio_venta") = "1/1/2018"
         rs1("id_anterior") = 0
         rs1("emite_etiqueta") = "N"
         rs1("texto_central") = " "
         rs1("vigente") = True
         rs1("reg_faltante") = 0
         rs1("tipo_carga_tique") = "M"
         rs1("abreviatura") = "*"
         rs1("id_tasaib") = 1
         rs1("id_prod_prov") = 0
         rs1("dto_compra2") = 0
              
              
              
      rs1.Update
     End If
    
    End If
    
Next F
Set rs1 = Nothing
 

End Sub

Sub productos_excel()
'On Error GoTo errtemp
  
  Set rs = New ADODB.Recordset
  q = "select * from migrar" 'aca va la tabla vinculada
  rs.Open q, cn1
  
  Set rs1 = New ADODB.Recordset
  q = "select * from  a2" '5a
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  l = 0
  While Not rs.EOF
     If (Not IsNull(rs("descripcion"))) Then
       rs1.AddNew
         rs1("Descripcion") = Left$(rs("descripcion"), 150)
         rs1("id_grupo") = 1
         rs1("id_departamento") = 1
         rs1("id_marca") = 1
         rs1("id_proveedor") = 1
         rs1("precio_ult_compra") = rs("precio_compra")
         rs1("fecha_ult_compra") = "01/01/2018"
         rs1("id_proveedor_ult_compra") = 1
         p = Format(rs("precio_venta"), "######0.00")
         rs1("pu") = Format(p / 1.21, "#####0.00")
         ti = 1
         rs1("cod_tasaiva") = ti
         rs1("id_unidad") = 7
         rs1("envase") = 1
         rs1("stock") = rs("cantidad")
         rs1("requeridos") = 0
         rs1("pedidos") = 0
         rs1("stock_minimo") = 0
         rs1("porc_utilidad") = 50
         rs1("costoreal") = rs("precio_compra")
         rs1("flete_compra") = 0
         rs1("dto_compra") = 0
         rs1("cod_barra") = 0
         rs1("precio_final") = p
         rs1("tasa_imp_interno") = 0
         rs1("tipo_producto") = "P"
         rs1("moneda") = "P"
         rs1("impuesto") = 0
         rs1("observaciones") = "*"
         rs1("ultima_compra") = "*"
         rs1("ultima_venta") = "*"
         rs1("fecha_actu_precio_venta") = "1/1/2018"
         rs1("id_anterior") = 0
         rs1("emite_etiqueta") = "N"
         rs1("texto_central") = " "
         rs1("vigente") = True
         rs1("reg_faltante") = 0
         rs1("tipo_carga_tique") = "M"
         rs1("abreviatura") = "*"
         rs1("id_tasaib") = 1
         rs1("id_prod_prov") = 0
         rs1("dto_compra2") = 0
       rs1.Update
       l = l + 1
       Label3 = l
       Label3.Refresh
     End If
     rs.MoveNext
  Wend
  MsgBox ("Operacion terminada")




  
  
  
  
  
  
  
  
  
End Sub

Sub clientes_excel()
'On Error GoTo errtemp
  
  Set rs = New ADODB.Recordset
  q = "select * from migrar" 'aca va la tabla vinculada
  rs.Open q, cn1
  
  Set rs1 = New ADODB.Recordset
  q = "select * from  vta_01" '5a
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  l = 0
  While Not rs.EOF
     If (Not IsNull(rs("F2"))) Then
       rs1.AddNew
         rs1("Denominacion") = rs("F2") & " "
         rs1("Direccion") = rs("F3") & " "
         rs1("cp") = "0"
         rs1("provincia") = "Buenos Aires"
         rs1("localidad") = "Rojas"
         rs1("te") = "*"
         
        rs1("cuit") = rs("F1")
        rs1("email") = "*"
        rs1("id_proveedor") = 1
        rs1("limite_credito") = 99999999.99
        rs1("exportacion") = "N"
        rs1("id_cliente_anterior") = 0
        rs1("id_vendedor") = 1
        rs1("id_tipoiva") = 1
        rs1("Observaciones") = "*"
        rs1("inscripto_operador_granos") = "N"
        rs1("percive_ib") = "N"
        rs1("saldo_incobrable") = "N"
        rs1("id_prov") = 2 'provincia ba
        rs1("direccion_local") = rs("F3")
        
        
        rs1.Update
      End If
      rs.MoveNext
      u = u + 1
     
     Wend
    MsgBox ("Operacion terminada")




  
  
  
  
  
  
  
  
  
End Sub

Private Sub Command8_Click()
J = InputBox$("Ingrese Clave de Administrador General")
If J = "0969" Then
  'On Error GoTo errtemp
  Set cnf = New ADODB.Connection
  gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\temp\laloma2.mdb;User id=claudio" & ";password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
  cnf.Open gconexion

  Set rs = New ADODB.Recordset
  q = "select * from a2" 'laloma
  rs.Open q, cnf
     
  Set rs1 = New ADODB.Recordset
  q = "select * from  a2" '5a
  rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
  l = 0
  While Not rs.EOF
       rs1.AddNew
         rs1("Descripcion") = Left$(rs("producto"), 150)
         rs1("id_grupo") = 1
         rs1("id_departamento") = rs("id_d")
         rs1("id_marca") = rs("id_m")
         rs1("id_proveedor") = 1
         rs1("precio_ult_compra") = 0
         rs1("fecha_ult_compra") = "02/01/2013"
         rs1("id_proveedor_ult_compra") = 1
         p = Format(rs("precio"), "######0.00")
         rs1("pu") = Format(p / 1.21, "#####0.00")
         ti = 1
         rs1("cod_tasaiva") = ti
         rs1("id_unidad") = 7
         rs1("envase") = 1
         rs1("stock") = 0
         rs1("requeridos") = 0
         rs1("pedidos") = 0
         rs1("stock_minimo") = 0
         rs1("porc_utilidad") = 0
         rs1("costoreal") = 0
         rs1("flete_compra") = 0
         rs1("dto_compra") = 0
         rs1("cod_barra") = 0
         rs1("precio_final") = p
         rs1("tasa_imp_interno") = 0
         rs1("tipo_producto") = "P"
         rs1("moneda") = "P"
         rs1("impuesto") = 0
         rs1("observaciones") = "*"
         rs1("ultima_compra") = "*"
         rs1("ultima_venta") = "*"
         rs1("fecha_actu_precio_venta") = "18/01/2013"
         rs1("id_anterior") = 0
         rs1("emite_etiqueta") = "N"
         rs1("texto_central") = " "
         rs1("vigente") = True
         rs1("reg_faltante") = 0
         rs1("tipo_carga_tique") = "M"
         rs1("abreviatura") = "*"
         rs1("id_tasaib") = 1
       rs1.Update
       l = l + 1
       Label3 = l
       Label3.Refresh
       rs.MoveNext
  Wend
  MsgBox ("Operacion terminada")
End If

End Sub

Private Sub Command9_Click()
i = InputBox("Ingrese cantidad de grupos consecutivos a generar", , 0)
If Val(i) > 0 Then
    Set rs = New ADODB.Recordset
    q = "select * from vta_01"
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    For i = 0 To Val(i)
      rs.AddNew
      rs.Update
    Next i
    Set rs = Nothing
    MsgBox ("Proceso terminado")
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 27
        
        Me.Hide
End Select
End Sub
Sub carga()
vta_recibo.armagrid
k = 0
r = 1
While k < List1.ListCount
  If List1.Selected(k) = True Then
   F = Mid$(List1.List(k), 1, 10)
   c = Mid$(List1.List(k), 11, 21)
   vta_recibo.msf1.AddItem F & Chr(9) & c & Chr(9) & Mid$(List1.List(k), 33, 10) & Chr(9) & Mid$(List1.List(k), 45, 3) & Chr(9) & Mid$(List1.List(k), 51, 8) & Chr(9) & Mid$(List1.List(k), 61, 8)
   r = r + 1
  End If
   k = k + 1
Wend

   
End Sub
Private Sub Form_Load()
Call barraesag(Me)
Option2 = True
End Sub

  
