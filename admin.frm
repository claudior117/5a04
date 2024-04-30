VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form admin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ADMINISTRADOR GENERAL SISTEMA "
   ClientHeight    =   2205
   ClientLeft      =   105
   ClientTop       =   645
   ClientWidth     =   3900
   Icon            =   "admin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2205
   ScaleWidth      =   3900
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingresar"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox T_password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox c_usuario 
         Height          =   315
         ItemData        =   "admin.frx":0ECA
         Left            =   1560
         List            =   "admin.frx":0ECC
         TabIndex        =   1
         Text            =   "c_usuario"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin.frx":0ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin.frx":1762
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin.frx":1FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin.frx":288A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin.frx":311E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin.frx":39B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin.frx":3CD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu M_tools 
      Caption         =   "Tools"
      Begin VB.Menu M_duplicar 
         Caption         =   "Duplicar sistema para Transporte"
      End
      Begin VB.Menu M_optimizar 
         Caption         =   "Optimizar Base de Datos"
      End
      Begin VB.Menu M_resguarda 
         Caption         =   "Resguradar Bases de Datos"
      End
      Begin VB.Menu M_habilitarterm 
         Caption         =   "Habilitar Terminal"
      End
      Begin VB.Menu M_salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu M_soporte 
      Caption         =   "SOPORTE REMOTO"
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub c_usuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If c_usuario.ListIndex > -1 Then
   T_password.Enabled = True
   T_password.SetFocus
 End If
End If
End Sub

Private Sub Command1_Click()
Attribute Command1_Click.VB_HelpID = 90834298
If inicializa Then
 
If c_usuario.ListIndex >= 0 Then
 
 If abrirconexion(c_usuario, T_password) = True Then
   para.id_usuario = c_usuario.ItemData(c_usuario.ListIndex)
   Set rs = New ADODB.Recordset
   q = "select * from g1 where [id_usuario] = " & para.id_usuario
   rs.Open q, cn1
   If Not rs.EOF And Not rs.BOF Then
     para.id_grupo = rs("grupo")
     glo.usuario = rs("usuario")
     para.usuario = rs("usuario")
     para.muestraagenda = rs("muestra_agenda")
     para.punto_venta_usuario = rs("punto_venta_inicio")
     para.IMPRESORA_PREDETERMINADA = Printer.DeviceName
     para.impresora_actual = para.IMPRESORA_PREDETERMINADA
     Call carga_estados_a
     para.imprime_pie_reportes = rs("imprime_pie_reportes")
     para.tipolistaprecios = rs("tipo_lista_precios")
     para.imprime_cabecera_reportes = rs("imprime_cabecera_reportes")
     para.usa_separador_miles = rs("usa_separador_miles")
     Set rs = Nothing


      If para.usa_separador_miles = "N" Then
         para.formato_numerico = "#0.00"
      Else
         para.formato_numerico = "#,##0.00"
      End If
     
     Set rs = New ADODB.Recordset
     q = "select * from g0 where [sucursal] = 0"
     rs.Open q, cn1
     If Not rs.BOF And Not rs.EOF Then
        para.cuenta_acreedores = rs("id_cuenta_acreedores")
        para.cuenta_deudores = rs("id_cuenta_deudores")
        para.cuenta_ventas = rs("id_cuenta_ventas")
        para.cuenta_iva_compras = rs("id_cuenta_iva_compras")
        para.cuenta_iva_ventas = rs("id_cuenta_iva_ventas")
        para.cuenta_retgan = rs("id_cuenta_ret_gan")
        para.cuenta_retib = rs("id_cuenta_ret_ib")
        para.cuenta_conceptos_nograbados = rs("id_cuenta_nograbados")
        para.cuenta_perc_IB = rs("id_cuenta_perc_ib")
        para.cuenta_perc_iva = rs("id_cuenta_perc_iva")
        para.cuenta_compras_varias = rs("id_cuenta_compras_varias")
        para.cuenta_retibbav = rs("id_cuenta_retibba_ventas")
        para.cuenta_retivav = rs("id_cuenta_retiva_ventas")
        para.cuenta_retganv = rs("id_cuenta_retgan_ventas")
        para.cuenta_retsussv = rs("id_cuenta_retsuss_ventas")
        para.numeracion_comun_Fact_nc = rs("numeracion_comun_fact_nc")
        para.cotizacion = rs("cotizacion")
        para.tasageneral = rs("tasa_general_iva")
        para.minimo_retib = rs("minimo_retib")
        para.producto_sel = 0
        para.id_periodo_contable = rs("id_periodo_contable")
        para.tipoactupreciocompcompra = rs("tipo_actu_precio_comp_compra")
        para.tipoprecioventa = rs("tipo_precio_venta")
        para.cuenta_inventario = rs("id_cuenta_inventario")
        para.cuenta_costo = rs("id_cuenta_costo_merc")
        'glo.sucursal = rs("sucursal_actual")
        glo.sucursalprueba = rs("sucursal_prueba")
        para.archivo_exportacion = "C:\exporta"
        para.fechacorte = rs("fecha_corte")
        para.tiporedondeo = rs("tipo_redondeo")
        para.ncenrecibo = rs("nc_en_recibo")
        para.muestrasaldofactventa = rs("muestra_saldo_fact_venta")
        
       para.nombre_fantasia = rs("nombre_fantasia")
       para.fecha_inicio_actividades = rs("fecha_inicio_actividades")
       para.numero_ib = rs("numero_ingresos_brutos")
       para.idsistema = rs("id_sistema")
       'define tipo de iva del cliente del sistema (1 fact ay b, el resto factura C)
       para.tipo_iva_empresa = rs("id_tipo_iva")
       
       
        
        Set rs1 = New ADODB.Recordset
        q = "select * from cyb_01 where [id_forma_pago] = 1"
        rs1.Open q, cn1
        If Not rs1.BOF And Not rs1.EOF Then
           para.cuenta_caja = rs1("id_cuenta_cont")
        End If
        Set rs1 = Nothing
     
     
        Set rs1 = New ADODB.Recordset
        q = "select * from cyb_01 where [id_forma_pago] = 3"
        rs1.Open q, cn1
        If Not rs1.BOF And Not rs1.EOF Then
           para.cuenta_valores_terceros = rs1("id_cuenta_cont")
        Else
           para.cuenta_valores_terceros = 0
        End If
        Set rs1 = Nothing

        Set rs1 = New ADODB.Recordset
        q = "select * from i_01 where [id_impuesto] = 1" 'percepcion ib
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
            para.calcula_perc_ib = rs1("calcula")
        Else
            para.calcula_perc_ib = "N"
        End If
        Set rs1 = Nothing
        
        Set rs1 = New ADODB.Recordset
        q = "select * from i_01 where [id_impuesto] = 50" 'ret ib
        rs1.Open q, cn1
        If Not rs1.EOF And Not rs1.BOF Then
            para.calcula_ret_ib = rs1("calcula")
        Else
            para.calcula_ret_ib = "N"
        End If
        Set rs1 = Nothing
        
        Set rs = New ADODB.Recordset
        q = "select * from a5 where [id_tipocomp] = 60"
        rs.MaxRecords = 1
        rs.Open q, cn1
        If Not rs.EOF And Not rs.BOF Then
           para.numint_regfaltante = rs("num_int")
        Else
           MsgBox ("ERROR. Falta inicializar el registro de Faltantes")
           para.numint_regfaltante = 0
           'End
        End If
        Set rs = Nothing
     
        Set rs1 = New ADODB.Recordset
        q = "select * from g12 where [id_tasaib] = 1"
        rs1.Open q, cn1
        If Not rs1.BOF And Not rs1.EOF Then
           para.tasaib = rs1("tasaib")
        Else
           MsgBox ("Error al inicializar tasa general de IB (g12)")
        End If
        Set rs1 = Nothing
     
     
     Else
        MsgBox ("Error al Inicializar el sistema")
        End
     End If
     Set rs = Nothing
     
     If abrirconexionib = True Then
       inicio.Show
       Unload Me
     Else
       MsgBox ("Error al abrir base de datos del padron de Ingresos Brutos")
     End If
   Else
     MsgBox ("Error al Inicializar el Usuario")
   End If
   Set rs = Nothing
 End If
End If
Else
 MsgBox ("Operacion no Permitida. Llame al servicio Tecnico")
 End
End If

End Sub
Sub carga_estados_a()
'estados materiales
a_estado_m(0) = "R"
a_estado_m(1) = "P"
a_estado_m(2) = "S"
a_estado_m(3) = "C"
a_estado_m(4) = "O"

'estados obras
a_estado_o(0) = "T"
a_estado_o(1) = "E"
a_estado_o(2) = "S"
a_estado_o(3) = "O"


End Sub
Private Sub Command2_Click()
End
End Sub


Private Sub Form_Load()
If App.PrevInstance Then
    J = MsgBox("Encontramos otra instancia de la aplicacion ejecutandose. ¿Desea continuar abriendo el programa?", 4)
    If J = 6 Then
        Call entrar
    Else
        End
    End If
Else
   Call entrar
End If

End Sub
Sub entrar()
        'para.empresa = "geser" 'command$
        'fotosv   sistema2
        para.empresa = "" 'prueba
        Call LEEINI
        X = Shell(App.Path & "\tools\confreg.exe")
        Call carga_usuarios_ini(c_usuario)
        c_usuario.ListIndex = para.usuario_inicio
        para.password_adm = "1975"

End Sub
Sub compactardb()
Dim JRO As New JRO.JetEngine
Dim BD_Original As String, Dest_DB As String

espere.Show
espere.Label1 = "Espere... Resguardando Bases Originales den carpeta BAK"
espere.Refresh

dbo = App.Path & "\dat\5a04.mdb" 'bd original
dbz = App.Path & "\dat\5a04z.mdb" 'bd compactada
dbpio = App.Path & "\dat\pib.mdb" 'bd padron ib original
dbpiz = App.Path & "\dat\pibz.mdb" 'bd padron ib compactada
dboc = App.Path & "\bak\5a04.mdb" 'bd original copia
dbpic = App.Path & "\bak\pib.mdb" 'bd pib copia
If Dir(dboc) <> "" Then Kill dboc
If Dir(dbpic) <> "" Then Kill dbpic
FileCopy dbo, dboc
FileCopy dbpio, dbpic

espere.Label1 = "Espere... Compactando Base de Datos del Sistema"
espere.Label1.Refresh

    DoEvents
        u = "Claudio"
        p = "0969"
        BD_Original = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\5a04.mdb;User id=" & u & ";password=" & p & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
        Dest_DB = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\5a04z.mdb;User id=" & u & ";password=" & p & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system1.mdw;"
        JRO.CompactDatabase BD_Original, Dest_DB 'compacto
        Set JRO = Nothing
        espere.Label1 = "Espere... Compactando Padron I.B."
        espere.Label1.Refresh
        BD_Original = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\pib.mdb;User id=" & u & ";password=" & p & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system2.mdw;"
        Dest_DB = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\pibz.mdb;User id=" & u & ";password=" & p & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system2.mdw;"
        JRO.CompactDatabase BD_Original, Dest_DB 'compacto
        Set JRO = Nothing
        espere.Label1 = "Espere... Reemplazando Bases Originales "
        espere.Label1.Refresh
        Kill (dbo)
        Kill (dbpio)
        Name dbz As dbo
        Name dbpiz As dbpio
        Unload espere
        MsgBox ("Operacion Terminada")
        
        Exit Sub
        
        
End Sub

Sub LEEINI()
'On Error GoTo errorini
's = 0 todas las empresas
's = 1 center clean que la maquina de cristina win 98 no ecuentar la ruta
s = 1
If para.empresa = "" Then
 If s = 1 Then
  Open "C:\5A04\GEN\5a04.INI" For Input As #1
 Else
  Open "C:\A5.INI" For Input As #1
 End If
Else
   carpeta = "C:\" & para.empresa & "\5A04\GEN\5a04.INI"
   Open carpeta For Input As #1
End If
Line Input #1, l
glo.nombrecli = l
Line Input #1, l
glo.direccioncli = l
Line Input #1, l
glo.TECLI = l
Line Input #1, l
glo.CUIT = l
Line Input #1, l
glo.SERIAL = Mid$(l, 1, 9)
Line Input #1, l
para.moneda = l
Line Input #1, l
para.sincroniza_bancos = l
Line Input #1, l
para.usuario_inicio = Val(l)
Line Input #1, l
glo.sucursalf = Val(l) '0 si no tiene conectada la impresora fiscal
Line Input #1, l
glo.sucursale = Val(l) '0 si no tiene conectada la impresora electronica
Line Input #1, l
glo.sucursal = Val(l) '0 punto de venta manual


Close #1

Exit Sub
errorini:
 MsgBox ("Error en archivo de Inicializacion")
 End

End Sub
Function inicializa() As Boolean
Call SACARSERIAL
Open "c:\windows\system\temp.txt" For Input As #1
Line Input #1, t
Close #1
If Mid$(t, 1, 9) <> glo.SERIAL Then
   inicializa = False
  
  Else
 ' p1 = GetDeviceCaps(HDC, horres)
 ' p2 = GetDeviceCaps(HDC, verres)

  'If p1 < 800 Or p2 < 600 Then
  ' MsgBox ("Su resolucion de Pantalla actual es de " & p1 & "x" & p2 & " el Sistema solo funcion en 800x600")
  ' End
  'End If
   inicializa = True
End If
End Function
Sub inicia()
      T_usuario = ""
      T_password = ""
End Sub



Private Sub M_duplicar_Click()
J = InputBox$("Ingrese Password de Administrador General")
If J = para.password_adm Then
   gen_duplica.Show
End If
  
End Sub

Private Sub M_habilitarterm_Click()
J = MsgBox("Este proceso habilita el uso de una terminal. Ejecutelo en la terminal a habilitar posterior a configurar el sistema en un servidor. El archivo de inicializacion en la carpeta del sistema debe estar configurado correctamente.", 4)
If J = 6 Then
  J = InputBox$("Ingrese Password de Supervisor")
  If J = "0969" Then
     Call habilitaterminal
  End If
End If

End Sub

Sub habilitaterminal()
d1 = "c:\5a04"
d2 = "c:\5a04\gen"
a1 = "c:\5a04\gen\5a04.ini"
o = App.Path & "\gen\5a04.ini"
d = "c:\5a04\gen\5a04.ini"
If Dir(d1) = "" Then
   MkDir (d1)
End If

If Dir(d2) = "" Then
   MkDir (d2)
End If

If Dir(a1) <> "" Then
   Kill (a1)
End If
FileCopy o, d

Call SACARSERIAL
Open "c:\temp.txt" For Input As #1
Line Input #1, t
Close #1

If IsNull(para.empresa) Then
  Open "C:\5A04\GEN\5a04.INI" For Input As #1
Else
   carpeta = "C:\" & para.empresa & "\5A04\GEN\5a04.INI"
   Open carpeta For Input As #1
End If


End Sub
Private Sub M_optimizar_Click()
J = MsgBox("Este proceso reemplaza las bases de datos originales por bases de datos optimizadas. Por seguridad las bases de datos originales seran resguardadas en la carpeta BAK. El proceso puede tardar varios minutos ¿Confirma?", 4)
If J = 6 Then
  J = InputBox$("Ingrese Password de Administrador General")
  If J = para.password_adm Then
     Call compactardb
  End If
End If
End Sub

Private Sub M_resguarda_Click()
J = InputBox$("Ingrese Password de Administrador General")
If J = para.password_adm Then
   gen_resguardo.Show
End If
End Sub

Private Sub M_salir_Click()
End
End Sub

Private Sub m_sincroniza_Click()
gen_sincronizar.Show
End Sub

Private Sub M_soporte_Click()
On Error GoTo errr
J = MsgBox("Este modulo habilitará su maquina para ser manejada remotamente. Una vez iniciada la sesion tendra que pasar al soporte SOLO POR TELEFONO  el Id y el password generado. Al salir del sistema el control remotot se desabilitará y su equipo no correra ningun riesgo", 4)
If J = 6 Then
  X = Shell("c:\5a04\tools\Teamviewerqs_es.exe", vbNormalFocus)
End If

Exit Sub

errr:
MsgBox ("ERROR! El Soporte Remoto NO se encuentra disponible. Se intentará acceder del servidor, si el error se repite llame al soporte tecnico ")
FileCopy App.Path & "\tools\teamviewerqs_es.exe", "c:\5a04\tools\teamviewerqs_es.exe"
Resume


End Sub

Private Sub T_password_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Command1.SetFocus
End If
End Sub
