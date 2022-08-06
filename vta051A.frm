VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_presup 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PRESUPUESTOS"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11880
   Begin VB.TextBox t_cl 
      Height          =   375
      Left            =   12120
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   9000
      TabIndex        =   44
      Top             =   5400
      Width           =   2775
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprime Descripcion Extra"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   41
      Top             =   6600
      Width           =   6015
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales"
      Height          =   855
      Left            =   6360
      TabIndex        =   36
      Top             =   6480
      Width           =   2535
      Begin VB.TextBox t_total 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox T_total2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Total U$s"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9720
      TabIndex        =   29
      Top             =   720
      Width           =   1935
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9720
      TabIndex        =   26
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contado "
         Height          =   255
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9960
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales del Comprobante"
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   5280
      Width           =   8775
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   6720
         Picture         =   "vta051A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox t_observaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   9
         Top             =   600
         Width           =   6855
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Vendedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5741
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "vta051A.frx":0105
         Left            =   1680
         List            =   "vta051A.frx":0107
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   8520
         Picture         =   "vta051A.frx":0109
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox t_fechavto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_cotizacion 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   7680
         Picture         =   "vta051A.frx":047B
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_letra 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_numcomp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox t_sucursal 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "c_prov"
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Vto.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   39
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cotizacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   35
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9840
      TabIndex        =   15
      Top             =   6480
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta051A.frx":0580
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta051A.frx":0E02
         Style           =   1  'Graphical
         TabIndex        =   12
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
      TabIndex        =   14
      Top             =   7590
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/08/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:59 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_presup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Dim calcula_perc_ib As String
Dim alicuota_perc_ib As Single
Dim minimo_perc_ib As Double
Dim gcuit As String
Dim numint As Long
Dim cuentaact As Long
Dim abreviatura As String
Dim cantlineas As Integer
Dim ubicacionctacte As String


Sub iniciacomp()
Set rs = New ADODB.Recordset
q = "select [imprime_desc_extra] from vta_06 where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = 40"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  If rs("imprime_desc_extra") = "S" Then
    Check2 = 1
  Else
    Check2 = 0
  End If
Else
  Check2 = 0
End If
Set rs = Nothing

Call mensaje
End Sub
Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
   Option1 = True
   
End Sub
Sub mensaje()
'activa mensaje de faturacion
tm = c_tipocomp & " [" & t_letra & "]"
If Option2 = True Then
  tm = tm & "  " & "CONTADO"
Else
  tm = tm & "  " & "CUENTA CORRIENTE"
End If
If c_prov.ListIndex = 0 Then
   tm = tm & " **" & vta_clientes.t_cli & "**"
Else
    tm = tm & " **" & c_prov & "**"
End If
Label20 = UCase$(tm)
Frame11.Visible = True

End Sub
Sub carga()
  Set rs = New ADODB.Recordset
  q = " select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = 40  and [num_comp] = " & Val(t_numcomp)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     t_fecha = rs("fecha")
     t_fechavto = rs("fecha_vto")
    c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     Call armagrid
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("unidad") & Chr$(9) & Format$(rs1("pu"), "######0.00") & Chr(9) & rs1("tasaiva") & Chr(9) & rs1("importe") & Chr(9) & rs1("pu_final") & Chr(9) & rs1("tasaib")
        
        Set rs2 = New ADODB.Recordset
        q = "select * from vta_015 where [num_int] = " & rs1("num_int") & " and [renglon] = " & rs1("renglon")
        rs2.Open q, cn1
        If Not rs2.EOF And Not rs2.BOF Then
           k = rs2("cant_lineas")
           msf1.AddItem 0 & Chr(9) & "" & Chr(9) & rs2("desc_ext") & Chr(9) & k
           msf1.RowHeight(msf1.Rows - 1) = k * 250
 
        End If
        Set rs2 = Nothing
        rs1.MoveNext
     Wend
     Call renumera
     Set rs1 = Nothing
     c_vend.ListIndex = buscaindice(c_vend, rs("id_vendedor"))
     t_total = Format$(rs("total"), "######0.00")
     If rs("contado") = "S" Then
        Option2 = True
     Else
        Option1 = True
     End If
     
     Set rs = Nothing
  
  
   'formas de pago
   
  
  
  Else
     EXISTE = "N"
  End If
  
End Sub

Private Sub btnacepta_Click()

      Call iniciagraba



End Sub

Sub iniciagraba()



If Val(t_total) > 0 Then
 Call mensaje
 J = MsgBox("Graba " & Label20, 4)
 If J = 6 Then
     para.z_actual = 0
     Call normal
 End If
Else
 MsgBox ("Imposible emitir comprobante. El total del comprobante debe ser > 0 ")
End If
  
    

End Sub


Sub normal()
  Set rs = New ADODB.Recordset
  q = "select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = 40 and [num_comp] = " & Val(t_numcomp)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
      EXISTE = "S"
      If para.id_grupo_modulo_actual >= 8 Then
         ni = rs("num_int")
         Set rs = Nothing
         J = MsgBox("Comprobante existente. ¿Desea Modificarlo? ", 4)
         If J = 6 Then
           Set cl_compvta = New comprobantes_venta
           cl_compvta.cargar2 (ni)
           cl_compvta.borrar
           Set cl_compvta = Nothing
           Call graba
         End If
       Else
         MsgBox ("El comprobante existe y Ud. no tiene permisos para modificarlo")
       End If
  Else
    Set rs = Nothing
    EXISTE = "N"
    Call graba
  End If

End Sub
Private Sub btnsale_Click()
J = MsgBox("Abandona el comprobante (S/N)", 4)
If J = 6 Then
  Unload Me
End If
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 12
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 700
msf1.ColWidth(2) = 5000
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 900
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 900
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1100
msf1.ColWidth(9) = 1100
msf1.ColWidth(10) = 1100
msf1.ColWidth(11) = 1000
msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "Unidad"
msf1.TextMatrix(0, 5) = "P.U."
msf1.TextMatrix(0, 6) = "% Iva"
msf1.TextMatrix(0, 7) = "Importe"
msf1.TextMatrix(0, 8) = "PU Final"
msf1.TextMatrix(0, 9) = "Iva"
msf1.TextMatrix(0, 10) = "Costo Tot."
msf1.TextMatrix(0, 11) = "Tasa IB "
End Sub



Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
  End If
End If

If c_prov.ItemData(c_prov.ListIndex) = 1 Then
      Option2 = True
      vta_clientes.t_id = 1
      vta_clientes.carga
      vta_clientes.Show
Else
    Call iniciacli
End If
End Sub


Sub inicia()
espere.Show
espere.Label1 = "Inicializando Comprobante....."
espere.Refresh
' Set cl_cli = New Clientes
'   cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
'   If cl_cli.id > 0 Then
   t_letra = vta_clientes.t_letrafact
   't_sucursal = Format$(val(glo.sucursal, "0000")
   gcuit = vta_clientes.t_cuit
   c_vend.ListIndex = buscaindice(c_vend, vta_clientes.t_idvend)
   Set cl_compvta = New comprobantes_venta
   cl_compvta.sucursal = Val(c_sucursal)
   cl_compvta.actual (40)
   cl_compvta.letra = t_letra
   cl_compvta.SACANUMCOMP
   t_numcomp = Format$(cl_compvta.numcomp, "00000000")
   cantlineas = cl_compvta.cant_lineas
   Set cl_compvta = Nothing
   t_cotizacion = para.cotizacion

     t_alicuotaib = "0.00"
     T_PERCIB = "0.00"
     'gcuit = "0"
   Call armagrid
   Unload espere

   





End Sub

Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)
End If
t_sucursal = Format$(c_sucursal, "0000")
t_numcomp = ""
End Sub


Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
End Sub

Private Sub Check2_LostFocus()
Set rs = New ADODB.Recordset
q = "select [imprime_desc_extra] from vta_06 where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = 40"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
 If Check2 = 1 Then
  rs("imprime_desc_extra") = "S"
 Else
   rs("imprime_desc_extra") = "N"
 End If
 rs.Update
End If
Set rs = Nothing

End Sub

Private Sub Command1_Click()
vta_ABM_vend.Show
End Sub

Private Sub Command1_LostFocus()
c_vend.clear
Call carga_vendedores(c_vend)
c_vend.ListIndex = 0

End Sub

Private Sub Command2_Click()
vta_ABM_cli.Show
End Sub

Private Sub Command2_LostFocus()
c_prov.clear
Call carga_clientes(c_prov)
c_prov.ListIndex = 0
End Sub


Private Sub Command5_Click()
vta_clientes.Show
End Sub


Private Sub Form_Activate()
Frame2.Enabled = False

End Sub
Sub captura()
MsgBox ("Estado Fiscal: " & epson1.FiscalStatus & "  Estado Impresora: " & epson1.PrinterStatus)
End Sub
Sub iniciacli()
 If c_prov.ListIndex > 0 Then
   vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
   vta_clientes.carga
 Else
   If Val(vta_clientes.t_id) <> 0 Then
      vta_clientes.t_id = 0
      vta_clientes.limpia
   
   
   End If
 End If
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
  
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 11)
End If


End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_clientes(c_prov)
c_prov.ListIndex = 0

Call carga_SUCURSALES(c_sucursal)
c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)




Set rs = New ADODB.Recordset
q = "select * from vta_05 order by [denominacion]"
rs.Open q, cn1
Call llena_combo(rs, "denominacion", "id_vendedor", c_vend, True)
Set rs = Nothing
c_vend.ListIndex = 0
Call armagrid
Call barraesag(Me)
Option2 = True
If para.moneda = "P" Then
  Option4 = True
Else
  Option3 = True
End If
t_sucursal = Format$(glo.sucursal, "0000")
Load vta_facturacion1
Load vta_facturacion2
Frame11.Visible = False


Load vta_clientes
vta_clientes.limpia
gcuit = "0"





End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_facturacion1
Unload vta_facturacion2
Unload vta_selremitos
Unload vta_clientes
Unload vta_formapago
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F3] Descipcion extra - [F5] Saca Renglon - [F9] Graba "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 If msf1.Rows > 1 Then
  If Val(msf1.TextMatrix(msf1.Rows - 1, 0)) > 0 Then
   Load gen_descextra
   gen_descextra.t_modulo = "P"
   gen_descextra.t_funcion = "A"
   gen_descextra.Show
  End If
 End If
End If




If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
   r = msf1.Row
   If r + 1 < msf1.Rows Then
      If Val(msf1.TextMatrix(r + 1, 0)) = 0 Then
        msf1.RemoveItem (r + 1)
      End If
   End If
   If msf1.Rows > 2 Then
     msf1.RemoveItem (r)
   Else
     Call armagrid
   End If
   Call renumera
  Else
   Call armagrid
   
 End If
 
End If


If KeyCode = vbKeyF9 Then
  
  Call sacatotales
  Call renumera
  Frame2.Enabled = True
  btnacepta.Enabled = True
  c_vend.SetFocus
End If

If KeyCode = vbKeyInsert Then
   vta_presup1.t_renglon = ""
   vta_presup1.t_cantidad = ""
   vta_presup1.t_pu = ""
   vta_presup1.t_importe = ""
   If msf1.Rows - 1 < cantlineas Then
     vta_presup1.Show
   Else
     MsgBox ("Se ha superado la cantidad maxima dde items para este comprobante")
   End If
End If

If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Sub renumera()
r = 1
For i = 1 To msf1.Rows - 1
 If Val(msf1.TextMatrix(i, 0)) <> 0 Then
    msf1.TextMatrix(i, 0) = r
    r = r + 1
 End If
Next i


End Sub
Sub graba()
  'On Error GoTo ERRORGRABA
  
  numint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(t_sucursal)
  cl_compvta.actual (40)
  cl_compvta.letra = t_letra
  cl_compvta.numcomp = Val(t_numcomp)
  abreviatura = cl_compvta.abreviatura
  ubicacionctacte = cl_compvta.ctacte
     If Option1 = True Then
         ep = "N"
         cp = "Cta.Cte."
         contado = "N"
      Else
         ep = "S"
         cp = "ctdo"
         contado = "S"
     End If
     ssi = 0
      
      If EXISTE = "N" Then
        cl_compvta.ACTUALIZA_NUMERADOR
      End If
      
      If Option4 = True Then
        moneda = "P"
      Else
        moneda = "D"
      End If
      
      
      
        
              
      tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
       
      If c_prov.ListIndex = 0 Then
        idcli = 1
      Else
        idcli = c_prov.ItemData(c_prov.ListIndex)
      End If
      
      T2 = Val(T_total2)
      
      cn1.BeginTrans
       
       
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], " & _
" [dominio_acoplado02], [SALDO_IMPAGO02], [num_z], [cae], [cae_vence], [tipo_op], [numint_asociado])"



QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', 40" & _
", " & idcli & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotal) & ", " & Val(t_nograbado) & ", " & Val(t_iva) & ", " & Val(t_total) & _
", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & t_observaciones & _
" ', " & Val(t_cotizacion) & ", " & T2 & ", '" & moneda & "', " & c_vend.ItemData(c_vend.ListIndex) & ", '" & cl_compvta.venta & "', '" & contado & "', 0" & _
", 0, 0, 1, 3, 0, 0, '" & t_fechavto & "', 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & _
", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & Left$(vta_clientes.t_localidad, 50) & _
"', " & tiporespiva & ", ' ', ' ', ' ', " & ssi & ", " & para.z_actual & ", 'u2', '01/01/2018', 1,0)"

'MsgBox (QUERY)

cn1.Execute QUERY
COSTOINV = 0
Set cl_cli = Nothing
For i = 1 To msf1.Rows - 1
  renglon = Val(msf1.TextMatrix(i, 0))
  If renglon > 0 Then
        
        If Val(msf1.TextMatrix(i, 1)) > 1 Then
          Set cl_prod = New productos
          cl_prod.cargar (Val(msf1.TextMatrix(i, 1)))
          costo = cl_prod.costoreal
          Set cl_prod = Nothing
        Else
          costo = 0
        End If
        
        QUERY = "INSERT INTO vta_03([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [pu], [importe], [tasaiva], [impuesto], [costo], [cantidad_original], [tunidad], [pu_final], [tasaib])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & " ', " & Val(msf1.TextMatrix(i, 3)) & ", " & Val(msf1.TextMatrix(i, 5)) & ", " & Val(msf1.TextMatrix(i, 7)) & ", " & Val(msf1.TextMatrix(i, 6)) & ", 0, " & costo & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & msf1.TextMatrix(i, 4) & "', " & Val(msf1.TextMatrix(i, 8)) & ", " & Val(msf1.TextMatrix(i, 11)) & ")"
        cn1.Execute QUERY
      
        

  Else
    'grabo desc extra
    QUERY = "INSERT INTO vta_015([num_int], [RENGLON], [desc_ext], [cant_lineas])"
    QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i - 1, 0)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 3)) & ")"
    cn1.Execute QUERY
  End If


Next i
      
      
     
      
      
QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
QUERY = QUERY & " VALUES ('Emitir Presupuesto NI:" & numint & "', " & para.id_usuario & ", 'V', " & numint & ", '" & Now & "', '[" & 40 & "] " & t_letra & " " & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 12, " & idcli & ")"
  
     cn1.Execute QUERY

      
      
      cn1.CommitTrans
      
      
     
        
     'End If
      
      
          J = MsgBox("Confirma Impresion del Comprobante", 4)
          If J = 6 Then
             Set cl_compvta = New comprobantes_venta
             cl_compvta.cargar2 (numint)
             cl_compvta.imprimir
          End If
          
      Call INICIALIZA2(Me)
      Call armagrid
      c_sucursal.SetFocus
      Frame2.Enabled = False
      t_sucursal = Format$(c_sucursal, "0000")
      Frame11.Visible = False
      
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos y vuelva a repetir la operacion")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
   If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
     vta_presup1.t_renglon = msf1.Row
     vta_presup1.t_basico = msf1.TextMatrix(msf1.Row, 1)
     vta_presup1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
     vta_presup1.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
     vta_presup1.t_unidad = msf1.TextMatrix(msf1.Row, 4)
     vta_presup1.t_pu = msf1.TextMatrix(msf1.Row, 5)
     vta_presup1.t_importe = msf1.TextMatrix(msf1.Row, 7)
     vta_presup1.Show
   Else
     Load gen_descextra
     gen_descextra.Text1 = msf1.TextMatrix(msf1.Row, 2)
     gen_descextra.t_modulo = "P"
     gen_descextra.t_funcion = "M"
     gen_descextra.Show
   End If
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True

End Sub

Private Sub Option1_GotFocus()
Call keyform(Me, "D")

End Sub

Private Sub Option1_LostFocus()
Call keyform(Me, "A")

End Sub

Private Sub Option2_GotFocus()
'all keyform(Me, "A")

End Sub

Private Sub Option2_LostFocus()
'Call keyform(Me, "D")

End Sub

Private Sub Option3_Click()
Label13 = "Total $"
End Sub

Private Sub Option4_Click()
Label13 = "Total U$s"
End Sub

Private Sub Option4_GotFocus()
'Call keyform(Me, "A")


End Sub

Private Sub Option4_LostFocus()
'Call keyform(Me, "D")

End Sub


Private Sub t_cotizacion_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_cotizacion_LostFocus()
If Val(t_cotizacion) <= 0 Then
   t_cotizacion = 1
End If
End Sub

Private Sub t_fecha_GotFocus()
If glo.sucursalf = Val(t_sucursal) Then
   t_fecha = Format$(Now, "dd/mm/yyyy")
   t_fecha.Locked = True
Else
   t_fecha.Locked = False
End If

End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub


Private Sub t_fechavto_LostFocus()
If Not IsDate(t_fechavto) Then
  t_fechavto = Format$(Now, "dd/mm/yyyy")
Else
  t_fechavto = Format$(t_fechavto, "dd/mm/yyyy")
End If

End Sub



Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub

Private Sub t_numcomp_LostFocus()
If IsNumeric(t_numcomp) Then
   t_numcomp = Format$(t_numcomp, "00000000")
   'If glo.sucursalf <> Val(c_sucursal) Then
     Call carga
   'Else
   '  EXISTE = "N"
   'End If
   
  
   Call iniciacomp

Else
  t_numcomp.SetFocus
End If
End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub


Sub sacatotales()
s = 0
For i = 1 To msf1.Rows - 1
 renglon = Val(msf1.TextMatrix(i, 0))
 If renglon > 0 Then
   r = Val(msf1.TextMatrix(i, 8)) * Val(msf1.TextMatrix(i, 3))
   s = s + r
 End If
Next i
   
t_total = Format$(s, "######0.00")
If Option4 = True Then
 If Val(t_cotizacion) < 1 Then
   t_cotizacion = 1
 End If
 T_total2 = Format$(Val(t_total) / Val(t_cotizacion), "#####0.00")
Else
  T_total2 = Format$(Val(t_total) * Val(t_cotizacion), "#####0.00")
End If
End Sub

Private Sub t_sucursal_GotFocus()
t_sucursal = Format$(Val(c_sucursal), "0000")
End Sub

Private Sub t_sucursal_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
End If
If c_prov.ItemData(c_prov.ListIndex) = 1 Then
   Call iniciacli
End If
Call inicia
End Sub

Private Sub t_total_LostFocus()
t_total = Format$(t_total, "######0.00")
End Sub

Private Sub T_total2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.Enabled = True
 btnacepta.SetFocus
End If

End Sub


