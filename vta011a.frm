VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_COMPVARIOS 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPROBANTES VARIOS"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10080
      TabIndex        =   56
      Top             =   1440
      Width           =   1455
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Moneda Unica"
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   10080
      TabIndex        =   39
      Top             =   720
      Width           =   1455
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   10080
      TabIndex        =   36
      Top             =   0
      Width           =   1455
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contado "
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9120
      TabIndex        =   33
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   34
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
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales del Comprobante"
      Height          =   2415
      Left            =   240
      TabIndex        =   28
      Top             =   5760
      Width           =   8775
      Begin VB.CommandButton Command1 
         Caption         =   "Percepciones"
         Height          =   255
         Left            =   2880
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox t_cae 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6600
         MaxLength       =   30
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Text            =   "c_cuenta"
         Top             =   1320
         Width           =   6015
      End
      Begin VB.ComboBox c_actividad 
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Text            =   "c_cuenta"
         Top             =   960
         Width           =   6015
      End
      Begin VB.TextBox t_observaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   9
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox T_total2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   43
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox t_cotizacion 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox t_total 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox t_iva 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox t_perc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox t_nograbado 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox t_subtotal 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "CAE"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   58
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Actividad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Contable:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Total U$s"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   44
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cotizacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   32
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   31
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "No Grabado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Subtotal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   3735
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   240
      TabIndex        =   23
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton Command3 
         Caption         =   " Remitos"
         Height          =   375
         Left            =   8160
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "vta011a.frx":0000
         Left            =   7560
         List            =   "vta011a.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox t_propio 
         Height          =   285
         Left            =   5640
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   8880
         Picture         =   "vta011a.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox t_fechavto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   8040
         Picture         =   "vta011a.frx":0376
         Style           =   1  'Graphical
         TabIndex        =   46
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
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
         Left            =   3480
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
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Text            =   "c_prov"
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   53
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Vto.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   49
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   20
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta011a.frx":047B
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta011a.frx":0CFD
         Style           =   1  'Graphical
         TabIndex        =   21
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
      TabIndex        =   19
      Top             =   8265
      Width           =   11760
      _ExtentX        =   20743
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
Attribute VB_Name = "vta_COMPVARIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Function verificatasaunica() As Boolean
   'devuelve true si la existe una sola tasa en la factura
 i = 1
 v = True
 While i <= msf1.Rows - 1
  If i = 1 Then
    tasa = Val(msf1.TextMatrix(i, 5))
  End If
  If tasa <> Val(msf1.TextMatrix(i, 5)) Then
    v = False
    i = msf1.Rows
  End If
  i = i + 1
 Wend
 verificatasaunica = v
End Function

Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
   Option1 = True
End Sub
Sub carga()
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(c_sucursal)
  cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
  If cl_compvta.PROPIO = "S" Then
    Set rs = New ADODB.Recordset
    q = " select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
    rs.Open q, cn1
  Else
    'COMP.ENTREGADOS POR EL CLIENTE
    Set rs = New ADODB.Recordset
    q = " select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp) & " AND [ID_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
    rs.Open q, cn1
  End If
  
  If Not rs.BOF And Not rs.EOF Then
     MsgBox ("Comprobante Existente")
     Call armagrid
     EXISTE = "S"
     t_fecha = rs("fecha")
     c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & Format$(rs1("pu"), "######0.00") & Chr(9) & rs1("tasaiva") & Chr(9) & rs1("importe")
        rs1.MoveNext
     Wend
     Set rs1 = Nothing
     
     
        'cargo percepciones
     Set rs2 = New ADODB.Recordset
     q = "select * from vta_016, a12 where [num_int] = " & rs("num_int") & " and vta_016.id_percepcion = a12.id_percepcion"
     rs2.Open q, cn1
     
     ABM_COMP_COMPRA2.armagrid
     i = 1
     While Not rs2.EOF
       ABM_COMP_COMPRA2.msf1.AddItem i & Chr$(9) & rs2("vta_016.id_percepcion") & Chr$(9) & rs2("descripcion") & Chr$(9) & rs2("importe") & Chr$(9) & rs2("vta_016.id_cuenta")
       rs2.MoveNext
       i = i + 1
     Wend
     Set rs2 = Nothing

     t_subtotal = Format$(rs("subtotal"), "######0.00")
     t_nograbado = Format$(rs("impuestos"), "######0.00")
     t_perc = Format$(rs("perc_iva") + rs("perc_gan") + rs("perc_ib"), "######0.00")
     t_iva = Format$(rs("iva"), "######0.00")
     t_total = Format$(rs("total"), "######0.00")
     Set rs = Nothing
     
     
  Else
     EXISTE = "N"
  End If
  Set cl_compvta = Nothing
End Sub

Private Sub btnacepta_Click()
If Val(t_iva) > 0 And msf1.Rows <= 1 Then
 J = MsgBox("Esta ingresando un comprobante con IVA pero sin definir la tasa. Si continua puede que los totals en el listado de Iva No coincdan. ¿Continua?", 4)
   If J = 6 Then
     c = 1
   Else
     c = 0
   End If
Else
  c = 1
End If
  

If c = 1 Then
J = MsgBox("Graba Comprobante", 4)
If J = 6 Then
 If verificaperiodog(t_fecha) = "A" Then
  Set rs = New ADODB.Recordset
  If t_propio = "S" Then
     q = "select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
  Else
      q = "select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp) & " and [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
      If para.id_grupo_modulo_actual >= 8 Then
         ni = rs("num_int")
         Set rs = Nothing
         J = MsgBox("Comprobante existente. ¿Desea Modificarlo? ", 4)
         If J = 6 Then
  
           
           Set cl_compvta = New comprobantes_venta
           cl_compvta.cargar2 (ni)
           cl_compvta.borrar
           Set cl_compvta = Nothing
           EXISTE = "S"
           Call verifica
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
 Else
  MsgBox ("Periodo cerrado, imposible realizar operacion")
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
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 4500
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 1200
msf1.ColWidth(7) = 1200

msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "P.U."
msf1.TextMatrix(0, 5) = "% Iva"
msf1.TextMatrix(0, 6) = "Importe"
msf1.TextMatrix(0, 7) = "PU Final"


End Sub


Private Sub c_actividad_LostFocus()
If c_actividad.ListIndex < 0 Then
  c_actividad.ListIndex = 0
End If

End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  If Val(c_cuenta) > 0 Then
    c_cuenta.ListIndex = buscaindice(c_cuenta, Val(c_cuenta))
  Else
    c_cuenta.ListIndex = 0
  End If
End If
End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
  End If
End If
Call iniciacli
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
Sub inicia()
Set cl_compvta = New comprobantes_venta
cl_compvta.sucursal = Val(c_sucursal)
cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
If cl_compvta.PROPIO = "N" Then
  'comprobantes ingresados por el cliente
   If cl_compvta.idtipocomp >= 60 And cl_compvta.idtipocomp <= 65 Then
     'liquidaciones por terceros llevan letra A, B
     Set cl_cli = New Clientes
     cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
     If cl_cli.id > 0 Then
       t_letra = cl_cli.letra
       cl_compvta.letra = t_letra
      Else
       MsgBox ("Error. No se puedo Inicializa el Cliente")
     End If
     Set cl_cli = Nothing
   Else
     'otros comprobantes de terceros llevan letra X
      t_letra = "X"
   End If
Else
     Set cl_cli = New Clientes
     cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
     If cl_cli.id > 0 Then
       t_letra = cl_cli.letra
       t_sucursal = Format$(c_sucursal, "0000")
       cl_compvta.letra = t_letra
       Call cl_compvta.SACANUMCOMP
       t_numcomp = Format$(cl_compvta.numcomp, "00000000")
     
     Else
       MsgBox ("Error. No se puedo Inicializa el Cliente")
     End If
     Set cl_cli = Nothing
End If
Set cl_compvta = Nothing
t_cotizacion = para.cotizacion
End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
End Sub


Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)
End If
t_sucursal = Format$(c_sucursal, "0000")
t_numcomp = ""

End Sub

Private Sub c_tipocomp_LostFocus()
Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & glo.sucursal & " and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex)
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  t_propio = rs("propio")
Else
  t_propio = "S"
End If
Call buscacuentacomp
Set rs = Nothing
End Sub
Sub buscacuentacomp()
Select Case c_tipocomp.ItemData(c_tipocomp.ListIndex)
    Case Is = 100
        c = para.cuenta_retibbav
    Case Is = 101
        c = para.cuenta_retivav
    Case Is = 102
        c = para.cuenta_retganv
    Case Is = 103
        c = para.cuenta_retsussv
        
End Select
c_cuenta.ListIndex = buscaindice(c_cuenta, c)
   
   
 
End Sub

Private Sub Command1_Click()
ABM_COMP_COMPRA2.t_modulo = "S"
ABM_COMP_COMPRA2.Show
End Sub

Private Sub Command2_Click()
vta_ABM_cli.Show
End Sub

Private Sub Command2_LostFocus()
c_prov.clear
Call carga_clientes(c_prov)
c_prov.ListIndex = 0
End Sub

Private Sub Command3_Click()
vta_selremitos2.carga
vta_selremitos2.Show
End Sub

Private Sub Command5_Click()
vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
vta_clientes.carga
vta_clientes.Show

End Sub

Sub actualizaremitos()
For J = 1 To msf1.Rows - 1
 If Val(msf1.TextMatrix(J, 1)) > 1 Then
  cantidadf = Val(msf1.TextMatrix(J, 3)) 'cantidad facturada
  codprodant = Val(msf1.TextMatrix(J, 1))  'cantidad facturada
  i = 1  'para cada articulo busco en los remitos seleccionados
  While i < vta_selremitos2.msf1.Rows
   If vta_selremitos2.msf1.TextMatrix(i, 0) = "**" Then
     nir = Val(vta_selremitos2.msf1.TextMatrix(i, 4))
     q = "SELECT * FROM VTA_02 WHERE [NUM_INT] = " & nir
     Set rs = New ADODB.Recordset
     rs.Open q, cn1, adOpenDynamic, adLockOptimistic
     If Not rs.EOF And Not rs.BOF Then
        'busco el producto en el remito
        q = "select * from vta_03 where [num_int] = " & nir & " and [id_producto] = " & codprodant
        Set rs1 = New ADODB.Recordset
        rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
        While Not rs1.EOF
             'si encontre el producto en el remito
             'verifico cantidad a facturar cantidadf contra lo que hay en el remito
             If cantidadf >= rs1.Fields("Cantidad") Then
                cantidadf = cantidadf - rs1.Fields("Cantidad")
                cpend = 0
                rs1("cantidad") = cpend
                rs1.Update
            
             Else
                cpend = rs1.Fields("cantidad") - cantidadf
                cantidadf = 0
                rs1("cantidad") = cpend
                rs1.Update
                rs1.MoveLast
                i = vta_selremitos2.msf1.Rows
             End If
              
            rs1.MoveNext
         Wend
         Set rs1 = Nothing
         
         If verificaremito(nir) = 0 Then
             rs("estado") = "F"
             rs.Update
         End If
        End If
      Set rs = Nothing
    End If
    i = i + 1
  Wend
 End If
Next J
End Sub

Function verificaremito(ByVal n As Long) As Integer
q = "select * from vta_03 where [num_int] = " & n
Set rs1 = New ADODB.Recordset
rs1.Open q, cn1
p = 0
While Not rs1.EOF
  If rs1("id_producto") > 1 Then
    If rs1("cantidad") > 0 Then
      p = 1
    End If
  End If
  rs1.MoveNext
Wend
verificaremito = p
End Function

Sub CALCULATOTALES()
vta_facturacion2.armagrid
If t_letra = "A" Then
  s = 0
  v = 0
  For i = 1 To msf1.Rows - 1
      r = Val(msf1.TextMatrix(i, 6))
      s = s + r
      v = v + (r * Val(msf1.TextMatrix(i, 5)) / 100)
      
      'agrega en composicion de iva
      X = 1
      While X < vta_facturacion2.msf1.Rows
        If Val(vta_facturacion2.msf1.TextMatrix(X, 0)) = Val(msf1.TextMatrix(i, 5)) Then
           vta_facturacion2.msf1.TextMatrix(X, 1) = Val(vta_facturacion2.msf1.TextMatrix(X, 1)) + r
           vta_facturacion2.msf1.TextMatrix(X, 2) = Val(vta_facturacion2.msf1.TextMatrix(X, 2)) + (r * Val(msf1.TextMatrix(i, 5)) / 100)
           X = vta_facturacion2.msf1.Rows
        Else
           X = X + 1
        End If
      Wend
  
      
  
  Next i
  vta_facturacion2.sacatotales
  t_subtotal = vta_facturacion2.msf1.TextMatrix(9, 1)
  t_iva = vta_facturacion2.msf1.TextMatrix(9, 2)
  Call sacatotales
  'Call sacaperc
 ' Call sacatotales
 Else
  s = 0
  v = 0
  t = 0
  For i = 1 To msf1.Rows - 1
      r = Val(msf1.TextMatrix(i, 6))
      r2 = Val(msf1.TextMatrix(i, 7))
      s = s + r
      t = t + (r2 * Val(msf1.TextMatrix(i, 3)))
  
            'agrega en composicion de iva
      X = 1
      While X < vta_facturacion2.msf1.Rows
        If Val(vta_facturacion2.msf1.TextMatrix(X, 0)) = Val(msf1.TextMatrix(i, 5)) Then
           vta_facturacion2.msf1.TextMatrix(X, 1) = Val(vta_facturacion2.msf1.TextMatrix(X, 1)) + r
           vta_facturacion2.msf1.TextMatrix(X, 2) = Val(vta_facturacion2.msf1.TextMatrix(X, 2)) + (r * Val(msf1.TextMatrix(i, 5)) / 100)
           X = vta_facturacion2.msf1.Rows
        Else
           X = X + 1
        End If
      Wend
  
  
  Next i
  t_subtotal = s
  t_iva = t - s
  Call sacatotales
  'Call sacaperc
  Call sacatotales
  
 End If
  
  
'  s = 0
'  V = 0
'  For i = 1 To vta_COMPVARIOS.msf1.Rows - 1
'      r = vta_COMPVARIOS.msf1.TextMatrix(i, 6)
'      s = s + r
'      V = V + (r * vta_COMPVARIOS.msf1.TextMatrix(i, 5) / 100)
'  Next i
'  vta_COMPVARIOS.t_subtotal = s
'  vta_COMPVARIOS.t_iva = V
'  vta_COMPVARIOS.sacatotales

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
  
End Select
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call TabEnter2(Me, 17)
End If


End Sub

Private Sub Form_Load()

Call INICIALIZA2(Me)
Call carga_clientes(c_prov)
c_prov.ListIndex = 0

Call carga_vendedores(c_vend)
c_vend.ListIndex = 0

Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "Sin Imputacion", 0
c_cuenta.ListIndex = 0
Call carga_SUCURSALES(c_sucursal)
c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)
Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & glo.sucursal & " and  [id_tipocomp] > 0  AND [id_tipocomp] < 500 order by descripcion"
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_tipocomp", c_tipocomp, True)
Set rs = Nothing

c_tipocomp.ListIndex = buscaindice(c_tipocomp, 1)


Call armagrid
Call barraesag(Me)
Option1 = True
Option4 = True
Load vta_COMPVARIOS1
Load ABM_COMP_COMPRA2
Call carga_actividades(c_actividad)

Load vta_clientes
Load vta_selremitos2
Check1 = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_COMPVARIOS1
Unload vta_clientes
Unload vta_selremitos2
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False

End Sub
Sub verifica()
If t_fecha = "" Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  If Not IsDate(t_fecha) Then
     t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
End If

If t_fechavto = "" Then
  t_fechavto = Format$(Now, "dd/mm/yyyy")
Else
  If Not IsDate(t_fechavto) Then
     t_fechavto = Format$(Now, "dd/mm/yyyy")
  End If
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
  Call CALCULATOTALES
  Frame2.Enabled = True
  t_cotizacion.SetFocus
End If

If KeyCode = vbKeyInsert Then
   vta_COMPVARIOS1.t_renglon = ""
   vta_COMPVARIOS1.t_cantidad = ""
   vta_COMPVARIOS1.t_pu = ""
   vta_COMPVARIOS1.t_importe = ""
   vta_COMPVARIOS1.Show
End If
End Sub

Sub graba()
  ' On Error GoTo ERRORGRABA
  numint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(c_sucursal)
  cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
  cl_compvta.letra = t_letra
  cl_compvta.numcomp = Val(t_numcomp)
      
      If Option1 = True Then
         ep = "N"
         cp = "0000-00000000"
         contado = "N"
         If Option4 = True Then
           ssi = Val(t_total)
         Else
           ssi = Val(T_total2)
         End If
      Else
         ep = "S"
         cp = "ctdo"
         contado = "S"
         'cl_compvta.ctacte = "N"
         ssi = 0
      End If
      
      cl_compvta.ACTUALIZA_NUMERADOR
      
      If Option4 = True Then
        moneda = "P"

      Else
        moneda = "D"
      End If
      
      
      
      Set rs = New ADODB.Recordset
      q = "select * from g8 where [id_actividad] = " & c_actividad.ItemData(c_actividad.ListIndex)
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
       codact = rs("id_actividad")
       alicuotaib = rs("alicuota_ib")
      Else
       codact = 0
       alicuotaib = 0
      End If
      Set rs = Nothing

      tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
      
      cn1.BeginTrans
      
      If Check1 = 0 Then
        tom = Val(T_total2)
      Else
        tom = 0
      End If
      
      'saco numz
      Set rs = New ADODB.Recordset
      q = "select * from fsc_001"
      rs.Open q, cn1
      numz = 0
      While Not rs.EOF
        If rs("sucursal_fiscal") = Val(t_sucursal) Then
          numz = rs("ult_z") + 1
        End If
        rs.MoveNext
      Wend
      Set rs = Nothing
      
      fvcae = Format$(DateValue(t_fecha) + 10, "dd/mm/yyyy")
      
      QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total], [estado], " & _
" [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], [VENTA], [CONTADO], " & _
" [id_actividad], [alicuota_ib],[alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos], [valor_declarado], [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], " & _
" [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], [dominio_acoplado02], [SALDO_IMPAGO02], [id_camion02], [dni_chofer02], [num_z], [cae], [cae_vence], [tipo_op], [perc_ib], [numint_asociado])"

      
QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & ", " & _
c_prov.ItemData(c_prov.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotal) & ", " & Val(t_nograbado) & ", " & Val(t_iva) & ", " & Val(t_total) & ", 'A', " & _
para.cuenta_ventas & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & t_observaciones & " ', " & Val(t_cotizacion) & ", " & _
tom & ", '" & moneda & "'," & c_vend.ItemData(c_vend.ListIndex) & ", '" & cl_compvta.venta & "', '" & contado & "', " & codact & ", " & alicuotaib & ", 0, 0, '" & t_fechavto & _
"', 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & ", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & Left$(vta_clientes.t_localidad, 50) & "', " & tiporespiva & _
", ' ', ' ', ' ', " & ssi & ", 1, 0, " & numz & ", '" & t_cae & "', '" & fvcae & "', 1, " & Val(t_perc) & ",0)"
      
      
     'MsgBox (QUERY2)
      cn1.Execute QUERY
      
      For i = 1 To msf1.Rows - 1
        If Val(msf1.TextMatrix(i, 1)) > 1 Then
          Set cl_prod = New productos
          cl_prod.cargar (Val(msf1.TextMatrix(i, 1)))
          costo = cl_prod.precio_ult_compra
          Set cl_prod = Nothing
        Else
          costo = 0
        End If
        
        QUERY = "INSERT INTO vta_03([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [pu], [importe], [tasaiva], [impuesto], [costo])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & " ', " & Val(msf1.TextMatrix(i, 3)) & ", " & Val(msf1.TextMatrix(i, 4)) & ", " & Val(msf1.TextMatrix(i, 6)) & ", " & Val(msf1.TextMatrix(i, 5)) & ", 0, " & costo & ")"
        cn1.Execute QUERY
      
        If cl_compvta.STOCK <> "N" Then
           QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo])"
           QUERY = QUERY & " VALUES ('" & t_fecha & "', " & Val(msf1.TextMatrix(i, 1)) & ", " & msf1.TextMatrix(i, 3) & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.abreviatura & t_letra & Format$(t_sucursal, "0000") & "-" & Format$(t_numcomp, "00000000") & "', '" & Left$(c_prov, 50) & "', " & numint & ",'V'" & ")"
           cn1.Execute QUERY
          
           If cl_compvta.STOCK = "E" Then
             c = Val(msf1.TextMatrix(i, 3))
           Else
             c = -Val(msf1.TextMatrix(i, 3))
           End If
           q = "update a2 set [stock] = [stock] + " & c & " where [id_producto] = " & Val(msf1.TextMatrix(i, 1))
           cn1.Execute q
        
        End If
        
        If cl_compvta.venta <> "N" Then
           ultvta = t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & " | " & Left$(c_prov, 25) & " | " & t_fecha & " | " & Format$(Val(msf1.TextMatrix(i, 4)), "#####0.00")
           QUERY = "update a2 set  [ultima_venta]='" & ultvta & "'"
           QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
           cn1.Execute QUERY
        End If
      Next i
      
      
      'actualizo tasa de iva
      If cl_compvta.grabado <> "N" Then
       If verificatasaunica Then
         If msf1.Rows > 1 Then
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
          QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(1, 5)) & ", " & Val(t_iva) & ", " & Val(t_subtotal) & ", " & tiporespiva & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ")"
          cn1.Execute QUERY
         Else
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
          QUERY = QUERY & " VALUES (" & numint & ", 21, " & Val(t_iva) & ", " & Val(t_subtotal) & ", " & tiporespiva & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ")"
          cn1.Execute QUERY
         End If
       Else
        For i = 1 To 7
        If Val(vta_facturacion2.msf1.TextMatrix(i, 1)) > 0 Then
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
          QUERY = QUERY & " VALUES (" & numint & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 0)) & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 2)) & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 1)) & ", " & tiporespiva & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ")"
          cn1.Execute QUERY
        End If
       Next i
      End If
     End If
      
      
     'actualizo percepciones
     If Val(t_perc) > 0 Then
        For i = 1 To ABM_COMP_COMPRA2.msf1.Rows - 1
          QUERY = "INSERT INTO vta_016([num_int], [secuencia], [id_percepcion], [importe], [id_cuenta], [cod_regimen])"
          QUERY = QUERY & " VALUES (" & numint & ", " & i & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 1) & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 3) & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 4) & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 6) & ")"
          cn1.Execute QUERY
        Next i
      End If
      ABM_COMP_COMPRA2.armagrid
       
      
          
     If Option2 = True Then
        'cobranza en efectivo
        'grabo mov caja
        Set rs = New ADODB.Recordset
        q = "select * from cyb_01 where [id_forma_pago] = 1"
        rs.Open q, cn1
        If Not rs.BOF And Not rs.EOF Then
          If rs("caja") = "S" Then
            If cl_compvta.ctacte = "H" Then
              t = -Val(t_total)
            Else
              t = Val(t_total)
            End If
              
            'grabo mov caja
             QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [id_usuario])"
             QUERY = QUERY & " VALUES (" & rs("id_cuenta_cont") & ", " & para.cuenta_deudores & ", '" & Left$(c_prov, 50) & "', " & t & ", 'D', '" & t_fecha & "', " & numint & ", 'V', 'Cdo. " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 1, " & para.id_usuario & ")"
             cn1.Execute QUERY
          End If
         End If
         Set rs = Nothing
      
        
        
      End If
      

    If Generaasientosauto Then
     If cl_compvta.contabilidad <> "N" And c_cuenta.ListIndex > 0 Then
         numintcgr = saca_ultnumero_int_comp("G")

         If Option1 = True Then
           cta = para.cuenta_deudores
         Else
           cta = para.cuenta_caja
         End If
         u1 = cl_compvta.contabilidad
          
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         
         Set rs = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & cta
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Ventas] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'V', " & numint & ", " & Val(t_total) & ", " & Val(t_total) & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_prov), 50) & "')"
         cn1.Execute QUERY
      
         
         'cuenta madre ctacte o caja
         ic = 1
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & Val(t_total) & ", '" & dcta & "')"
         
         cn1.Execute QUERY
         ic = ic + 1
      
         If Val(t_nograbado) > 0 Then
           'cuenta nogbra
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_conceptos_nograbados & ", '" & u2 & "', " & Val(t_nograbado) & ", 'No Grabado')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
                   
         If Val(t_perc) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_conceptos_nograbados & ", '" & u2 & "', " & Val(t_perc) & ", 'Perc. y Ret')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
          
         If Val(t_iva) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_ventas & ", '" & u2 & "', " & Val(t_iva) & ", 'IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         'contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & u2 & "', " & Val(t_subtotal) & ", '" & "Comp. Varios Vta." & "')"
         cn1.Execute QUERY
      
        End If
     End If
      
      'actualizo remitos
     'If Val(vta_selremitos.t_seleccionados) > 0 Then
        For i = 1 To vta_selremitos.msf1.Rows - 1
          If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
             QUERY = "INSERT INTO vta_08([id_factura], [id_remito])"
             QUERY = QUERY & " VALUES (" & numint & ", " & Val(vta_selremitos.msf1.TextMatrix(i, 4)) & ")"
             cn1.Execute QUERY
          End If
        Next i
     
        Call actualizaremitos
      
      
      
     If EXISTE = "S" Then
       QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
       QUERY = QUERY & " VALUES ('Modif. Comprobante manual NI " & numint & "', " & para.id_usuario & ", 'V'," & numint & ", '" & Now & "', '" & Left$(c_prov, 50) & "', 6, " & c_prov.ItemData(c_prov.ListIndex) & ")"
     Else
       QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
       QUERY = QUERY & " VALUES ('Ingreso Comprobante manual NI " & numint & "', " & para.id_usuario & ", 'V'," & numint & ", '" & Now & "', '" & Left$(c_prov, 50) & "', 7, " & c_prov.ItemData(c_prov.ListIndex) & ")"
     End If
    cn1.Execute QUERY
    cn1.Execute QUERY
      
      
      cn1.CommitTrans
           
       Set cl_compvta = Nothing
      Set cl_cli = Nothing
      Call INICIALIZA2(Me)
      Call armagrid
      c_tipocomp.SetFocus
   

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    vta_COMPVARIOS1.t_renglon = msf1.Row
    vta_COMPVARIOS1.t_basico = msf1.TextMatrix(msf1.Row, 1)
    vta_COMPVARIOS1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    vta_COMPVARIOS1.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
    vta_COMPVARIOS1.t_pu = msf1.TextMatrix(msf1.Row, 4)
    vta_COMPVARIOS1.t_importe = msf1.TextMatrix(msf1.Row, 6)
    vta_COMPVARIOS1.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True

End Sub

Private Sub Option3_Click()
Label13 = "Total $"
End Sub

Private Sub Option4_Click()
Label13 = "Total U$s"
End Sub

Private Sub t_cotizacion_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_cotizacion_LostFocus()
If Val(t_cotizacion) <= 0 Then
   t_cotizacion = 1
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

Private Sub t_iva_LostFocus()
Call sacatotales

End Sub

Private Sub t_nograbado_LostFocus()
Call sacatotales

End Sub


Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub

Private Sub t_numcomp_LostFocus()
   t_numcomp = Format$(Val(t_numcomp), "00000000")
   Call carga
End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub

Private Sub t_perc_LostFocus()
Call sacatotales

End Sub

Private Sub t_subtotal_LostFocus()
Call sacatotales
End Sub
Sub sacatotales()
t_subtotal = Format$(Val(t_subtotal), "######0.00")
t_nograbado = Format$(Val(t_nograbado), "######0.00")
t_perc = Format$(Val(t_perc), "######0.00")
t_iva = Format$(Val(t_iva), "######0.00")
t_total = Format$(Val(t_subtotal) + Val(t_nograbado) + Val(t_perc) + Val(t_iva), "######0.00")
If Option4 = True Then
  T_total2 = Format$(Val(t_total) / Val(t_cotizacion), "#####0.00")
Else
  T_total2 = Format$(Val(t_total) * Val(t_cotizacion), "#####0.00")
End If
End Sub

Private Sub t_sucursal_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_sucursal_LostFocus()
t_sucursal = Format$(Val(t_sucursal), "0000")
  Call inicia
End Sub

Private Sub t_total_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_total_LostFocus()
t_total = Format$(t_total, "######0.00")
End Sub

Private Sub Text1_LostFocus()
If Not IsDate(t_fechavto) Then
  t_fechavto = Format$(Now, "dd/mm/yyyy")
Else
  t_fechavto = Format$(t_fechavto, "dd/mm/yyyy")
End If
End Sub
