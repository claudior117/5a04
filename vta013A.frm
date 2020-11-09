VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0A6BE9FC-5039-11D5-98EC-0800460222F0}#1.0#0"; "IFEpson.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_remitos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "REMITOS"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   9120
      TabIndex        =   63
      Top             =   6840
      Width           =   2655
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprime Descripcion Extra"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
      End
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal epson1 
      Left            =   11400
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   10080
      TabIndex        =   56
      Top             =   960
      Width           =   1455
      Begin VB.CommandButton Command4 
         Caption         =   "Transporte"
         Height          =   615
         Left            =   120
         Picture         =   "vta013A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contrato de Consignacion"
      Height          =   975
      Left            =   9120
      TabIndex        =   50
      Top             =   5880
      Width           =   2655
      Begin VB.TextBox t_fechac 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   52
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   10080
      TabIndex        =   40
      Top             =   0
      Width           =   1455
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   9600
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox t_cantlineas 
         Enabled         =   0   'False
         Height          =   405
         Left            =   960
         MaxLength       =   3
         TabIndex        =   62
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox t_letrafact 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   60
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   38
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
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales del Comprobante"
      Height          =   2295
      Left            =   240
      TabIndex        =   30
      Top             =   6000
      Width           =   8775
      Begin VB.TextBox t_valor 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_bultos 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_alicuotaib 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   16
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   7800
         Picture         =   "vta013A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox t_observaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox T_total2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   44
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox t_cotizacion 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox t_total 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox t_iva 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox t_perc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox t_nograbado 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox t_subtotal 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Valor Declarado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   55
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Total Bultos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   54
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "% Perc."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Total U$s"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7440
         TabIndex        =   45
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Cotizacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   35
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5040
         TabIndex        =   34
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Perc.IB"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "No Grabado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Subtotal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   3735
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   240
      TabIndex        =   25
      Top             =   0
      Width           =   9735
      Begin VB.TextBox t_enviara 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   7
         Top             =   1560
         Width           =   7215
      End
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "vta013A.frx":040F
         Left            =   7680
         List            =   "vta013A.frx":0411
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox t_stock 
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
         Left            =   9000
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   8880
         Picture         =   "vta013A.frx":0413
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   8040
         Picture         =   "vta013A.frx":0785
         Style           =   1  'Graphical
         TabIndex        =   47
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
         TabIndex        =   20
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "vta013A.frx":088A
         Left            =   2160
         List            =   "vta013A.frx":088C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1200
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
         Left            =   3480
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
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
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Text            =   "c_prov"
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   65
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         TabIndex        =   61
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Stock:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8160
         TabIndex        =   59
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Enviar a:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   22
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta013A.frx":088E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta013A.frx":1110
         Style           =   1  'Graphical
         TabIndex        =   23
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
      TabIndex        =   21
      Top             =   8340
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
            TextSave        =   "25/4/2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:08"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_remitos"
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

Sub renumera()
r = 1
For i = 1 To msf1.Rows - 1
 If Val(msf1.TextMatrix(i, 0)) <> 0 Then
    msf1.TextMatrix(i, 0) = r
    r = r + 1
 End If
Next i


End Sub


Sub CALCULATOTALES()
 s = 0
  v = 0
  b = 0
  For i = 1 To msf1.Rows - 1
      b = b + Val(msf1.TextMatrix(i, 7))
      r = Val(msf1.TextMatrix(i, 8))
      s = s + r
      v = v + (r * (Val(msf1.TextMatrix(i, 6)) / 100))
  Next i
  t_subtotal = s
  t_iva = v
  t_bultos = b
  If Option3 = True Then
    t_valor = Format$((s + v) * Val(t_cotizacion), "#######0.00")
  Else
     t_valor = Format$((s + v), "#######0.00")
  End If
    Call sacatotales
      

End Sub
Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
   Option1 = True
   vta_transporte.limpia
End Sub
Sub carga()
  
  Set rs = New ADODB.Recordset
  q = " select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     t_fecha = rs("fecha")
     c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select * from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad_original") & Chr(9) & rs1("tunidad") & Chr$(9) & Format$(rs1("pu"), "######0.00") & Chr(9) & rs1("tasaiva") & Chr(9) & rs1("bultos") & Chr$(9) & rs1("importe") & Chr$(9) & rs1("cantidad") & Chr$(9) & rs1("pu_final")
        rs1.MoveNext
     Wend
     Set rs1 = Nothing
     c_vend.ListIndex = buscaindice(c_vend, rs("id_vendedor"))
     t_subtotal = Format$(rs("subtotal"), "######0.00")
     t_nograbado = Format$(rs("impuestos"), "######0.00")
     t_perc = Format$(rs("perc_iva") + rs("perc_gan") + rs("perc_ib"), "######0.00")
     t_iva = Format$(rs("iva"), "######0.00")
     t_total = Format$(rs("total"), "######0.00")
     t_bultos = rs("total_bultos")
     t_valor = rs("valor_declarado")
     t_observaciones = rs("observaciones")
     If rs("moneda") = "P" Then
       Option4 = True
     Else
       Option3 = True
     End If
     t_stock = rs("stock")
     vta_transporte.t_transp = rs("transporte")
     vta_transporte.t_direccion = rs("direccion_transp")
    ' vta_transporte.t_localidad = rs("localidad_trans")
     vta_transporte.t_cuit = rs("cuit_transp")
     vta_transporte.t_chofer = rs("chofer02")
     vta_transporte.t_dominio = rs("dominio02")
     vta_transporte.t_acoplado = rs("dominio_acoplado02")
     Set rs = Nothing
  Else
     EXISTE = "N"
  End If
  
End Sub

Private Sub btnacepta_Click()
J = MsgBox("Graba Comprobante ", 4)
If J = 6 Then
 Call renumera
 If verificaperiodog(t_fecha) = "A" Then
  If Val(t_sucursal) = glo.sucursalf Then
     Call fiscal
  Else
     Call normal
  End If
 Else
  MsgBox ("Periodo cerrado. Imposible grabar comprobante")
 End If
End If





End Sub
Sub normal()
  Set rs = New ADODB.Recordset
  q = "select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
      EXISTE = "S"
      If para.id_grupo_modulo_actual >= 8 Then
         ni = rs("num_int")
         Set rs = Nothing
         J = MsgBox("Remito existente. ¿Desea Modificarlo?. Recuerde que se perderan facturaciones parciales o totales del comprobante", 4)
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
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 11
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 4000
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 800
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 800
msf1.ColWidth(7) = 800
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 1000
msf1.ColWidth(10) = 1000

msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "Unidad"
msf1.TextMatrix(0, 5) = "P.U."
msf1.TextMatrix(0, 6) = "% Iva"
msf1.TextMatrix(0, 7) = "Bultos"
msf1.TextMatrix(0, 8) = "Importe"
msf1.TextMatrix(0, 9) = "Remitido"
msf1.TextMatrix(0, 10) = "PU Final"


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
    MsgBox ("Imposible realizar Remitos a Cliente CONTADO")
Else
    Call iniciacli
End If

End Sub
Sub inicia()
Set cl_cli = New Clientes
cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
If cl_cli.id > 0 Then
   t_letra = "X"
   
   c_vend.ListIndex = buscaindice(c_vend, cl_cli.idvendedor)
   Set cl_compvta = New comprobantes_venta
   cl_compvta.sucursal = Val(t_sucursal)
   cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
   cl_compvta.letra = t_letra
   cl_compvta.SACANUMCOMP
   t_cantlineas = cl_compvta.cant_lineas
   t_numcomp = Format$(cl_compvta.numcomp, "00000000")
   t_stock = cl_compvta.STOCK
   Set cl_compvta = Nothing
   t_cotizacion = para.cotizacion
Else
  MsgBox ("Error. No se puedo Inicializa el Cliente")
End If
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
q = "select [imprime_desc_extra] from vta_06 where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex)
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

Private Sub Command4_Click()
vta_transporte.Show
End Sub

Private Sub Command5_Click()
vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
vta_clientes.carga
vta_clientes.Show

End Sub

Private Sub Form_Activate()
Frame2.Enabled = False

End Sub
Sub iniciacli()
 If c_prov.ListIndex >= 0 Then
   vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
   vta_clientes.carga
   t_enviara = vta_clientes.t_dirlocal
 Else
   If Val(vta_clientes.t_id) <> 0 Then
      vta_clientes.t_id = 0
      vta_clientes.limpia
   
   
   End If
 End If
 
End Sub



Sub fiscal()
Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
If cl_fiscal.imprimerto = "S" Then
  seguir = True
  While seguir
    If imprime_remitofiscal2 Then
        espere.ProgressBar1.Value = 5
        espere.Label1 = "Espere... Grabando Comprobante Fiscal"
        Call graba
        seguir = False
    Else
        J = MsgBox("Error al Imprimir el Comprobante. Verifique la Impresora para continuar.  Reintente o Cancele", 5)
        If J = 4 Then
           seguir = True
        Else
           seguir = False
        End If
    End If
    Unload espere
  Wend
Else
  seguir = True
  While seguir
    If imprime_remitofiscal Then
        espere.ProgressBar1.Value = 5
        espere.Label1 = "Espere... Grabando Comprobante Fiscal"
        Call graba
        seguir = False
    Else
        J = MsgBox("Error al Imprimir el Comprobante. Verifique la Impresora para continuar.  Reintente o Cancele", 5)
        If J = 4 Then
           seguir = True
        Else
           seguir = False
        End If
    End If
    Unload espere
  Wend
 End If
End Sub
Function imprime_remitofiscal2()
Dim CUIT As String
Dim identifica As String
Dim tpago As String
Dim t As String
Dim tipocompf As String
Dim cliz As String
Dim dirz As String
Dim direz As String
Dim locz As String
Dim de1z As String
Dim uz As String
espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 6
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Comprobando Impresora"

Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
tipocompf = cl_fiscal.CODRTO
caracteresmax = cl_fiscal.caracteresmax
Set cl_fiscal = Nothing

If vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex) <> 3 Then
   identifica = "CUIT"
   CUIT = vta_clientes.t_cuit
 Else
   identifica = "DNI"
   CUIT = "0"
 End If
 
tpago = "Cta.Cte. Nro. " & Format$(Val(vta_clientes.t_id), "00000")


espere.ProgressBar1.Value = 2
espere.Label1 = "Espere... Abriendo Comprobante Fiscal:" & c_tipocomp
 
cliz = textofiscal(Left$(vta_clientes.t_cli & "-", caracteresmax))
dirz = textofiscal(Left$(vta_clientes.t_direccion & "-", caracteresmax))
locz = textofiscal(Left$(vta_clientes.t_localidad, caracteresmax))
direz = textofiscal(Left$(t_enviara, caracteresmax))

r = epson1.OpenInvoice(tipocompf, "C", "X", "1", "P", "17", "I", vta_clientes.t_codfiscal, cliz, " ", identifica, CUIT, "N", dirz, locz, direz, tpago, " ", "C")
 
 
 'envia items a facturar
espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Imprimiendo Productos"
 
i = 1
uz = "    "
While r And i < msf1.Rows
      If r Then
         uz = Left$(Format$(RTrim$(msf1.TextMatrix(i, 4)), "@@@@!"), 4)
         de1z = textofiscal(Left$(uz & " " & msf1.TextMatrix(i, 2), caracteresmax))
         r = epson1.SendInvoiceItem(de1z, Format$(Val(msf1.TextMatrix(i, 3)) * 1000, "00000000"), Format$(Val(msf1.TextMatrix(i, 5)) * 100, "000000000"), Format$(Val(msf1.TextMatrix(i, 6)) * 100, "0000"), "M", "0", "0", " ", " ", " ", "0", "0")
      Else
        i = msf1.Rows
      End If
   i = i + 1
 Wend
 
 'no realiza pago copmprobante y automaticanmene cuando cierro se genera un pago pr el total
  
 'subtotal para obtener el importe neto, iva y total impreso en la factura
espere.ProgressBar1.Value = 4
espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

If r Then r = epson1.CloseInvoice("E", "X", " ")
  
If r Then t_numcomp = epson1.AnswerField_3

imprime_remitofiscal2 = r
    
   'si hay error tratarlo en un proceso global de errores fiscales

End Function

Function imprime_remitofiscal() As Boolean
Dim CUIT As String
Dim identifica As String
Dim tpago As String

Set cl_cli = New Clientes
cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))

espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 5
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Abriendo Comprobante Fiscal"

'abrir nc

r = epson1.OpenNoFiscal
If r Then r = epson1.SendNoFiscalText(" ")
If r Then r = epson1.SendNoFiscalText("      ****" & c_tipocomp & " ****    ")
If r Then r = epson1.SendNoFiscalText(" Documento No Valido como Factura")
If r Then r = epson1.SendNoFiscalText(" ")
If r Then r = epson1.SendNoFiscalText("Nro: " & t_letra & "  " & t_sucursal & "-" & t_numcomp)
If r Then r = epson1.SendNoFiscalText(" ")
If r Then r = epson1.SendNoFiscalText("Cliente: " & cl_cli.razonsocial)
If r Then r = epson1.SendNoFiscalText("         (" & Format$(cl_cli.id, "00000") & ")")
If r Then r = epson1.SendNoFiscalText("Domicilio: " & cl_cli.direccion)
If r Then r = epson1.SendNoFiscalText("Localidad: " & Format$(Left$(cl_cli.direccion, 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!") & "   Cuit:" & cl_cli.abreviatura_tipoiva & " " & cl_cli.CUIT)
If r Then r = epson1.SendNoFiscalText("Cuit: " & cl_cli.abreviatura_tipoiva & "  " & cl_cli.CUIT)
If r Then r = epson1.SendNoFiscalText(" ")
If r Then r = epson1.SendNoFiscalText("----------------------------------------")
If r Then r = epson1.SendNoFiscalText("Basico  Detalle                   Cant ")
If r Then r = epson1.SendNoFiscalText("----------------------------------------")
 'envia items al nc
espere.ProgressBar1.Value = 2
espere.Label1 = "Espere... Imprimiendo Productos"
 
 i = 1
 While r And i < msf1.Rows
      If r Then
         r = epson1.SendNoFiscalText(Format$(msf1.TextMatrix(i, 1), "00000") & " " & Format$(Left$(msf1.TextMatrix(i, 2), 40), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!") & " " & msf1.TextMatrix(i, 3))
      Else
        i = msf1.Rows
      End If
   i = i + 1
 Wend
 
 If r Then r = epson1.SendNoFiscalText(" ")
 
 espere.ProgressBar1.Value = 3
 espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

 
  If r Then r = epson1.CloseNoFiscal
   
  imprime_remitofiscal = r
    
  Set cl_cli = Nothing
   'si hay error tratarlo en un proceso global de errores fiscales
   
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call TabEnter2(Me, 19)
End If


End Sub

Private Sub Form_Load()

Call INICIALIZA2(Me)
Call carga_clientes(c_prov)
c_prov.RemoveItem 0
c_prov.ListIndex = 0
Call carga_SUCURSALES(c_sucursal)
c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)


Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & glo.sucursal & " and  [id_tipocomp] > 40  and [id_tipocomp] < 50 "
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_tipocomp", c_tipocomp, True)
Set rs = Nothing
c_tipocomp.ListIndex = buscaindice(c_tipocomp, 45)

Set rs = New ADODB.Recordset
q = "select * from vta_05 order by [denominacion]"
rs.Open q, cn1
Call llena_combo(rs, "denominacion", "id_vendedor", c_vend, True)
Set rs = Nothing
c_vend.ListIndex = 0
Call armagrid
Call barraesag(Me)
If para.moneda = "P" Then
  Option4 = True
Else
  Option3 = True
End If
Load vta_remitos1
Load vta_transporte
vta_transporte.limpia
calcula_perc_ib = "N"
t_alicuotaib = "0.00"
minimo_perc_ib = 0
t_sucursal = Format$(glo.sucursal, "0000")
Set rs = Nothing

Load vta_clientes

Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
epson1.PortNumber = cl_fiscal.puerto
Set cl_fiscal = Nothing


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_remitos1
Unload vta_transporte
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F3] Lote/Vto - [F5] Elimina - [F9] Graba"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If

If KeyCode = vbKeyF3 Then
 If msf1.Row > 0 Then
    g = InputBox$("Agregue Nro. Lote", "LOTE/VTO", "LOTE: ")
    If g <> "" Then
      msf1.TextMatrix(msf1.Row, 2) = msf1.TextMatrix(msf1.Row, 2) & " " & g
    End If
 End If
End If


If KeyCode = vbKeyF9 Then
  Call renumera
  Call sacatotales

  Frame2.Enabled = True
  c_vend.SetFocus
End If

If KeyCode = vbKeyInsert Then
   vta_remitos1.limpia
   vta_remitos1.Show
End If

If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Sub graba()
  'On Error GoTo ERRORGRABA
  numint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(t_sucursal)
  cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
  cl_compvta.letra = t_letra
  cl_compvta.numcomp = Val(t_numcomp)
      
  If cl_compvta.idtipocomp = 45 Then
      ep = "S"
      cp = "Rto"
      tc = "R"
      estado = "S" 'sin facturar
  Else
      ep = "S"
      cp = "Dev"
      tc = "D"
      estado = "F" 'ver???
  End If
      
     
     contado = "N"
     
     'cl_compvta.ctacte = "N"
      
     cl_compvta.sucursal = Val(t_sucursal)
     
     If EXISTE = "N" Then
          cl_compvta.ACTUALIZA_NUMERADOR
     End If
     
     If Option4 = True Then
        moneda = "P"
      Else
        moneda = "D"
      End If
      
      Set cl_cli = New Clientes
      cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
       
      cn1.BeginTrans
      
      If vta_transporte.c_camion.ItemData(vta_transporte.c_camion.ListIndex) <= 0 Then
        idcamion = 1
      Else
        idcamion = vta_transporte.c_camion.ItemData(vta_transporte.c_camion.ListIndex)
      End If
      
      
      QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total], [estado], [id_cuenta], [stock], [cta_cte], [grabado], " & _
" [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva], [total_bultos], [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [id_transporte], [fecha_vto], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], [dominio_acoplado02], [id_camion02], [dni_chofer02], " & _
"[cae], [cae_vence], [tipo_op])"

QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & ", " & c_prov.ItemData(c_prov.ListIndex) & _
", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotal) & ", " & Val(t_nograbado) & ", " & Val(t_iva) & ", " & Val(t_total) & ", '" & estado & "', " & para.cuenta_ventas & ", '" & t_stock & "', '" & _
cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & t_observaciones & " ', " & Val(t_cotizacion) & ", " & Val(T_total2) & ", '" & moneda & "', " & c_vend.ItemData(c_vend.ListIndex) & _
", '" & cl_compvta.venta & "', '" & contado & "', " & Val(t_perc) & ", 0, 0, " & Val(t_bultos) & ", " & Val(t_valor) & ", '" & vta_transporte.t_transp & "', '" & vta_transporte.t_direccion & "', '" & vta_transporte.t_cuit & _
"', " & Val(vta_transporte.t_id) & ", '" & t_fecha & "', 0, " & Val(c_sucursal) & ", '" & Left$(cl_cli.razonsocial, 50) & "', '" & _
Left$(cl_cli.direccion, 50) & "', '" & Left$(cl_cli.CUIT, 20) & "', '" & Left$(cl_cli.localidad, 50) & "', " & cl_cli.idtipoiva & ", '" & vta_transporte.t_chofer & "', '" & vta_transporte.t_dominio & "', '" & vta_transporte.t_acoplado & _
"', " & idcamion & ", " & Val(vta_transporte.t_dni) & ",'u2','01/01/2018',2)"


      cn1.Execute QUERY
      
      Set cl_cli = Nothing
      For i = 1 To msf1.Rows - 1
        If Val(msf1.TextMatrix(i, 1)) > 1 Then
          Set cl_prod = New productos
          cl_prod.cargar (Val(msf1.TextMatrix(i, 1)))
          costo = cl_prod.precio_ult_compra
          Set cl_prod = Nothing
        
        
            'descargar de remitos`pendientyes las cantidaddes de la devolucion
          If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 46 Then
            q = "SELECT * FROM VTA_02, VTA_03 WHERE VTA_02.[NUM_INT] = VTA_03.[NUM_INT] AND [ID_TIPOCOMP] = 45 AND [ID_PRODUCTO] = " & Val(msf1.TextMatrix(i, 1)) & " AND [CANTIDAD] > 0 and [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex) & " and [estado] = 'S'"
            q = q & " order by [fecha]"
            Set rs2 = New ADODB.Recordset
            rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
            c = Val(msf1.TextMatrix(i, 3))
            While Not rs2.EOF And c > 0
              If rs2("cantidad") >= c Then
               rs2("cantidad") = rs2("cantidad") - c
               rs2.Update
               c = 0
              Else
               c = c - rs2("cantidad")
               rs2("cantidad") = 0
               rs2.Update
              End If
              rs2.MoveNext
             Wend
             Set rs2 = Nothing
           End If
        Else
          costo = 0
        End If
        
        QUERY = "INSERT INTO vta_03([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [pu], [importe], [tasaiva], [impuesto], [costo], [cantidad_original], [tunidad], [bultos], [pu_final])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & " ', " & Val(msf1.TextMatrix(i, 3)) & ", " & Val(msf1.TextMatrix(i, 5)) & ", " & Val(msf1.TextMatrix(i, 8)) & ", " & Val(msf1.TextMatrix(i, 6)) & ", 0, " & costo & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & Left$(msf1.TextMatrix(i, 4), 8) & "', " & Val(msf1.TextMatrix(i, 7)) & ", " & Val(msf1.TextMatrix(i, 10)) & ")"
        cn1.Execute QUERY
      
        'productos pendientes de facturacion
        QUERY = "INSERT INTO vta_07([num_int], [secuencia], [id_producto], [id_cliente], [cantidad], [tipo])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", " & c_prov.ItemData(c_prov.ListIndex) & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & tc & "')"
        cn1.Execute QUERY
      
        
        
        
             
        
        If t_stock <> "N" Then
           QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo], [id_cliente])"
           QUERY = QUERY & " VALUES ('" & t_fecha & "', " & Val(msf1.TextMatrix(i, 1)) & ", " & msf1.TextMatrix(i, 3) & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.abreviatura & t_letra & Format$(t_sucursal, "0000") & "-" & Format$(t_numcomp, "00000000") & _
           "', '" & Left$(c_prov, 50) & "', " & numint & ",'V', " & c_prov.ItemData(c_prov.ListIndex) & ")"
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
           ultvta = t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & " | " & Left$(c_prov, 28) & " | " & t_fecha & " | " & Format$(Val(msf1.TextMatrix(i, 4)), "#####0.00")
           QUERY = "update a2 set  [ultima_venta]='" & ultvta & "'"
           QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
           cn1.Execute QUERY
        End If
      Next i
      
      
    
    'verifica los remitos que debn pasar a pendiente
    If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 46 Then
       q = "SELECT * FROM VTA_02 WHERE [ID_TIPOCOMP] = 45 AND  [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex) & " and [estado] = 'S'"
       q = q & " order by [fecha]"
       Set rs2 = New ADODB.Recordset
       rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
       While Not rs2.EOF
         Set rs3 = New ADODB.Recordset
         q = "select * from vta_03 where [num_int] = " & rs2("num_int")
         rs3.Open q, cn1
         p = 0
         While Not rs3.EOF
           If rs3("cantidad") > 0 Then
             p = 1
           End If
           rs3.MoveNext
         Wend
         Set rs3 = Nothing
         If p = 0 Then
           rs2("estado") = "F"
           rs2.Update
         End If
         rs2.MoveNext
        Wend
        Set rs2 = Nothing
     End If
       



     If cl_compvta.contabilidad <> "N" Then
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
      
         ic = 1
         'cuenta madre ctacte o caja
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
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_perc_IB & ", '" & u2 & "', " & Val(t_perc) & ", 'Perc. IB')"
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
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_ventas & ", '" & u2 & "', " & Val(t_subtotal) & ", '" & "Ventas" & "')"
         cn1.Execute QUERY
      
      End If
      
      cn1.CommitTrans
      Set rs = Nothing
      Set cl_compvta = Nothing
      Set cl_cli = Nothing

      If Val(t_sucursal) <> glo.sucursalf Then
       If glo.sucursalf = 0 Then
        J = MsgBox("Confirma Impresion del Comprobante", 4)
        If J = 6 Then
           Set cl_compvta = New comprobantes_venta
           cl_compvta.cargar2 (numint)
           cl_compvta.imprimir
        End If
       Else
         MsgBox ("Por disposicion del AFIP teniendo una impresora fiscal definida no se permite imprimir otro tipo de comprobantes. Gracias")
       End If
      End If
      Call INICIALIZA2(Me)
      Call armagrid
      t_sucursal = c_sucursal
      c_tipocomp.SetFocus
       
  
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos y vuelva a repetir la operacion")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    vta_remitos1.t_renglon = msf1.Row
    vta_remitos1.t_basico = msf1.TextMatrix(msf1.Row, 1)
    vta_remitos1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    vta_remitos1.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
    vta_remitos1.t_unidad = msf1.TextMatrix(msf1.Row, 4)
    vta_remitos1.t_pu = msf1.TextMatrix(msf1.Row, 5)
    vta_remitos1.t_bultos = msf1.TextMatrix(msf1.Row, 7)
    vta_remitos1.t_importe = msf1.TextMatrix(msf1.Row, 8)
    
    vta_remitos1.Show
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

Private Sub t_alicuotaib_LostFocus()
If Val(t_alicuotaib) < 0 Then
  t_alicuotaib = "0.00"
End If

End Sub

Private Sub t_cotizacion_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_cotizacion_LostFocus()
If Val(t_cotizacion) <= 0 Then
   t_cotizacion = 1
End If
End Sub

Private Sub t_enviara_LostFocus()
If Len(t_enviara) <= 0 Then
  MsgBox ("El campo [Enviar a] no puede estar vacio")
  t_enviara.SetFocus
End If
End Sub

Private Sub t_fecha_GotFocus()
If Val(t_sucursal) = glo.sucursalf Then
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


Private Sub t_iva_LostFocus()
Call sacatotales

End Sub

Private Sub t_nograbado_LostFocus()
Call sacatotales

End Sub


Private Sub t_numcomp_GotFocus()
Call inicia

End Sub

Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub

Private Sub t_numcomp_LostFocus()
If IsNumeric(t_numcomp) Then
   t_numcomp = Format$(t_numcomp, "00000000")
   Call carga
   Call iniciacomp
End If
End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub

Private Sub t_perc_LostFocus()
Call sacatotales

End Sub
Sub iniciacomp()
Set rs = New ADODB.Recordset
q = "select [imprime_desc_extra] from vta_06 where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex)
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
End Sub

Private Sub t_stock_LostFocus()
t_stock = Format$(t_stock, ">@")
Select Case t_stock
   Case Is = "S", Is = "N", Is = "E"
      
   Case Else
   t_stock = "N"
End Select
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
't_valor = t_total
t_bultos = Format$(Val(t_bultos), "####0")
If Option4 = True Then
  T_total2 = Format$(Val(t_total) / Val(t_cotizacion), "#####0.00")
  
Else
  T_total2 = Format$(Val(t_total) * Val(t_cotizacion), "#####0.00")
End If
End Sub

Private Sub t_sucursal_LostFocus()
  'Call inicia
End Sub

Private Sub t_total_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_total_LostFocus()
t_total = Format$(t_total, "######0.00")
End Sub
