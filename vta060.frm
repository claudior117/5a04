VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_porterceros 
   BackColor       =   &H00E0E0E0&
   Caption         =   "LIQUIDACION DE VENTA POR TERCEROS"
   ClientHeight    =   8490
   ClientLeft      =   -60
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.TextBox t_idprov 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   69
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00808080&
      Caption         =   "Deducciones ( a compras)"
      Height          =   1455
      Left            =   120
      TabIndex        =   65
      Top             =   3120
      Width           =   11415
      Begin VB.TextBox t_ngd 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox t_ivad 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox t_td 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Las deducciones  serán registradas en el modulo de Compras en la cuenta corriente del proveedor"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1080
         Width           =   8055
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00400040&
         Caption         =   "Neto Grav. Ded."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00400040&
         Caption         =   "Iva Dedudcciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3120
         TabIndex        =   67
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00400040&
         Caption         =   "Total Deducciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   66
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00808080&
      Caption         =   "Operaciones de Venta"
      Height          =   1095
      Left            =   120
      TabIndex        =   60
      Top             =   1920
      Width           =   11535
      Begin VB.TextBox t_ivav 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox t_in1 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox t_iva 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox c_tasa 
         Height          =   315
         ItemData        =   "vta060.frx":0000
         Left            =   4680
         List            =   "vta060.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox t_ib 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox t_pn 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox t_po 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Total 1:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8760
         TabIndex        =   64
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Iva:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   63
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Tasa Iva:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   62
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Importe Bruto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   54
      Top             =   7560
      Width           =   6015
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   4800
      TabIndex        =   49
      Top             =   6360
      Width           =   3255
      Begin VB.TextBox t_total 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox T_total2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Total U$s"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total Liquidacion"
      Height          =   855
      Left            =   120
      TabIndex        =   48
      Top             =   6360
      Width           =   4695
      Begin VB.TextBox t_ivat 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox t_bruto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox t_ng 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   71
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Neto a Cobrar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "No grabado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   58
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   9720
      TabIndex        =   40
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9000
      TabIndex        =   37
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
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
      Height          =   1335
      Left            =   120
      TabIndex        =   35
      Top             =   4680
      Width           =   8655
      Begin VB.ComboBox c_actividad 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   6720
         Picture         =   "vta060.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox t_observaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   16
         Top             =   960
         Width           =   6855
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Actividad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Vendedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos del Comprobante"
      Height          =   1935
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   9375
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "vta060.frx":0109
         Left            =   7440
         List            =   "vta060.frx":010B
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   8520
         Picture         =   "vta060.frx":010D
         Style           =   1  'Graphical
         TabIndex        =   53
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
         Picture         =   "vta060.frx":047F
         Style           =   1  'Graphical
         TabIndex        =   44
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
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "vta060.frx":0584
         Left            =   1680
         List            =   "vta060.frx":0586
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
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
         BackColor       =   &H00008000&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   57
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Vto.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   52
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
         TabIndex        =   47
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Nro. Liquidacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Cliente/Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   27
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta060.frx":0588
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta060.frx":0E0A
         Style           =   1  'Graphical
         TabIndex        =   28
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
      TabIndex        =   26
      Top             =   8235
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
            TextSave        =   "12/8/2018"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "19:46"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_porterceros"
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

Sub limpia()
   
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
   Option1 = True
   Call ABM_COMP_COMPRA2.armagrid
End Sub
Sub carga()
  Set rs = New ADODB.Recordset
  q = "select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = 401 and [num_comp] = " & Val(t_numcomp) & " and [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
 ' MsgBox (q)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     t_fecha = rs("fecha")
     
     c_vend.ListIndex = buscaindice(c_vend, rs("id_vendedor"))
     t_ib = Format$(rs("subtotal"), "######0.00")
     't_nograbado = Format$(rs("impuestos"), "######0.00")
     't_perc = Format$(rs("perc_iva") + rs("perc_gan") + rs("perc_ib"), "######0.00")
     t_ivav = Format$(rs("iva"), "######0.00")
     t_in1 = Format$(rs("total"), "######0.00")
     
         
   
    'cargo proveedores
    Set cl_comp = New COMPROBANTES
    Call cl_comp.cargar(401, t_letra, Val(t_sucursal), Val(t_numcomp), vta_clientes.t_idproveedor)
    If cl_comp.numint <> 0 Then
        t_ngd = Format$(rs("subtotal"), "######0.00")
        t_ivad = Format$(rs("iva"), "######0.00")
        t_td = Format$(rs("total"), "######0.00")
    End If
    
    Call CALCULATOTALES
  Else
     EXISTE = "N"
  End If
  
End Sub

Private Sub btnacepta_Click()
    Call iniciagraba

End Sub
Sub iniciagraba()
 J = MsgBox("Graba Comprobante ", 4)
 If J = 6 Then
  If verificaperiodog(t_fecha) = "A" Then
     para.z_actual = 0
     Call normal
  Else
   MsgBox ("Periodo Cerrado. Imposible grabar comprobante")
  End If
 End If
  
    

End Sub
Function verificatasaunica() As Boolean
   'devuelve true si la existe una sola tasa en la factura
 i = 1
 v = True
 While i <= msf1.Rows - 1
  If i = 1 Then
    tasa = Val(msf1.TextMatrix(i, 6))
  End If
  If tasa <> Val(msf1.TextMatrix(i, 6)) Then
    v = False
    i = msf1.Rows
  End If
  i = i + 1
 Wend
 verificatasaunica = v
End Function



Sub normal()
  Set rs = New ADODB.Recordset
  q = "select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = 401 and [num_comp] = " & Val(t_numcomp)
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
Unload Me
End Sub



Private Sub c_actividad_LostFocus()
If c_actividad.ListIndex < 0 Then
  c_actividad.ListIndex = 0
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
Call iniciaprov

End Sub
Sub iniciaprov()
Set cl_prov = New proveedores
cl_prov.carga (Val(vta_clientes.t_idproveedor))
If cl_prov.idprov <> 0 Then
   t_letra = cl_prov.letra
Else
   t_letra = "X"
End If
Set cl_prov = Nothing
End Sub

Sub inicia()
espere.Show
espere.Label1 = "Inicializando Comprobante....."
espere.Refresh
   gcuit = vta_clientes.t_cuit
   c_vend.ListIndex = buscaindice(c_vend, vta_clientes.t_idvend)
   Set cl_compvta = New comprobantes_venta
   cl_compvta.sucursal = Val(c_sucursal)
   cl_compvta.actual (400)
   cantlineas = cl_compvta.cant_lineas
   Set cl_compvta = Nothing
   t_cotizacion = para.cotizacion

     t_alicuotaib = "0.00"
     T_PERCIB = "0.00"
     'gcuit = "0"
   
   
   Unload espere







End Sub

Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)
End If
t_sucursal = Format$(c_sucursal, "0000")
t_numcomp = ""
End Sub

Private Sub c_tasa_LostFocus()
If c_tasa.ListIndex < 0 Then
  c_tasa.ListIndex = 0
End If
t_ivav = Format$(Val(t_ib) * Val(c_tasa) / 100, "#####0.00")

End Sub

Private Sub c_tipocomp_GotFocus()
btnacepta.Enabled = False
End Sub

Private Sub c_tipocomp_LostFocus()
Call iniciacli
End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
End If
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
vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
vta_clientes.carga
vta_clientes.Show
End Sub









Sub iniciacli()
   vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
   vta_clientes.carga
  
   
End Sub

Sub CALCULATOTALES()
t_bruto = Format$(Val(t_ib) - Val(t_ngd), "#######0.00")
t_ivat = Format$(Val(t_ivav) - Val(t_ivad), "#######0.00")
t_ng = Format$(Val(t_ng), "#######0.00")
t_total = Format$(Val(t_bruto) + Val(t_ivat) + Val(t_ng), "#######0.00")
T_total2 = Format$(Val(t_total) / Val(t_cotizacion), "#######0.00")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
  
End Select

End Sub
Sub grabacontcompras()

     
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 21)
End If


End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_clientesprov(c_prov)
c_prov.ListIndex = 0

Call carga_SUCURSALES(c_sucursal)
c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)
Set rs = New ADODB.Recordset




c_tipocomp.clear
c_tipocomp.AddItem "Liquidacion", 0
c_tipocomp.ListIndex = 0

Set rs = New ADODB.Recordset
q = "select * from vta_05 order by [denominacion]"
rs.Open q, cn1
Call llena_combo(rs, "denominacion", "id_vendedor", c_vend, True)
Set rs = Nothing
c_vend.ListIndex = 0



Call barraesag(Me)

Option1 = True
If para.moneda = "P" Then
  Option4 = True
Else
  Option3 = True
End If
t_sucursal = Format$(glo.sucursal, "0000")
Load vta_liqcereal1
Frame11.Visible = False

Call carga_actividades(c_actividad)

Load vta_clientes
vta_clientes.limpia
gcuit = "0"

Call carga_tasaiva(c_tasa)
c_tasa.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_facturacion1
Unload vta_facturacion2
Unload vta_selremitos
Unload vta_clientes
Unload ABM_COMP_COMPRA2
End Sub


Sub graba()
  'On Error GoTo ERRORGRABA
  
  numint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = 1
  cl_compvta.actual (401)
  cl_compvta.letra = t_letra
  cl_compvta.numcomp = Val(t_numcomp)
  abreviatura = cl_compvta.abreviatura
  ubicacionctacte = cl_compvta.ctacte
         
  ep = "N"
  cp = "0000-00000000"
  contado = "N"
  If Option4 = True Then
       ssi = Val(t_total)
  Else
       ssi = Val(T_total2)
  End If
      
      
  If EXISTE = "N" Then
       cl_compvta.ACTUALIZA_NUMERADOR
  End If
      
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
       cuentaact = rs("cuenta_contable_venta")
      Else
       codact = 0
       alicuotaib = 0
       cuentaact = para.cuenta_ventas
      End If
      Set rs = Nothing
      
        
              
      tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
       
      idcli = c_prov.ItemData(c_prov.ListIndex)
      
      cn1.BeginTrans
       
       
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], " & _
" [dominio_acoplado02], [SALDO_IMPAGO02], [num_z], [cae], [cae_vence], [tipo_op])"



QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', 401" & _
", " & idcli & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_ib) & ", 0, " & Val(t_ivav) & ", " & Val(t_in1) & _
", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & t_observaciones & _
" ', " & Val(t_cotizacion) & ", " & Val(T_total2) & ", '" & moneda & "', " & c_vend.ItemData(c_vend.ListIndex) & ", '" & cl_compvta.venta & "', '" & contado & "', 0" & _
", 0, 0, " & codact & ", 0, 0, 0, '" & t_fechavto & "', 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & _
", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & _
Left$(vta_clientes.t_localidad, 50) & "', " & tiporespiva & ", ' ', ' ', ' ', " & ssi & ", " & para.z_actual & ", 'u2', '01/01/2018', 1)"

        'MsgBox (QUERY)
      cn1.Execute QUERY
      
      Set cl_cli = Nothing
        
       
              
      
      'actualizo tasa de iva
      If cl_compvta.grabado <> "N" Then
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
          QUERY = QUERY & " VALUES (" & numint & ", " & Val(c_tasa) & ", " & Val(t_ivav) & ", " & Val(t_ib) & ", " & tiporespiva & ", " & cuentaact & ")"
          cn1.Execute QUERY
      End If
     
      
      
           
   'contabilida
   If Generaasientosauto Then
     If cl_compvta.contabilidad <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")

         u1 = cl_compvta.contabilidad
          
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         
         'grabo asiento
         'If Option3 = True Then
           'fact en dolares
         '  tot = Val(T_total2)
         '  m = Val(t_cotizacion)
         'Else
           tot = Val(t_in1)
           m = 1
         'End If
         
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Ventas] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'V', " & numint & ", " & tot & ", " & tot & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_prov), 50) & "')"
         cn1.Execute QUERY
      
         
         
           'ingresa deudores
           cta = para.cuenta_deudores
           ic = 1
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & tot & ", '" & dcta & "')"
           cn1.Execute QUERY
           ic = ic + 1
          

         
                   
         
         If Val(t_ivav) > 0 And cl_compvta.grabado <> "N" Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_ventas & ", '" & u2 & "', " & Format(Val(t_ivav) * m, "#####0.00") & ", 'IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         'contrapartida
         
         If cl_compvta.grabado <> "N" Then
           importe = Val(t_ib) * m
         Else
           importe = Val(t_in1) * m
         End If
         
         cuentaact = para.cuenta_ventas
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cuentaact & ", '" & u2 & "', " & Format(importe, "######0.00") & ", '" & "Ventas" & "')"
         cn1.Execute QUERY
         ic = ic + 1
      
      
      End If
      
    End If
      
      
     
      Set rs = Nothing
      Set cl_compvta = Nothing
      Set cl_cli = Nothing

         
      
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Emitir liq. cereal NI:" & numint & "', " & para.id_usuario & ", 'V', " & numint & ", '" & Now & "', '[400] " & t_letra & " " & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 12, " & idcli & ")"
  
     cn1.Execute QUERY

     cn1.CommitTrans
          
     If Val(t_td) > 0 Then
       Call grabacompcompras
     
     End If
      
      
      
      
      
      
       
      
      
      Call INICIALIZA2(Me)
      
      c_tipocomp.SetFocus
      t_sucursal = Format$(c_sucursal, "0000")
      
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos y vuelva a repetir la operacion")
  

End Sub

Sub grabacompcompras()
      
     ' On Error GoTo errc
      numintc = saca_ultnumero_int_comp("C")
      
      Set rs = New ADODB.Recordset
      q = "select [id_proveedor] from vta_01 where [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        idprov = rs("id_proveedor")
      Else
        idprov = 1
      End If
      Set rs = Nothing
      
      Set cl_comp = New COMPROBANTES
      cl_comp.actual (401)
      
      cc = cl_comp.ctacte
      ctdo = "N"
     
      tc_obs = " "
      
      If Option4 = True Then
        moneda = "P"
      Else
        moneda = "D"
      End If
      tom = Val(t_td)
      cp = "0000-00000000"
      cn1.BeginTrans
      
      QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], " & _
" [no_grabado], [percep_ret], [iva], [total], [fecha_prob_entrega], [fecha_recepcion], [estado], [id_codretgan], [id_cuenta], [stock], [ctacte], [grabado], " & _
" [estado_pago], [num_op], [id_codretib], [obs], [condiciones], [info_contacto], [moneda], [cotiz_dolar], [contado], [TOTAL_D], [monto_suj_ret], " & _
"[alicuota_ret], [ret_mes], [pagos_realizados], [pago_actual], [minimo_no_imp], [fecha_vto], [COMPRA], [saldo_impago])"
      
 QUERY = QUERY & " VALUES (" & numintc & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', 401" & _
 ", " & idprov & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_ngd) & ", 0, 0, " & Val(t_ivad) & _
 ", " & Val(t_td) & ", '" & Format$(Now, "dd/mm/yyyy") & "', '" & t_fecha & "', 'A', 0, 110101, '" & _
 cl_comp.STOCK & "', '" & cc & "', '" & cl_comp.grabado & "', 'P', '" & cp & "', 0, 'Por VenTas', ' ', ' ', '" & moneda & "', " & _
 Val(t_cotizacion) & ", '" & ctdo & "', " & tom & ", 0, 0, 0, 0, 0, 0, '" & t_fechavto & "', '" & cl_comp.compra & "', 0)"
 
 'MsgBox (QUERY)
 cn1.Execute QUERY
   
  
  
 If Generaasientosauto Then
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
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Compras] " & cl_comp.abreviatura & " " & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'C', " & numint & ", " & Val(t_td) & ", " & Val(t_td) & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_prov), 50) & "')"
         cn1.Execute QUERY
      
         ic = 1
        
         
         'cuenta madre ctacte o caja
           'ingresa deudores
           cta = para.cuenta_acreedores
           Set rs = New ADODB.Recordset
           q = "select * from c_01 where [id_cuenta] = " & cta
           rs.Open q, cn1
           If Not rs.EOF And Not rs.BOF Then
             dcta = rs("descripcion")
           Else
             dcta = "Cuenta Inexistente"
           End If
           Set rs = Nothing
           
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & Format(Val(t_td), "######0.00") & ", '" & dcta & "')"
           cn1.Execute QUERY
           ic = ic + 1
         
          
         If Val(t_ivad) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_compras & ", '" & u2 & "', " & Val(t_ivad) & ", 'IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         'contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_acreedores & ", '" & u2 & "', " & Val(t_ngd) & ", '" & c_cuenta & "')"
         cn1.Execute QUERY
      
       End If
          
      End If
  
  
  cn1.CommitTrans
  
  
Exit Sub

errc:
  MsgBox ("Error al actualizar Deducciones en el modulo Compras. Verifque")
  Exit Sub

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



Private Sub t_bruto_GotFocus()
t_bruto = Format$(Val(t_in1) - Val(t_td) - Val(t_perc), "#####0.00")
End Sub

Private Sub t_bruto_LostFocus()
Call sacatotales
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


Private Sub t_ib_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_ib_LostFocus()
t_ib = Format$(Val(t_ib), "######0.00")
End Sub

Private Sub t_iva_GotFocus()
t_iva = Format$(Val(t_ib) * Val(c_tasa) / 100, "######0.00")
End Sub

Private Sub t_ivad_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_ivad_LostFocus()
t_td = Format$(Val(t_ngd) + Val(t_ivad), "######0.00")
End Sub

Private Sub t_ivav_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_ivav_LostFocus()
t_in1 = Format$(Val(t_ib) + Val(t_ivav), "######0.00")
End Sub

Private Sub t_ngd_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub

Private Sub t_numcomp_LostFocus()
If IsNumeric(t_numcomp) Then
   t_numcomp = Format$(t_numcomp, "00000000")
   'If glo.sucursalf <> Val(c_sucursal) Then
     Call carga
  ' Else
   '  EXISTE = "N"
   'End If
   
   c_actividad.ListIndex = buscaindice(c_actividad, sacaactividadsucursal(Val(t_sucursal)))
   

Else
  t_numcomp.SetFocus
End If
End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub



Sub sacatotales()
t_bruto = Format$(Val(t_bruto), "######0.00")
t_nograbado = Format$(Val(t_nograbado), "######0.00")
t_perc = Format$(Val(t_perc), "######0.00")
t_iva = Format$(Val(t_iva), "######0.00")
t_total = Format$(Val(t_bruto), "######0.00")
If Option4 = True Then
 If Val(t_cotizacion) < 1 Then
   t_cotizacion = 1
 End If
 T_total2 = Format$(Val(t_total) / Val(t_cotizacion), "#####0.00")
Else
  T_total2 = Format$(Val(t_total) * Val(t_cotizacion), "#####0.00")
End If
t_ngd = Format$(Val(t_ngd), "#####0.00")
t_ivad = Format$(Val(t_ivad), "#####0.00")
t_td = Format$(Val(t_td), "#####0.00")



End Sub
Sub sacaperc()
If Option3 = True Then
   s = Val(t_subtotal) * Val(t_cotizacion)
 Else
   s = Val(t_subtotal)
 End If

q = "select * from i_01 where [id_impuesto] = 1"
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
If Not rs2.EOF And Not rs2.BOF Then
 impmin = rs2("importe_minimo_sujeto_ret")
 retmin = rs2("retencion-minima")
  If rs2("calcula") = "S" And s >= impmin Then
   tp = s * (Val(t_alicuotaib) / 100)
   If tp >= retmin Then
       t_perc = Format$(tp, "#####0.00")
   Else
       t_perc = "0.00"
   End If
 Else
  t_perc = "0.00"
 End If
Else
 t_perc = "0.00"
End If

If Option3 = True Then 'dolares
   p$ = Val(t_perc) / Val(t_cotizacion)
Else
   p$ = Val(t_perc)
End If
t_perc = Format$(p$, "####0.00")
Set rs2 = Nothing

If Check1 = 1 Then
  'calcula perciva rg 2459
   Set rs2 = New ADODB.Recordset
   q = "select * from  i_01, i_02 where i_01.[id_impuesto] = i_02.[id_impuesto] and i_01.[id_impuesto] = 2"
   rs2.Open q, cn1
   If Not rs2.EOF And Not rs2.BOF Then
     Set cl_cli = New Clientes
     cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
     If cl_cli.id > 0 Then
        If cl_cli.operador_granos = "S" Then
           t_alicuotaperciva = rs2("alicuota_i")
        Else
           t_alicuotaperciva = rs2("alicuota_n")
        End If
     Else
        t_alicuotaperciva = "0.00"
     End If
    ' Set cl_cli = Nothing
   Else
     t_alicuotaperciva = "0.00"
   End If
   
   
   If s >= rs2("importe_minimo_sujeto_ret") Then
     pi = s * Val(t_alicuotaperciva) / 100
   Else
     pi = 0
   End If
   
   If Option3 = True Then 'dolares
    p$ = pi / Val(t_cotizacion)
   Else
    p$ = pi
   End If
           
   If p$ >= rs2("retencion-minima") Then
      t_perciva = Format$(pi, "######0.00")
   Else
      t_perciva = "0.00"
   End If
   Set rs2 = Nothing
Else
  t_alicuotaperciva = "0.00"
  t_perciva = "0.00"
End If




End Sub

Private Sub t_po_GotFocus()
t_po = Format$(Val(t_pr) * Val(t_factor) / 100, "####0.00")
End Sub

Private Sub t_sucursal_GotFocus()
t_sucursal = Format$(Val(c_sucursal), "0000")
End Sub

Private Sub t_sucursal_LostFocus()
Call inicia
End Sub

Private Sub t_td_GotFocus()
Call CALCULATOTALES
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



