VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0A6BE9FC-5039-11D5-98EC-0800460222F0}#1.0#0"; "IFEpson.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_facturacion 
   BackColor       =   &H00E0E0E0&
   Caption         =   "FACTURACION"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11880
   Begin VB.Frame Frame14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Plan de pago en cuotas"
      Height          =   855
      Left            =   240
      TabIndex        =   83
      Top             =   7440
      Width           =   6015
      Begin VB.TextBox t_fechacuota1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   80
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_cuotas_sininteres 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   88
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox t_interes_cuota 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   87
         Top             =   480
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   4440
         TabIndex        =   86
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox t_valorcuota 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   82
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_cantcuotas 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   81
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Fecha cuota1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   91
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Sin Int."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Interes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   89
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Valor Cuota"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   85
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Cant.Cuotas"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   84
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox t_cae_vence 
      Height          =   285
      Left            =   12240
      TabIndex        =   75
      Text            =   "20180101"
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox t_cae 
      Height          =   285
      Left            =   12240
      TabIndex        =   74
      Text            =   "u2"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox t_cantlineas 
      Height          =   285
      Left            =   10440
      TabIndex        =   73
      Text            =   "Text1"
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox t_cl 
      Height          =   285
      Left            =   10440
      TabIndex        =   72
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   9720
      TabIndex        =   70
      Top             =   1440
      Width           =   1935
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Moneda Unica"
         Height          =   315
         Left            =   120
         TabIndex        =   71
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   9000
      TabIndex        =   68
      Top             =   5880
      Width           =   2775
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprime Descripcion Extra"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal epson1 
      Left            =   0
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   240
      TabIndex        =   61
      Top             =   7560
      Width           =   6015
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   9000
      TabIndex        =   56
      Top             =   5280
      Width           =   2775
      Begin VB.CommandButton Command4 
         Caption         =   "Cal"
         Height          =   255
         Left            =   2280
         TabIndex        =   58
         ToolTipText     =   "Para calcular la Percepcion por canje de cereal RG2459 es necesario que el cliente este marcado como OPERADOR DE GRANOS"
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por canje Cereal RG 2459"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales"
      Height          =   855
      Left            =   6840
      TabIndex        =   53
      Top             =   7440
      Width           =   2535
      Begin VB.TextBox t_total 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox T_total2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   21
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
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Parciales"
      Height          =   855
      Left            =   240
      TabIndex        =   50
      Top             =   6600
      Width           =   6375
      Begin VB.TextBox t_subtotal 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_porcdescuento 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   15
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox t_descuento 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "IVA"
         Height          =   195
         Left            =   5160
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox t_iva 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_subtotal2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox t_nograbado 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Subtotal2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         Caption         =   "%"
         Height          =   255
         Left            =   2760
         TabIndex        =   78
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Descuento"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   77
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Subtotal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "No Grabado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Percepciones"
      Height          =   855
      Left            =   6840
      TabIndex        =   49
      Top             =   6600
      Width           =   2295
      Begin VB.CommandButton Command6 
         Caption         =   "Percepciones"
         Height          =   195
         Left            =   360
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_perc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   360
         MaxLength       =   10
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   " Remitos"
      Height          =   255
      Left            =   8160
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9720
      TabIndex        =   40
      Top             =   720
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
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9720
      TabIndex        =   37
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command8 
         Caption         =   "F.P."
         Height          =   255
         Left            =   1080
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contado "
         Height          =   255
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9960
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   35
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
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales del Comprobante"
      Height          =   1335
      Left            =   240
      TabIndex        =   32
      Top             =   5280
      Width           =   8655
      Begin VB.ComboBox c_tipoop 
         Height          =   315
         ItemData        =   "vta003A.frx":0000
         Left            =   6720
         List            =   "vta003A.frx":000D
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox c_actividad 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   600
         Width           =   5055
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   6720
         Picture         =   "vta003A.frx":003E
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
         TabIndex        =   12
         Top             =   960
         Width           =   6855
      End
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Actividad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
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
         BackColor       =   &H00800080&
         Caption         =   "Vendedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
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
      TabIndex        =   27
      Top             =   0
      Width           =   9375
      Begin VB.CheckBox ch_plan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Habilita Plan Cuotas"
         Height          =   375
         Left            =   4920
         TabIndex        =   92
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox c_transferencia 
         Height          =   315
         ItemData        =   "vta003A.frx":0143
         Left            =   6000
         List            =   "vta003A.frx":014D
         TabIndex        =   4
         Top             =   1200
         Width           =   3135
      End
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "vta003A.frx":0189
         Left            =   7440
         List            =   "vta003A.frx":018B
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   8520
         Picture         =   "vta003A.frx":018D
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox t_fechavto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_cotizacion 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   7680
         Picture         =   "vta003A.frx":04FF
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
         TabIndex        =   22
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "vta003A.frx":0604
         Left            =   1680
         List            =   "vta003A.frx":0606
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
         TabIndex        =   5
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Metodo Transf"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4800
         TabIndex        =   76
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   66
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Vto.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   59
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
         TabIndex        =   48
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   24
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta003A.frx":0608
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta003A.frx":0E8A
         Style           =   1  'Graphical
         TabIndex        =   25
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
      TabIndex        =   23
      Top             =   8385
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
            TextSave        =   "14/11/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:36 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_facturacion"
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

Dim Fiscaltf As Driver

Sub electronica()

    Dim seguir As Boolean
    
    seguir = True
    
    'On Error GoTo ManejoError
    
    If Not fe_valida_tique() Then
        'el tique esta vencido y tenemos que generarlo de nuevo
        If Not fe_genera_wsaa() Then
          MsgBox ("Error al generar tique WSAA, verificar conexion y regisar log")
          seguir = False
        End If
    End If
    
    Debug.Print ("Tiquet correcto")
    
    
    If (c_tipocomp.ItemData(c_tipocomp.ListIndex) >= 2 And c_tipocomp.ItemData(c_tipocomp.ListIndex) <= 3) Or (c_tipocomp.ItemData(c_tipocomp.ListIndex) >= 31 And c_tipocomp.ItemData(c_tipocomp.ListIndex) <= 32) Then
       If Val(vta_selcomp.t_seleccionados) = 0 Then
           MsgBox ("Para realizar NC/ND es necesario que seleccione la factura asociada")
           seguir = False
       End If
    End If
    
       
    
    
    
    If seguir Then
      'se va a emitir el comprobante
       Debug.Print ("Ingreso a la generacion del comprobante")
       ' Crear objeto interface Web Service de Factura Electrónica de Mercado Interno
        Set WSFEv1 = CreateObject("WSFEv1")
    
        ' Setear tocken y sing de autorización (pasos previos)
        WSFEv1.Token = para.facte_token
        WSFEv1.Sign = para.facte_sign
    
        ' CUIT del emisor (debe estar registrado en la AFIP)
        cuitemisor = Mid$(glo.CUIT, 1, 2) & Mid$(glo.CUIT, 4, 8) & Mid$(glo.CUIT, 13, 1)
        WSFEv1.CUIT = cuitemisor
        
        WSFEv1.LanzarExcepciones = False
    
        ' Conectar al Servicio Web de Facturación
        proxy = "" ' "usuario:clave@localhost:8000"
        wsdl = para.facte_servidor_wsfe
        cache = "" 'Path
        wrapper = "" ' libreria http (httplib2, urllib2, pycurl)
        cacert = ""
    
        ok = WSFEv1.Conectar(cache, wsdl, proxy, wrapper, cacert)
        ControlarExcepcion_fe WSFEv1
    
        WSFEv1.Dummy
        ControlarExcepcion_fe WSFEv1
    
        If (WSFEv1.AppServerStatus = "OK" And WSFEv1.DbServerStatus = "OK" And WSFEv1.AuthServerStatus = "OK") Then
       
             Debug.Print ("Ingreso al servidor")
             Set cl_compvta = New comprobantes_venta
             cl_compvta.sucursal = glo.sucursale
             cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
            ' Establezco los valores de la factura a autorizar:
             Select Case t_letra
             
             Case Is = "A"
                tipo_cbte = cl_compvta.cod_afip_a
             Case Is = "B"
                tipo_cbte = cl_compvta.cod_afip_b
             Case Is = "C"
                tipo_cbte = cl_compvta.cod_afip_c
           
             End Select
             punto_vta = glo.sucursale
             cbte_nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta)
             ControlarExcepcion_fe WSFEv1
            
             If cbte_nro = "" Then
                cbte_nro = 0                ' no hay comprobantes emitidos
             Else
                cbte_nro = CLng(cbte_nro)   ' convertir a entero largo
             End If
             cbte_nro = cbte_nro + 1
             
             ' hacer esto solo si los numeros comprobantes coinciden
             If cbte_nro = Val(t_numcomp) Then
                     fecha = Format(t_fecha, "yyyymmdd")
                     concepto = c_tipoop.ListIndex + 1
                     If vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex) <> 3 Then
                       tipo_doc = 80
                       CUIT2 = RTrim$(vta_clientes.t_cuit)
                     Else
                        Set rs3 = New ADODB.Recordset
                        q = "select minimo_informar_cons_final from g0 where sucursal= 0"
                        rs3.Open q, cn1
                        If Not rs3.EOF And Not rs3.BOF Then
                           If Val(t_total) < rs3("minimo_informar_cons_final") Then 'minimo de consumidor final sin informar datos
                               tipo_doc = 99
                               CUIT2 = "0"
                           Else
                              tipo_doc = 96
                              CUIT2 = RTrim$(vta_clientes.t_cuit)
                           End If
                        Else
                            tipo_doc = 99
                               CUIT2 = "0"
                        End If
                        Set rs3 = Nothing
                     End If
                     
                     cbt_desde = cbte_nro
                     cbt_hasta = cbte_nro
                     
                     If t_letra <> "C" Then
                        imp_total = t_total
                        imp_tot_conc = t_nograbado
                        imp_neto = t_subtotal
                        imp_iva = t_iva
                        imp_trib = "0.00"
                        imp_op_ex = "0.00"
                     Else
                        imp_total = t_total
                        imp_tot_conc = "0.00"
                        imp_neto = t_subtotal
                        imp_iva = "0.00"
                        imp_trib = "0.00"
                        imp_op_ex = "0.00"
                     
                     
                     End If
                     fecha_cbte = fecha
                     If Option1 = True Then
                          'cc 10 dias a partir factura
                          fechapago = DateValue(t_fecha) + 10
                    Else
                          fechapago = ""
                    End If
                     
                     
                     
                   
                     
                     
                     If c_tipoop.ListIndex >= 1 Then
                       
                       fecha_serv_desde = fecha
                       fecha_serv_hasta = fecha
                     Else
                       fecha_serv_desde = ""
                       fecha_serv_hasta = ""
                     End If
                    
                    If Option4 = True Then
                        moneda_id = "PES"
                        moneda_ctz = "1.000"
                    Else
                        moneda_id = "DOL"
                        moneda_ctz = Format$(Val(t_cotizacion), "####0.000")
                    End If
                   
                    
                    If c_tipocomp.ItemData(c_tipocomp.ListIndex) >= 30 And c_tipocomp.ItemData(c_tipocomp.ListIndex) <= 32 Then
                      fecha_venc_pago = Format(t_fechavto, "yyyymmdd")
                      If c_transferencia.ListIndex = 0 Then
                        mtransf = "SCA"
                      Else
                       mtransf = "ADC"
                      End If
                    Else
                        fecha_venc_pago = ""
                    End If
                    
                    
                    
        
                    ok = WSFEv1.CrearFactura(concepto, tipo_doc, CUIT2, tipo_cbte, punto_vta, _
                        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
                        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
                        fecha_serv_desde, fecha_serv_hasta, _
                        moneda_id, moneda_ctz)
            
                  
                  
                    'agrego opcionales factura credito
                     If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 30 Then
                      
                      
                      
                      Set rs31 = New ADODB.Recordset
                      q = "select * from g0 where [sucursal] = 0"
                      rs31.Open q, cn1
                      
                      'parametros
                      ok = WSFEv1.AgregarOpcional(2101, rs31("cbu")) 'cbu
                      ok = WSFEv1.AgregarOpcional(2102, rs31("alias")) 'alias
                      ok = WSFEv1.AgregarOpcional(27, mtransf) 'transmisión (desde el 01/04/2021)
                      
                      Set rs31 = Nothing
                     Else
                      fecha_venc_pago = ""
                      
                     End If
                  
                     If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 32 Then
                            ok = WSFEv1.AgregarOpcional(22, "S") 'anula
                      End If
                  
                  ' Agrego los comprobantes asociados:
                  
                  tc21 = c_tipocomp.ItemData(c_tipocomp.ListIndex)
                  
                  If tc21 = 2 Or tc21 = 3 Or tc21 = 31 Or tc21 = 32 Then ' solo nc/nd
                     F = vta_selcomp.msf1.Rows - 1
                     compasocnc = ""
                     For i = 1 To F
                        If vta_selcomp.msf1.TextMatrix(i, 0) = "**" Then
                             tipo_cbte_asoc_1 = CInt(vta_selcomp.msf1.TextMatrix(i, 4))
                             punto_vta_asoc_1 = CInt(vta_selcomp.msf1.TextMatrix(i, 1))
                             cbte_nro_asoc_1 = CLng(vta_selcomp.msf1.TextMatrix(i, 2))
                             cuit_cbte_asoc_1 = cuitemisor 'RTrim$(vta_clientes.t_cuit) 'cuitemisor  ´
                             fecha_cbte_asoc_1 = vta_selcomp.msf1.TextMatrix(i, 3)
                             'MsgBox (tipo_cbte_asoc_1 & " " & punto_vta_asoc_1 & " " & cbte_nro_asoc_1 & " " & cuit_cbte_asoc_1 & " " & fecha_cbte_asoc_1)
                             compasocnc = vta_selcomp.msf1.TextMatrix(i, 1) & "-" & vta_selcomp.msf1.TextMatrix(i, 2)
                             i = F
                         End If
                    Next i
                    
                    
                    
                    ok = WSFEv1.AgregarCmpAsoc(tipo_cbte_asoc_1, punto_vta_asoc_1, cbte_nro_asoc_1, cuit_cbte_asoc_1, fecha_cbte_asoc_1)
                    'ok = WSFEv1.AgregarCmpAsoc(tipo_cbte_asoc_1, punto_vta_asoc_1, cbte_nro_asoc_1, cuitemisor)
                    'ok = WSFEv1.AgregarCmpAsoc(tipo_cbte_asoc_1, punto_vta_asoc_1, cbte_nro_asoc_1)
                    
                    
                  End If
                
                
                
               
            If t_letra <> "C" Then
                ' Agrego percepcion ib
                If Val(t_perc) > 0 Then
                    id = 99
                    Desc = "Percepcion Ing Brutos"
                    base_imp = t_subtotal
                    alic = Format$(Val(t_alicuotaib) / 100, "#0.00")
                    importe = t_perc
                    ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)
                End If
                
                
                     ' Agrego percepcion iva
                If Val(t_perciva) > 0 Then
                    id = 99
                    Desc = "Percepcion Iva"
                    base_imp = t_subtotal
                    alic = Format$(Val(t_alicuotaperciva) / 100, "#0.00")
                    importe = t_perciva
                    ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)
                End If
                
                                
                ' Agrego tasas de IVA
                 
                    For i = 1 To 7
                      If Val(vta_facturacion2.msf1.TextMatrix(i, 1)) > 0 Then
                          Select Case Val(vta_facturacion2.msf1.TextMatrix(i, 0))
                            Case Is = 0
                                 id = 3
                            Case Is = 10.5
                                 id = 4
                            Case Is = 21
                                 id = 5
                            Case Is = 27
                                 id = 6
                            Case Is = 5
                                 id = 8
                            Case Is = 2.5
                                 id = 9
                          End Select
                          base_imp = Format$(Val(vta_facturacion2.msf1.TextMatrix(i, 1)), "0.00")
                          importe = Format$(Val(vta_facturacion2.msf1.TextMatrix(i, 2)), "0.00")
                          ok = WSFEv1.AgregarIva(id, base_imp, importe)
                       End If
                     Next i
             End If
            
            ' Habilito reprocesamiento automático (predeterminado):
            WSFEv1.Reprocesar = True
        
            ' Solicito CAE:
            cae = WSFEv1.CAESolicitar()
            ControlarExcepcion WSFEv1
            
            If WSFEv1.cae <> "" Then
               t_cae = WSFEv1.cae
               t_cae_vence = WSFEv1.Vencimiento
               t_numcomp = WSFEv1.CbteNro
               Debug.Print t_cae
               Debug.Print t_cae_vence
               Debug.Print t_numcomp
               
               MsgBox ("Comprobante Electronico Cae: " & t_cae & " generado correctamente, ahora se procederá a grabar localmente e imprimir")
               
               
               'guardo xml
               fd = FreeFile
                Open "c:\5a04\log\xml.txt" For Append As fd
                If Not WSFEv1 Is Nothing Then
                        Print #fd, WSFEv1.XmlRequest
                        Print #fd, WSFEv1.XmlResponse
                End If
                Close fd
                  
               
               Call graba
               
            Else
                
               MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
            
              
              ' Muestro los errores
              If WSFEv1.errmsg <> "" Then
                MsgBox WSFEv1.errmsg, vbExclamation, "Error"
              End If
            
              ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
               For Each evento In WSFEv1.eventos:
                 MsgBox evento, vbInformation, "Evento"
               Next
            End If
     
     
     Else 'else si los numeros de comprobantes no coinciden
       
         
                  
            MsgBox ("Los numeros de comprobantes entre el sistema y el AFIP no coinciden. Último comprobante del AFIP: " & cbte_nro - 1 & " Ultimo comprobante del sistema: " & Val(t_numcomp) - 1 & " Verifique antes de continuar!!")
         
        
     End If
 Else
   MsgBox ("Error al conectar al servidor WSFE")
 End If
 End If
    
    
Exit Sub

ManejoError:
    fd = FreeFile
    Open "c:\5a04\log\error.txt" For Append As fd
    If Not WSFEv1 Is Nothing Then
            Print #fd, WSFEv1.Excepcion
            Print #fd, WSFEv1.Traceback
            Print #fd, WSFEv1.XmlRequest
            Print #fd, WSFEv1.XmlResponse
            Print #fd, WSFEv1.DebugLog()
            ' guardo mensaje de error para mostrarlo:
            Excepcion = WSFEv1.Excepcion
    End If
    Close fd
        
  
End Sub

Sub ControlarExcepcion_fe(obj As Object)
    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0
    If obj.Excepcion <> "" Then
        ' Depuración (grabar a un archivo los detalles del error)
        fd = FreeFile
        Open "c:\excepcion.txt" For Append As fd
        Print #fd, obj.Excepcion
        Print #fd, obj.Traceback
        Print #fd, obj.XmlRequest
        Print #fd, obj.XmlResponse
        Close fd
        MsgBox obj.Excepcion, vbExclamation, "Excepción"
        End
    End If
End Sub

Sub iniciacomp()
Set rs = New ADODB.Recordset
q = "select [imprime_desc_extra], [cant_lineas] from vta_06 where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex)
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  If rs("imprime_desc_extra") = "S" Then
    Check2 = 1
  Else
    Check2 = 0
  End If
  t_cantlineas = rs("cant_lineas")
Else
  Check2 = 0
  t_cantlineas = 25
End If
Set rs = Nothing
t_cae = "u2"
t_cae_vence = "20180101"
'Call mensaje
vta_selremitos.limpia

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

If para.muestrasaldofactventa = "N" Then
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
    If Option4 = True Then
     tm = tm & "  " & " en $"
    Else
     tm = tm & "  " & " en U$s"
    End If

    
Else
    Set cl_cli = New Clientes
    cl_cli.carga (Val(vta_clientes.t_id))
    If cl_cli.id > 1 Then
     tm = "SALDO A LA FECHA $" & Format$(cl_cli.saldo(True, t_fecha, True), "######0.00")
    Else
     tm = ""
    End If
    Set cl_cli = Nothing
    Label20.FontBold = True
End If
Label20 = UCase$(tm)
Frame11.Visible = True


End Sub
Sub carga()
  Set rs = New ADODB.Recordset
  q = "select [fecha], [fecha_vto], [cotizacion_dolar], [id_cliente], [num_int], [id_vendedor], [subtotal], [impuestos], [total], [perc_ib], [perc_gan], [perc_iva], [iva], " & _
  " [contado], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [observaciones] from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
  rs.MaxRecords = 1
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     t_fecha = rs("fecha")
     t_fechavto = rs("fecha_vto")
     t_cotizacion = rs("cotizacion_dolar")
     
     c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select [id_producto], [descripcion], [cantidad], [unidad], [pu], [tasaiva], [importe], [pu_final], [tasaib], [num_int], [renglon] from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     Call armagrid
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("unidad") & Chr$(9) & Format$(rs1("pu"), "######0.00") & Chr(9) & rs1("tasaiva") & Chr(9) & rs1("importe") & Chr(9) & rs1("pu_final") & Chr(9) & rs1("tasaib")
        
        Set rs2 = New ADODB.Recordset
        q = "select [desc_ext], [cant_lineas] from vta_015 where [num_int] = " & rs1("num_int") & " and [renglon] = " & rs1("renglon")
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
     t_subtotal = Format$(rs("subtotal"), "######0.00")
     t_nograbado = Format$(rs("impuestos"), "######0.00")
     t_perc = Format$(rs("perc_iva") + rs("perc_gan") + rs("perc_ib"), "######0.00")
     t_iva = Format$(rs("iva"), "######0.00")
     t_total = Format$(rs("total"), "######0.00")
     If Not IsNull(rs("observaciones")) Then
            t_observaciones = rs("observaciones")
     Else
       t_observaciones = "*"
     End If
     vta_formapago.armagrid2
     If rs("contado") = "S" Then
        vta_clientes.t_cli = rs("cliente02")
        vta_clientes.t_direccion = rs("direccion02")
        vta_clientes.t_cuit = rs("cuit02")
        vta_clientes.t_localidad = rs("localidad02")
        vta_clientes.c_iva.ListIndex = buscaindice(vta_clientes.c_iva, rs("id_tipo_iva02"))
        Option2 = True
        
        
        
        
        
    End If
     
     
  
  
   
  
  Else
     EXISTE = "N"
  End If
  Set rs = Nothing
  
End Sub

Sub carga2()
  Set rs = New ADODB.Recordset
  If para.numeracion_comun_Fact_nc = "S" Then
      q = "select [num_int] from vta_02 where  [id_tipocomp] = 1 and [sucursal]= " & Val(t_sucursal) & " and [letra]= '" & t_letra & "' and [num_comp]= " & Val(t_numcomp)
  Else
      q = "select [num_int] from vta_02 where  [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [sucursal]= " & Val(t_sucursal) & " and [letra]= '" & t_letra & "' and [num_comp]= " & Val(t_numcomp)
 End If
 
 rs.MaxRecords = 1
 rs.Open q, cn1
 ni = 0
 If Not rs.EOF And Not rs.BOF Then
      ni = rs("num_int")
 End If
 Set rs = Nothing
 
 If ni <> 0 Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     Set rs = New ADODB.Recordset
     q = "select [fecha], [fecha_vto], [cotizacion_dolar], [id_cliente], [num_int], [id_vendedor], [subtotal], [impuestos], [total], [perc_ib], [perc_gan], [perc_iva], [iva], " & _
     " [contado], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [observaciones] from vta_02 where [num_int] = " & ni
     rs.MaxRecords = 1
     rs.Open q, cn1
  
     t_fecha = rs("fecha")
     t_fechavto = rs("fecha_vto")
     t_cotizacion = rs("cotizacion_dolar")
     
     c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select [id_producto], [descripcion], [cantidad], [tunidad], [pu], [tasaiva], [importe], [pu_final], [tasaib], [num_int], [renglon] from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     Call armagrid
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("tunidad") & Chr$(9) & Format$(rs1("pu"), "######0.00") & Chr(9) & rs1("tasaiva") & Chr(9) & rs1("importe") & Chr(9) & rs1("pu_final") & Chr(9) & rs1("tasaib")
        
        Set rs2 = New ADODB.Recordset
        q = "select [desc_ext], [cant_lineas] from vta_015 where [num_int] = " & rs1("num_int") & " and [renglon] = " & rs1("renglon")
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
     t_subtotal = Format$(rs("subtotal"), "######0.00")
     t_nograbado = Format$(rs("impuestos"), "######0.00")
     t_perc = Format$(rs("perc_iva") + rs("perc_gan") + rs("perc_ib"), "######0.00")
     t_iva = Format$(rs("iva"), "######0.00")
     t_total = Format$(rs("total"), "######0.00")
     If Not IsNull(rs("observaciones")) Then
            t_observaciones = rs("observaciones")
     Else
       t_observaciones = "*"
     End If
     vta_formapago.armagrid2
     If rs("contado") = "S" Then
        vta_clientes.t_cli = rs("cliente02")
        vta_clientes.t_direccion = rs("direccion02")
        vta_clientes.t_cuit = rs("cuit02")
        vta_clientes.t_localidad = rs("localidad02")
        vta_clientes.c_iva.ListIndex = buscaindice(vta_clientes.c_iva, rs("id_tipo_iva02"))
        Option2 = True
        
        
    End If
     
     
  
  
   
   'cargo percepciones
     Set rs2 = New ADODB.Recordset
     q = "select vta_016.id_percepcion, descripcion, importe, vta_016.id_cuenta  from vta_016, a12 where [num_int] = " & rs("num_int") & " and vta_016.id_percepcion = a12.id_percepcion"
     rs2.Open q, cn1
    
     ABM_COMP_COMPRA2.armagrid
     i = 1
     While Not rs2.EOF
       ABM_COMP_COMPRA2.msf1.AddItem i & Chr$(9) & rs2("id_percepcion") & Chr$(9) & rs2("descripcion") & Chr$(9) & rs2("importe") & Chr$(9) & rs2("id_cuenta")
       rs2.MoveNext
       i = i + 1
     Wend
     Set rs2 = Nothing
      
      
      Set rs = Nothing
   
  
  Else
     EXISTE = "N"
  End If
  
  
End Sub

Sub carga3()
  Set rs = New ADODB.Recordset
  q = "select [num_int], [id_tipocomp], [sucursal], [letra]  from vta_02 where  [num_comp]= " & Val(t_numcomp)
  rs.Open q, cn1
  ni = 0
  If Not rs.EOF And Not rs.BOF Then
      ni = 0
      
      While Not rs.EOF And ni = 0
       If para.numeracion_comun_Fact_nc = "S" Then
          If rs("id_tipocomp") < 10 And rs("sucursal") = Val(t_sucursal) And rs("letra") = t_letra Then
            ni = rs("num_int")
         End If
      Else
         If rs("id_tipocomp") = c_tipocomp.ItemData(c_tipocomp.ListIndex) And rs("sucursal") = Val(t_sucursal) And rs("letra") = t_letra Then
            ni = rs("num_int")
         End If
      
      End If
         rs.MoveNext
      Wend
   Else
    ni = 0
   End If
 Set rs = Nothing
 
 If ni <> 0 Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     Set rs = New ADODB.Recordset
     q = "select [fecha], [fecha_vto], [cotizacion_dolar], [id_cliente], [num_int], [id_vendedor], [subtotal], [impuestos], [total], [perc_ib], [perc_gan], [perc_iva], [iva], " & _
     " [contado], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [observaciones] from vta_02 where [num_int] = " & ni
     rs.MaxRecords = 1
     rs.Open q, cn1
  
     t_fecha = rs("fecha")
     t_fechavto = rs("fecha_vto")
     t_cotizacion = rs("cotizacion_dolar")
     
     c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select [id_producto], [descripcion], [cantidad], [tunidad], [pu], [tasaiva], [importe], [pu_final], [tasaib], [num_int], [renglon] from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     Call armagrid
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("tunidad") & Chr$(9) & Format$(rs1("pu"), "######0.00") & Chr(9) & rs1("tasaiva") & Chr(9) & rs1("importe") & Chr(9) & rs1("pu_final") & Chr(9) & rs1("tasaib")
        
        Set rs2 = New ADODB.Recordset
        q = "select [desc_ext], [cant_lineas] from vta_015 where [num_int] = " & rs1("num_int") & " and [renglon] = " & rs1("renglon")
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
     t_subtotal = Format$(rs("subtotal"), "######0.00")
     t_nograbado = Format$(rs("impuestos"), "######0.00")
     t_perc = Format$(rs("perc_iva") + rs("perc_gan") + rs("perc_ib"), "######0.00")
     t_iva = Format$(rs("iva"), "######0.00")
     t_total = Format$(rs("total"), "######0.00")
     If Not IsNull(rs("observaciones")) Then
            t_observaciones = rs("observaciones")
     Else
       t_observaciones = "*"
     End If
     vta_formapago.armagrid2
     If rs("contado") = "S" Then
        vta_clientes.t_cli = rs("cliente02")
        vta_clientes.t_direccion = rs("direccion02")
        vta_clientes.t_cuit = rs("cuit02")
        vta_clientes.t_localidad = rs("localidad02")
        vta_clientes.c_iva.ListIndex = buscaindice(vta_clientes.c_iva, rs("id_tipo_iva02"))
        Option2 = True
        
        
    End If
     
     
  
  
   
   'cargo percepciones
     Set rs2 = New ADODB.Recordset
     q = "select vta_016.id_percepcion, descripcion, importe, vta_016.id_cuenta  from vta_016, a12 where [num_int] = " & rs("num_int") & " and vta_016.id_percepcion = a12.id_percepcion"
     rs2.Open q, cn1
    
     ABM_COMP_COMPRA2.armagrid
     i = 1
     While Not rs2.EOF
       ABM_COMP_COMPRA2.msf1.AddItem i & Chr$(9) & rs2("id_percepcion") & Chr$(9) & rs2("descripcion") & Chr$(9) & rs2("importe") & Chr$(9) & rs2("id_cuenta")
       rs2.MoveNext
       i = i + 1
     Wend
     Set rs2 = Nothing
      
      
      Set rs = Nothing
   
  
  Else
     EXISTE = "N"
  End If
  
  
End Sub




Private Sub btnacepta_Click()
If msf1.Rows - 1 <= Val(t_cantlineas) Then
 
If Option2 = True Then
 'contado
   If estadocaja(t_fecha) = "A" Then
    If Val(vta_formapago.t_diferencia) = 0 And vta_formapago.msf2.Rows > 1 Then
       Call iniciagraba
    Else
       If vta_formapago.msf2.Rows <= 1 Then
           J = MsgBox("No ha ingresado forma de pago, acepta pago total en Efectivo", 4)
           If J = 6 Then
              'pone forma de pago efectivo
              vta_formapago.msf2.AddItem "001" & Chr(9) & 1 & Chr(9) & "-" & Chr(9) & "Efectivo $" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(Val(t_total), "######0.00") & Chr(9) & Format$(t_fecha, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & para.cuenta_caja
              Call iniciagraba
           Else
              vta_formapago.Show
              vta_formapago.t_total = t_total
           End If
       Else
          MsgBox ("El pago ingresado no coincide con el total del comprobante")
       End If
    End If
   Else
     MsgBox ("Caja Cerrada. Imposible ingresar movimientos de contado en la fecha indicada")
   End If
  
Else
  'ctacte
  If c_prov.ListIndex > 0 Then
    'verifca credito
    If verificacredito Then
      Call iniciagraba
    End If
  
  Else
    MsgBox ("El Cliente Manual solo puede utilizarse para facturacion de contado")
  End If
End If

Else
  MsgBox ("La cantidad de lineas del comprobante supera el maximo permitido. Imposible emitir comprobante")
End If


End Sub

Function verificacredito() As Boolean
vc = True
If vta_facturacion.c_tipocomp.ListIndex = 0 Then
    If vta_facturacion.Option4 Then
     'pesos
     tpl = Val(vta_facturacion.t_total)
    Else
     tpl = Val(vta_facturacion.T_total2)
    End If
     
    If Val(vta_clientes.t_saldo1) + tpl > Val(vta_clientes.t_limite) Then
      If t_cl = False Then
         J = MsgBox("El comprobante actual ha superado el LIMITE de CREDITO establecido para el cliente. ¿Acepta Emision?", 4)
         If J = 6 Then
           vc = True
         Else
           vc = False
         End If
     Else
       MsgBox ("El comprobante actual ha superado el LIMITE de CREDITO establecido para el cliente. Imposible generar comprobante")
       vc = False
    End If
   End If
End If
verificacredito = vc
End Function
Sub iniciagraba()



If Val(t_total) > 0 Then
  If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 25 And t_letra <> "E" Then
    MsgBox ("Para realizar Facturas de Pro Forma es necesario que el cliente sea de Exportacion")
    c_tipocomp.SetFocus
    Exit Sub
  End If


 Call mensaje
 J = MsgBox("Graba " & Label20, 4)
 If J = 6 Then
  
  
  If verificaperiodog(t_fecha) = "A" Then
   
   If Val(t_sucursal) = glo.sucursalf Then
     
    
       
       Call fiscal
     
   Else
    If Val(t_sucursal) = glo.sucursale Then
     para.z_actual = 0
     If EXISTE = "S" Then
       k = MsgBox("Los comprobantes electronicos no se pueden modificar, el numero de comrobante actual existe, si usted continua se creara un nuevo comprobante", 4)
       If k = 6 Then
          EXISTE = "N"
          Call electronica
       End If
     Else
       Call electronica
     End If
    Else
     para.z_actual = 0
     Call normal
    End If
   End If
   
   'vuelve a cta cte
   Option1 = True
   
  Else
   MsgBox ("Periodo Cerrado. Imposible grabar comprobante")
  End If
 End If
Else
 MsgBox ("Imposible emitir comprobante. El total del comprobante debe ser > 0 ")
End If
  

End Sub
Function verificafechafiscal() As Boolean
'verifica horario fiscal
If para.fiscal <> 0 Then
  r = epson1.SetGetDateTime("G")
  If r = True Then
     F = epson1.AnswerField_3
     h = epson1.AnswerField_4
     ff = Format$(Mid$(F, 5, 2), "00") & "/" & Format$(Mid$(F, 3, 2), "00") & "/" & Format$(Mid$(F, 1, 2), "00")
    ' hf = Format$(Mid$(h, 1, 2), "00") & ":" & Format$(Mid$(h, 3, 2), "00") & ":" & Format$(Mid$(h, 5, 2), "00")
     If DateValue(ff) <> DateValue(t_fecha) Then
       g = MsgBox("La fecha de la impresora fiscal y del comprobante son diferentes. Fecha impreso fiscal " & ff & ". Continua sin correjir", 4)
       If g = 6 Then
          verificafechafiscal = True
       Else
          verificafechafiscal = False
       End If
     Else
       verificafechafiscal = True
     End If
  Else
   MsgBox ("La Impresora Fiscal esta desconectada o no se encuentra. Verifique y vuelva a intentar")
   verificafechafiscal = False
  End If
End If
End Function
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
Sub fiscal()

estadograba = 0
Call bloquea_comp
If Option4 = True Then
 J = MsgBox("Imprime Comprobante Fiscal", 4)
 If J = 6 Then
  
  seguir = True
  While seguir
    Set cl_fiscal = New fiscal
    cl_fiscal.carga (glo.sucursalf)
    
     If cl_fiscal.idmodelo = 24 Then 'tm-900 then
          'if vta_clientes.c_iva
             resulta = imprime_facturafiscal2
             
       
       Else
             resulta = imprime_facturafiscal
       End If
    Set cl_fiscal = Nothing
      
    If resulta Then
        espere.ProgressBar1.Value = 5
        espere.Label1 = "Espere... Grabando Comprobante Fiscal"
        If Val(t_total) <= 0 Or Val(t_numcomp) <= 0 Then
          seguir = False
          estadograba = 1
        Else
            Set rs = New ADODB.Recordset
            q = " select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
            rs.Open q, cn1
            If Not rs.BOF And Not rs.EOF Then
                MsgBox ("Problema detectado!!! El comprobante generado por el impresor fiscal ya existe en el sistema")
                estadograba = 1
                Set rs = Nothing
            Else
             Set rs = Nothing
             Call graba
            End If
            seguir = False
        End If
    Else
        J = MsgBox("Error al Imprimir el Comprobante. Verifique la Impresora para continuar.  Reintente o Cancele", 5)
        If J = 4 Then
           seguir = True
        Else
           seguir = False
           estadograba = 1
        End If
    End If
    Unload espere
  Wend
  If estadograba = 1 Then
     MsgBox ("El comprobante Fiscal ha tenido problemas y no pudo grabarse. Si el impresor termino de emitir el comprobante ingreselo por comprobantes manuales, sino vuelva a emitirlo por el controlador fiscal")
  End If
Else
  'graba pero no emite
   Set rs = New ADODB.Recordset
   q = " select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
   rs.Open q, cn1
   If Not rs.BOF And Not rs.EOF Then
      ni = rs("num_int")
      Set rs = Nothing
      k = MsgBox("El comprobante existe, desea modifiarlo", 4)
      If k = 6 Then
         'borrar y grabar
         If para.id_grupo_modulo_actual >= 8 Then
           Set cl_compvta = New comprobantes_venta
           cl_compvta.cargar2 (ni)
           cl_compvta.borrar
           Set cl_compvta = Nothing
           Call graba
         Else
           MsgBox ("El comprobante existe y Ud. no tiene permisos para modificarlo")
         End If
       End If
   Else
       Set rs = Nothing
       Call graba
   End If
End If
Else
  MsgBox ("La facturacion fiscal no admite comprobantes en U$s")
End If
Call libera_comp
End Sub

Sub bloquea_comp()
Frame3.Enabled = False
Frame5.Enabled = False
Frame2.Enabled = False
Frame6.Enabled = False
End Sub
Sub libera_comp()
Frame3.Enabled = True
Frame5.Enabled = True
Frame2.Enabled = True
Frame6.Enabled = True
End Sub
Function imprime_facturafiscal2() As Boolean
Dim a(5) As String

Dim CUIT As String
Dim identifica As String
Dim tpago As String
Dim t  As String
Dim de1 As String
Dim tipocompfz As String
Dim tv2 As String
Dim td As String
Dim cliz As String
Dim dirz As String
Dim locz As String
Dim de1z As String
Dim tivacz As String
Dim letraz As String
Dim rk As Boolean
Dim remitosz As String
Dim remitosz2 As String
'Dim r As Boolean
Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
para.z_actual = cl_fiscal.ultimo_z + 1
Select Case c_tipocomp.ItemData(c_tipocomp.ListIndex)
Case Is = 1
   If t_letra = "A" Then
     tipocompfz = 1  'fact A
   Else
      tipocompfz = 2 'fact b
   End If
Case Is = 2
    If cl_fiscal.imprimend = "S" Then
       If t_letra = "A" Then
          tipocompfz = 4 'nd A
       Else
          tipocompfz = 5 'nd B
       End If
    Else
        MsgBox ("La impresora fiscal no puede imprimir ND")
        imprime_facturafiscal2 = False
        Exit Function
    End If
 Case Is = 3
    If cl_fiscal.imprimenc = "S" Then
       If t_letra = "A" Then
           tipocompfz = 7
       Else
           tipocompfz = 8
       End If
    Else
        MsgBox ("La impresora fiscal no puede imprimir NC")
        imprime_facturafiscal2 = False
         Exit Function
    End If
 Case Else
    para.z_actual = 0
    imprime_facturafiscal2 = False
    Exit Function
End Select
caracteresmax = cl_fiscal.caracteresmax
Set cl_fiscal = Nothing


'copias
'cantidad copias
Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & Val(t_sucursal) & " and  [id_tipocomp] = " & 10
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
   copias21 = rs("cant_copias_b")
Else
   copias21 = 1
End If
Set rs = Nothing




espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 6
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Comprobando Impresora"
'abrir factura
If vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex) <> 3 Then
   identifica = 0 'cuit
   'CUIT = Mid$(vta_clientes.t_cuit, 1, 11) '& Mid$(vta_clientes.t_cuit, 4, 8) & Mid$(vta_clientes.t_cuit, 13, 1)
    CUIT = RTrim$(vta_clientes.t_cuit)
 Else
   identifica = 1 'ninguno
   CUIT = RTrim$(vta_clientes.t_cuit)
 End If
 
 If Option1 = True Then
    tpago = "Cta.Cte. Nro. " & Format$(c_prov.ItemData(c_prov.ListIndex), "00000")
 Else
    tpago = "CONTADO"
 End If

   tv2 = " "

 espere.ProgressBar1.Value = 2
 espere.Label1 = "Espere... Abriendo Comprobante Fiscal:" & c_tipocomp
 
remitosz = " "
remitosz2 = " "

For i = 1 To vta_selremitos.msf1.Rows - 1
   If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
       remitosz = remitosz & "#" & Val(Mid$(vta_selremitos.msf1.TextMatrix(i, 1), 6, 8))
   End If
Next i
      
 
 'On Error GoTo errf
 cliz = textofiscal(Left$(vta_clientes.t_cli & " ", caracteresmax))
 dirz = textofiscal(Left$(vta_clientes.t_direccion & " ", caracteresmax))
 locz = textofiscal(Left$(vta_clientes.t_localidad & " ", caracteresmax))
 letraz = t_letra
 tivacz = vta_clientes.t_codfiscal2
 
 
 'abrir factura
 
 On Error GoTo DepuraErrores
 If Not Fiscaltf.Inicializar Then
    Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
  End If
  
  Fiscaltf.CancelarComprobante
    
  
  
 'datos del cliente
 If Not Fiscaltf.DatosCliente(cliz, identifica, CUIT, tivacz, dirz) Then
      Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
 End If
     
  If remitosz <> " " Then
       If Not Fiscaltf.ImprimirTextoNoFiscal("Rtos:" & remitosz) Then
          Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
       End If
  End If
  
  
  
  If Not Fiscaltf.AbrirComprobante(tipocompfz) Then
     Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
  End If
  
  
   
'envia items a facturar
espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Imprimiendo Productos"
 
 i = 1
 While i < msf1.Rows
     
        
         If Val(msf1.TextMatrix(i, 0)) = 0 Then
          If Check2 = 1 Then 'tiene desc extra
            de1 = " "
            dex = msf1.TextMatrix(i, 2)
            
            Call lee_desc_extra(a, dex)
            
            For k = 0 To 2
             If a(k) <> "%%" Then
                 de1 = Left$(a(k), caracteresmax)
                 de1z = textofiscal(de1)
                 If Not Fiscaltf.ImprimirTextoFiscal(delz) Then
                   Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
                 End If
                
             Else
               k = 2
             End If
           Next k
          End If
        Else
          de1z = textofiscal(Left$(msf1.TextMatrix(i, 2), caracteresmax))
          
          If t_letra = "A" Then
            precio = Val(msf1.TextMatrix(i, 5))
          Else
            precio = Val(msf1.TextMatrix(i, 8))
          End If
          
          If Not Fiscaltf.ImprimirItem2g(de1z, Val(msf1.TextMatrix(i, 3)), precio, Val(msf1.TextMatrix(i, 6)), 0, IFUniversal.Gravado, 0, 1, msf1.TextMatrix(i, 1), "", 0) Then
             Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
          End If
        End If
      
     
   i = i + 1
 Wend
 
 
 'pagos
  espere.Label1 = "Espere... Grabando Pagos"
  
  
  t_subtotal = Fiscaltf.subtotal.MontoNeto
  t_iva = Fiscaltf.subtotal.MontoIVA
  t_total = Fiscaltf.subtotal.MontoVentas
 
  
  
  If Option2 = True Then 'contado
  
    resto = Val(t_total)
  For i = 1 To fsc_formapago.msf2.Rows - 1
     td = Left$(RTrim$(fsc_formapago.msf2.TextMatrix(i, 2)), 15)
     mp = Format$(Val(fsc_formapago.msf2.TextMatrix(i, 6)), "######0.00")
     dp = "T"
        
     Set rs2 = New Recordset
     q = "select * from cyb_01 where [id_forma_pago] = " & Val(fsc_formapago.msf2.TextMatrix(i, 0))
     rs2.Open q, cn1
     If Not rs2.EOF And Not rs2.BOF Then
        codpago = rs2("codigo_driver_fiscal")
     Else
       codpago = 8
     End If
     Set rs2 = Nothing
        
     If Not Fiscaltf.ImprimirPago2g(td, mp, "", codpago, 1, "", "") Then
       Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
     End If
     
     resto = resto - mp
   Next i

  
  If resto > 0 Then
      If Not Fiscaltf.ImprimirPago2g("Pago", Format$(resto, "######0.00"), "", IFUniversal.CuentaCorriente, 1, "", "") Then
       Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
      End If
     
  End If
    
    
  
  Else
    td = "Cta. Cte. Nro. " & Format$(c_prov.ItemData(c_prov.ListIndex), "00000")
    mp = Val(t_total)
    dp = "T"
    If Not Fiscaltf.ImprimirPago2g(td, Format$(mp, "######0.00"), "", IFUniversal.CuentaCorriente, 1, "", "") Then
       Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
    End If
     
    
    
  End If
  
 'subtotal para obtener el importe neto, iva y total impreso en la factura
espere.ProgressBar1.Value = 4
espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

      
      
      
 
  espere.Label1 = "Espere Cerrando Tique...."
  espere.Label1.Refresh
  Fiscaltf.CerrarComprobante
  
  t_numcomp = Format$(Fiscaltf.UltimoComprobante(tipocompfz), "00000000")
 
    
  'copias
   l = InputBox("Indique cantidad de Copias", , copias21)
   If Val(l) > 0 And Val(l) <= 6 Then
        For Y = 1 To Val(l)
            If Fiscaltf.CopiarComprobante(tipocompfz, Val(t_numcomp)) Then
                     'Err.Raise Fiscaltf.Error, "", Fiscaltf.ErrorDesc
            End If
        Next
   End If
   
  Fiscaltf.Finalizar
  
  imprime_facturafiscal2 = True
  
    
    
 Exit Function
DepuraErrores:
  'Fiscaltf.Finalizar
  MsgBox Fiscaltf.ErrorDesc
  imprime_facturafiscal2 = False
  Exit Function
   
End Function



Function imprime_facturafiscal() As Boolean
Dim a(5) As String

Dim CUIT As String
Dim identifica As String
Dim tpago As String
Dim t  As String
Dim de1 As String
Dim tipocompfz As String
Dim tv2 As String
Dim td As String
Dim cliz As String
Dim dirz As String
Dim locz As String
Dim de1z As String
Dim tivacz As String
Dim letraz As String
Dim rk As Boolean
Dim remitosz As String
Dim remitosz2 As String
'Dim r As Boolean
Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
para.z_actual = cl_fiscal.ultimo_z + 1
Select Case c_tipocomp.ItemData(c_tipocomp.ListIndex)
Case Is = 1
     tipocompfz = cl_fiscal.CODFACT 'codigo para tique fact y fact
Case Is = 2
    If cl_fiscal.imprimend = "S" Then
        tipocompfz = "D"
    Else
        MsgBox ("La impresora fiscal no puede imprimir ND")
        imprime_facturafiscal = False
        Exit Function
    End If
 Case Is = 3
    If cl_fiscal.imprimenc = "S" Then
        tipocompfz = cl_fiscal.CODNC
    Else
        MsgBox ("La impresora fiscal no puede imprimir NC")
        imprime_facturafiscal = False
         Exit Function
    End If
 Case Else
    para.z_actual = 0
    imprime_facturafiscal = False
    Exit Function
End Select
caracteresmax = cl_fiscal.caracteresmax
Set cl_fiscal = Nothing




espere.Show
espere.Refresh
espere.ProgressBar1.Min = 0
espere.ProgressBar1.Max = 6
espere.ProgressBar1.Value = 1
espere.Label1 = "Espere... Comprobando Impresora"
'abrir factura
If vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex) <> 3 Then
   identifica = "CUIT"
   'CUIT = Mid$(vta_clientes.t_cuit, 1, 11) '& Mid$(vta_clientes.t_cuit, 4, 8) & Mid$(vta_clientes.t_cuit, 13, 1)
    CUIT = RTrim$(vta_clientes.t_cuit)
 Else
   identifica = "DNI"
   CUIT = RTrim$(vta_clientes.t_cuit)
 End If
 
 If Option1 = True Then
    tpago = "Cta.Cte. Nro. " & Format$(c_prov.ItemData(c_prov.ListIndex), "00000")
 Else
    tpago = "CONTADO"
 End If

If vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex) = 4 Then
   If caracteresmax > 37 Then
      tv2 = "Receptor del comprobante: Responsable Monotributo"
   Else
      tv2 = "Receptor: Monotributo"
   End If
Else
   tv2 = " "
End If

 'Call NULOS(t_remito)
 espere.ProgressBar1.Value = 2
 espere.Label1 = "Espere... Abriendo Comprobante Fiscal:" & c_tipocomp
 
remitosz = " "
remitosz2 = " "

For i = 1 To vta_selremitos.msf1.Rows - 1
   If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
      'If Len(remitosz & "#" & CStr(Val(Mid$(vta_selremitos.msf1.TextMatrix(i, 1), 6, 8)))) + 8 <= caracteresmax Then
       remitosz = remitosz & "#" & Val(Mid$(vta_selremitos.msf1.TextMatrix(i, 1), 6, 8))
      'Else
      ' remitosz2 = remitosz2 & "#" & CStr(Val(Mid$(vta_selremitos.msf1.TextMatrix(i, 1), 6, 8)))
      'End If
   End If
Next i
      
 
 'On Error GoTo errf
 cliz = textofiscal(Left$(vta_clientes.t_cli & " ", caracteresmax))
 dirz = textofiscal(Left$(vta_clientes.t_direccion & " ", caracteresmax))
 locz = textofiscal(Left$(vta_clientes.t_localidad & " ", caracteresmax))
 letraz = t_letra
 tivacz = vta_clientes.t_codfiscal
 rk = epson1.OpenInvoice(tipocompfz, "C", letraz, "1", "P", "17", "I", tivacz, cliz, " ", identifica, CUIT, "N", dirz, locz, tpago, Left$("Remitos:" & remitosz, caracteresmax), tv2, "C")
 
 'rk = epson1.OpenInvoice(tipocompfz, "C", letraz, "1", "P", "17", "I", "I", cliz, "otro", identifica, "20202956034", "N", "Pellegrini 304", "dir2", "dir3", "001", "002", "C")
 If rk Then

Else
   Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
   'MsgBox (tipocompf & " " & t_letra & " " & vta_clientes.t_codfiscal)
  ' MsgBox ("Error F001 al Inicializar Comprobante. Estado Fiscal: " & epson1.FiscalStatus & "  Estado Impresor. " & epson1.PrinterStatus)

End If
 'envia items a facturar
espere.ProgressBar1.Value = 3
espere.Label1 = "Espere... Imprimiendo Productos"
 
 i = 1
 While rk And i < msf1.Rows
     
      If rk Then
        
          If Val(msf1.TextMatrix(i, 0)) = 0 Then
           If Check2 = 1 Then 'tiene desc extra
            de1 = " "
            dex = msf1.TextMatrix(i, 2)
            
            Call lee_desc_extra(a, dex)
            
            For k = 0 To 2
             If a(k) <> "%%" Then
                 de1 = Left$(a(k), caracteresmax)
                 de1z = textofiscal(de1)
                If t_letra = "A" Then
                  r = epson1.SendInvoiceItem(de1z, "00001000", "000000000", "0000", "M", "0", "0", " ", " ", " ", "0", "0")
                Else
                  r = epson1.SendInvoiceItem(de1z, "00001000", "000000000", "0000", "M", "0", "0", " ", " ", " ", "0", "0")
                End If
             Else
               k = 2
             End If
            Next k
         End If
        Else
          de1z = textofiscal(Left$(msf1.TextMatrix(i, 2), caracteresmax))
          If t_letra = "A" Then
            pu = Val(msf1.TextMatrix(i, 5))
            If pu >= 0 Then
              ci = "M"
            Else
              ci = "R"
              pu = -pu
            End If
                        
            rk = epson1.SendInvoiceItem(de1z, Format$(Val(msf1.TextMatrix(i, 3)) * 1000, "00000000"), Format$(pu * 100, "000000000"), Format$(Val(msf1.TextMatrix(i, 6)) * 100, "0000"), Format$(ci, "@"), "0", "0", " ", " ", " ", "0", "0")
          Else
            
            pu = Val(msf1.TextMatrix(i, 8))
            If pu >= 0 Then
              ci = "M"
            Else
              ci = "R"
              pu = -pu
            End If
            
            rk = epson1.SendInvoiceItem(de1z, Format$(Int(Val(msf1.TextMatrix(i, 3)) * 1000), "00000000"), Format$(pu * 100, "000000000"), Format$(Val(msf1.TextMatrix(i, 6)) * 100, "0000"), Format$(ci, "@"), "0", "0", " ", " ", " ", "0", "0")
          End If
      
        End If
      
      Else
        Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
        i = msf1.Rows
      End If
   i = i + 1
 Wend
 
 
 'pagos
  espere.Label1 = "Espere... Grabando Pagos"
  
  If Option2 = True Then 'contado
  
     td = Left$(RTrim$(vta_formapago.msf2.TextMatrix(1, 1)) & " " & RTrim$(vta_formapago.msf2.TextMatrix(1, 3)), caracteresmax)
    mp = Format$(Val(t_total) * 100, "00000000")
    dp = "T"
    If rk Then
       rk = epson1.SendInvoicePayment(td, Format$(Val(t_total) * 100, "00000000"), "T")
     Else
       Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
     End If
  
  Else
    td = "Cta. Cte. Nro. " & Format$(c_prov.ItemData(c_prov.ListIndex), "00000")
    mp = Format$(Val(t_total) * 100, "00000000")
    dp = "T"
    If rk Then
       rk = epson1.SendInvoicePayment("Cta. Cte. Nro. " & Format$(c_prov.ItemData(c_prov.ListIndex), "00000"), Format$(Val(t_total) * 100, "00000000"), "T")
     Else
       Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
     End If
  End If
  
 'subtotal para obtener el importe neto, iva y total impreso en la factura
espere.ProgressBar1.Value = 4
espere.Label1 = "Espere... Cerrando Comprobante Fiscal"

 If rk Then
    rk = epson1.GetInvoiceSubtotal("N", "xx")
 Else
   Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
 End If
 If rk Then
      t_subtotal = Format$(Val(epson1.AnswerField_10) / 100, "######0.00")
      t_iva = Format$(Val(epson1.AnswerField_6) / 100, "####0.00")
      t_total = Format$(Val(epson1.AnswerField_5) / 100, "######0.00")
 Else
     Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
 End If
 
 
  If rk Then rk = epson1.CloseInvoice(tipocompfz, letraz, " ")
   
  If rk Then
     t_numcomp = epson1.AnswerField_3
  Else
     Call verificaerrfiscal(epson1.FiscalStatus, epson1.PrinterStatus)
  End If
  imprime_facturafiscal = rk
    
 Exit Function
errf:
 MsgBox ("Error al comunicarse con el impresor fiscal. Verifique que esta encendido y reintente")
 Exit Function
   
End Function


Sub normal()
  Set rs = New ADODB.Recordset
  q = "select num_int from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
      EXISTE = "S"
      If para.id_grupo_modulo_actual >= 8 Then
         ni = rs("num_int")
         
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
  Set rs = Nothing

End Sub
Private Sub btnsale_Click()
J = MsgBox("Abandona el comprobante (S/N)", 4)
If J = 6 Then
   
    Unload Me
End If
End Sub

Sub cerra_all()


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


Private Sub c_actividad_LostFocus()
If c_actividad.ListIndex < 0 Then
  c_actividad.ListIndex = 0
End If
End Sub

Private Sub c_prov_LostFocus()
vta_selremitos.limpia
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
    c_vend.ListIndex = buscaindice(c_vend, vta_clientes.t_idvend)
  Else
    c_prov.ListIndex = 0
  End If
End If

If c_prov.ItemData(c_prov.ListIndex) = 1 Then
      Option2 = True
      vta_clientes.t_id = 1
       t_alicuotaib = 0
      vta_clientes.carga
      vta_clientes.Show
Else
    Call iniciacli
    
End If


End Sub


Sub inicia()

 
   t_letra = vta_clientes.t_letrafact
   gcuit = vta_clientes.t_cuit
  
   Set cl_compvta = New comprobantes_venta
   cl_compvta.sucursal = Val(c_sucursal)
   cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
   cl_compvta.letra = t_letra
   cl_compvta.SACANUMCOMP
   t_numcomp = Format$(cl_compvta.numcomp, "00000000")
   cantlineas = cl_compvta.cant_lineas
   Set cl_compvta = Nothing
   t_cotizacion = para.cotizacion

   vta_formapago.armagrid2
   
   If para.calcula_perc_ib = "S" And t_letra = "A" Then
     Set cl_padronib = New padron_ib
     cl_padronib.cuit_texto = gcuit
     cl_padronib.buscar_perc
     t_alicuotaib = Format$(cl_padronib.tasa_percib, "##0.00")
     Select Case cl_padronib.estado_consulta
     Case Is = "OK"
       Label20 = "¡COMPROBANTE SUJETO A PERCEPCION IB! Consulta del Padron de IB Satistactoria"
     Case Is = "NO"
       Label20 = "¡ATENCION! El contribuyente NO se encuentra en el padron, se aplicará la tasa maxima para percpciones de IB. Verifique si corresponde"
     Case Is = "ER"
       Label20 = "¡CUIDADO! Numero de cuit con formato invalido. Padron NO consultado"
     End Select
     Frame11.Visible = True
     
     Set cl_padronib = Nothing
     
     
   Else
     t_alicuotaib = "0.00"
     T_PERCIB = "0.00"
     'gcuit = "0"
   End If
   Call armagrid
   Unload espere

   If Option2 = True Then
      Command8.Enabled = True
   Else
       Command8.Enabled = False
   End If
'Else
'  Unload espere
'  MsgBox ("Error. No se puedo Inicializa el Cliente")
'End If





End Sub

Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = buscaindice(c_sucursal, para.punto_venta_usuario)
End If
t_sucursal = Format$(c_sucursal, "0000")
t_numcomp = ""
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

Private Sub ch_plan_Click()

Call habilitaplandepago
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

Private Sub Command3_Click()
If c_tipocomp.ItemData(c_tipocomp.ListIndex) <> 3 And c_tipocomp.ItemData(c_tipocomp.ListIndex) <> 32 And c_tipocomp.ItemData(c_tipocomp.ListIndex) <> 2 And c_tipocomp.ItemData(c_tipocomp.ListIndex) <> 31 Then
    Set cl_compvta = New comprobantes_venta
    cl_compvta.sucursal = Val(c_sucursal)
    cl_compvta.actual (1)
    vta_selremitos.t_r2 = cl_compvta.cant_lineas
    Set cl_compvta = Nothing
    vta_selremitos.carga
    vta_selremitos.Show
Else
    vta_selcomp.carga
    vta_selcomp.Show
 
End If
End Sub

Private Sub Command4_Click()
Call CALCULATOTALES
End Sub

Private Sub Command5_Click()
vta_clientes.Show
End Sub

Private Sub Command6_Click()
ABM_COMP_COMPRA2.t_modulo = "P"
ABM_COMP_COMPRA2.Show
End Sub

Private Sub Command7_Click()
vta_facturacion2.t_modulo = "F"
vta_facturacion2.Show
End Sub

Private Sub Command8_Click()
 
  vta_formapago.Show
  vta_formapago.t_total = t_total
 
End Sub

Private Sub Form_Activate()
Frame2.Enabled = False

End Sub
Sub captura()
MsgBox ("Estado Fiscal: " & epson1.FiscalStatus & "  Estado Impresora: " & epson1.PrinterStatus)
End Sub
Sub grabaformapago()
  For i = 1 To vta_formapago.msf2.Rows - 1
         If Val(vta_formapago.msf2.TextMatrix(i, 0)) = 3 Then
                'ch. terceros
                
                QUERY = "insert into cyb_03 (fecha_emision, num_cheque, banco, sucursal, titular, importe,   "
                QUERY = QUERY & "estado, fecha_dif, origen, destino, num_mov_banco_i, num_mov_banco_e, num_int_op, num_int_rbo, fecha_salida, fecha_ingreso,tipo_salida) values("
                QUERY = QUERY & "'" & t_fecha & "', " & Val(msf2.TextMatrix(i, 2)) & ",'" & msf2.TextMatrix(i, 3) & "', '" & msf2.TextMatrix(i, 4) & "', '" & msf2.TextMatrix(i, 5) & "', "
                QUERY = QUERY & Val(msf2.TextMatrix(i, 6)) & ",'C', '" & vta_formapago.msf2.TextMatrix(i, 7) & "', '" & Left$(vta_clientes.t_cli, 50) & "', ' ', 0,0,0, " & numint & ", '" & t_fecha & "', '" & t_fecha & "', 'C')"
                'MsgBox (QUERY)
                cn1.Execute QUERY
                
                
                qr = "SELECT @@IDENTITY AS NewID"
                Set rs = cn1.Execute(qr)
                numintch = rs.Fields("NewID").Value

                
                Set rs = Nothing
         
         Else
           numintch = 0
         End If
         
         
         If Val(vta_formapago.msf2.TextMatrix(i, 0)) = 4 Then
                q = "select * from cyb_04"
                Set rs = New ADODB.Recordset
                rs.Open q, cn1, adOpenDynamic, adLockOptimistic
                rs.AddNew
                 rs("id_banco") = Val(vta_formapago.msf2.TextMatrix(i, 8))
                 rs("fecha") = vta_formapago.msf2.TextMatrix(i, 7)
                 rs("importe") = Val(vta_formapago.msf2.TextMatrix(i, 6))
                 rs("id_tipomov") = 60 'transf
                 rs("fecha_dif") = vta_formapago.msf2.TextMatrix(i, 7)
                 rs("ubicacion") = "H"
                 rs("entro") = "N"
                 rs("fecha_acreed") = vta_formapago.msf2.TextMatrix(i, 7)
                 rs("num_comp") = Val(vta_formapago.msf2.TextMatrix(i, 2))
                 rs("detalle") = "Transf." & Left$(vta_formapago.msf2.TextMatrix(i, 5), 30)
                 rs("modulo") = "V"
                 rs("num_mov_int") = numint
                 rs("id_tipodbcr") = 1
                rs.Update
                
                Set rs = Nothing
         End If
         
         
         q = "select caja, id_cuenta_cont from cyb_01 where [id_forma_pago] = " & Val(vta_formapago.msf2.TextMatrix(i, 0))
         Set rs = New ADODB.Recordset
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
          If rs("CAJA") = "S" Then
             If ubicacionctacte <> "D" Then
                importe = -Val(vta_formapago.msf2.TextMatrix(i, 6))
             Else
                importe = Val(vta_formapago.msf2.TextMatrix(i, 6))
             End If
             
             ctach = rs("id_cuenta_cont")
             QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
             QUERY = QUERY & " VALUES (" & ctach & ", " & cuentaact & ", '" & RTrim$(Left$(vta_clientes.t_cli, 49)) & " ', " & importe & ", 'D', '" & t_fecha & "', " & numint & ", 'V', '" & Left$(abreviatura, 5) & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', " & Val(vta_formapago.msf2.TextMatrix(i, 0)) & ", " & numintch & ", " & para.id_usuario & ")"
             cn1.Execute QUERY
          End If
         End If
         Set rs = Nothing

                 
        'formas de pago
        QUERY = "INSERT INTO vta_04([num_int], [secuencia], [id_formapago], [formapago], [num_ch], [detalle_banco], [sucursal], [titular], [importe], [fecha_dif], [num_int_fp])"
        QUERY = QUERY & " VALUES (" & numint & ", " & i & ", " & Val(vta_formapago.msf2.TextMatrix(i, 0)) & ", '" & Left$(RTrim$(vta_formapago.msf2.TextMatrix(i, 1)), 9) & " ', " & Val(vta_formapago.msf2.TextMatrix(i, 2)) & ", '" & RTrim$(vta_formapago.msf2.TextMatrix(i, 3)) & " ', '" & RTrim$(vta_formapago.msf2.TextMatrix(i, 4)) & " ', '" & RTrim$(vta_formapago.msf2.TextMatrix(i, 5)) & " ', " & Val(vta_formapago.msf2.TextMatrix(i, 6)) & ", '" & RTrim$(vta_formapago.msf2.TextMatrix(i, 7)) & " ', " & numintch & ")"
        cn1.Execute QUERY

      Next i

      
          

End Sub
Sub iniciacli()
 If c_prov.ListIndex > 0 Then
   vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
   vta_clientes.carga
   t_letra = vta_clientes.t_letrafact
   c_vend.ListIndex = buscaindice(c_vend, Val(vta_clientes.t_idvend))
 Else
   If Val(vta_clientes.t_id) <> 0 Then
      vta_clientes.t_id = 0
      vta_clientes.limpia
   
   
   End If
 End If
 
 If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 2 Or c_tipocomp.ItemData(c_tipocomp.ListIndex) = 3 Or c_tipocomp.ItemData(c_tipocomp.ListIndex) = 31 Or c_tipocomp.ItemData(c_tipocomp.ListIndex) = 32 Then
    Command3.Caption = "Facturas"
 Else
    Command3.Caption = "Remitos"
    
 End If
 
 Call habilitaplandepago
 
 
 
 
 
 If c_tipocomp.ItemData(c_tipocomp.ListIndex) >= 30 And c_tipocomp.ItemData(c_tipocomp.ListIndex) <= 32 Then
    c_transferencia.Visible = True
    Label15.Visible = True
 Else
    c_transferencia.Visible = False
    Label15.Visible = False
End If
 
 
End Sub
Sub habilitaplandepago()
 If ch_plan.Value = 0 Then
   Frame14.Visible = False
   Frame11.Visible = True
 Else
  If Option2 = False Then
   Frame14.Visible = True
   Frame11.Visible = False
   t_cantcuotas = 1
  Else
    MsgBox ("Los planes de cuotas deben realizarse en CUENTA CORRIENTE")
    Option1 = True
  End If
  
  Set rs0 = New ADODB.Recordset
  q = "select cuotas_sininteres, interes_cuota  from g0 where sucursal=0"
  rs0.Open q, cn1
  If Not rs0.EOF And Not rs0.BOF Then
    t_interes_cuota = rs0("interes_cuota")
    t_cuotas_sininteres = rs0("cuotas_sininteres")
  
  
  Else
    t_interes_cuota = 0
    t_cuotas_sininteres = 0
  
  End If
  
  Set rs0 = Nothing
  
  t_fechacuota1 = Format$(Now, "dd/mm/yyyy")

End If

End Sub
Sub actualizaremitos()
For J = 1 To msf1.Rows - 1
 If Val(msf1.TextMatrix(J, 1)) > 1 Then
  cantidadf = Val(msf1.TextMatrix(J, 3)) 'cantidad facturada
  codprodant = Val(msf1.TextMatrix(J, 1))  'cod producto
  i = 1  'para cada articulo busco en los remitos seleccionados
  While i < vta_selremitos.msf1.Rows
   If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
     nir = Val(vta_selremitos.msf1.TextMatrix(i, 4))
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
                i = vta_selremitos.msf1.Rows
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

'vuelvo a verificar todos los remitos
      
      i = 1  'buscolos remitos seleccionados
      While i < vta_selremitos.msf1.Rows
        If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
          nir = Val(vta_selremitos.msf1.TextMatrix(i, 4))
          If verificaremito(nir) = 0 Then
             q = "SELECT * FROM VTA_02 WHERE [NUM_INT] = " & nir
             Set rs = New ADODB.Recordset
             rs.Open q, cn1, adOpenDynamic, adLockOptimistic
             If Not rs.EOF And Not rs.BOF Then
                rs("estado") = "F"
                rs.Update
             End If
          End If
        End If
        i = i + 1
     Wend
          
          
          
          
End Sub



Sub actualizaremitos2()
For J = 1 To msf1.Rows - 1
 If Val(msf1.TextMatrix(J, 1)) > 1 Then
  cantidadf = Val(msf1.TextMatrix(J, 3)) 'cantidad facturada
  codprodant = Val(msf1.TextMatrix(J, 1))  'codigo producto
  i = 1  'para cada articulo busco en los remitos seleccionados
  While i < vta_selremitos.msf1.Rows
   If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
     nir = Val(vta_selremitos.msf1.TextMatrix(i, 4))
     q = "select cantidad from vta_03 where [num_int] = " & nir & " and [id_producto] = " & codprodant
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
                'rs1.MoveLast
                i = vta_selremitos.msf1.Rows
             End If
              
            rs1.MoveNext
      Wend
      Set rs1 = Nothing
   End If
     i = i + 1
  Wend
 End If
Next J


'vuelvo a verificar todos los remitos
      
      i = 1  'buscolos remitos seleccionados
      While i < vta_selremitos.msf1.Rows
        If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
          nir = Val(vta_selremitos.msf1.TextMatrix(i, 4))
          If verificaremito(nir) = 0 Then
             q = "SELECT estado FROM VTA_02 WHERE [NUM_INT] = " & nir
             Set rs = New ADODB.Recordset
             rs.Open q, cn1, adOpenDynamic, adLockOptimistic
             If Not rs.EOF And Not rs.BOF Then
                rs("estado") = "F"
                rs.Update
             End If
          End If
        End If
        i = i + 1
     Wend
          
          
          
          
End Sub



Function verificaremito(ByVal n As Long) As Integer
q = "select id_producto, cantidad from vta_03 where [num_int] = " & n
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
     renglon = Val(msf1.TextMatrix(i, 0))
     If renglon > 0 Then
      r = Val(msf1.TextMatrix(i, 7))
      s = s + r
      'v = v + (r * Val(msf1.TextMatrix(i, 6)) / 100)
         
      
      'agrega en composicion de iva
      X = 1
      While X < vta_facturacion2.msf1.Rows
        If Val(vta_facturacion2.msf1.TextMatrix(X, 0)) = Val(msf1.TextMatrix(i, 6)) Then
           vta_facturacion2.msf1.TextMatrix(X, 1) = Val(vta_facturacion2.msf1.TextMatrix(X, 1)) + r
           vta_facturacion2.msf1.TextMatrix(X, 2) = Val(vta_facturacion2.msf1.TextMatrix(X, 2)) + (r * Val(msf1.TextMatrix(i, 6)) / 100)
           X = vta_facturacion2.msf1.Rows
        Else
           X = X + 1
        End If
      Wend
     End If
      
  
  Next i
  'vta_facturacion2.sacatotales
  't_subtotal = vta_facturacion2.msf1.TextMatrix(9, 1)
  't_iva = vta_facturacion2.msf1.TextMatrix(9, 2)
  Call sacatotales
  Call sacaperc
  Call sacatotales
 Else
  
 If t_letra = "B" Then
  s = 0
  v = 0
  t = 0
  For i = 1 To msf1.Rows - 1
     renglon = Val(msf1.TextMatrix(i, 0))
     If renglon > 0 Then
      
      r = Val(msf1.TextMatrix(i, 7))
      R2 = Val(msf1.TextMatrix(i, 8))
      's = s + r
      't = t + (R2 * Val(msf1.TextMatrix(i, 3)))
      't = t + (r * Val(msf1.TextMatrix(i, 6)) / 100)
            'agrega en composicion de iva
      X = 1
      While X < vta_facturacion2.msf1.Rows
        If Val(vta_facturacion2.msf1.TextMatrix(X, 0)) = Val(msf1.TextMatrix(i, 6)) Then
           vta_facturacion2.msf1.TextMatrix(X, 1) = Val(vta_facturacion2.msf1.TextMatrix(X, 1)) + r
           vta_facturacion2.msf1.TextMatrix(X, 2) = Val(vta_facturacion2.msf1.TextMatrix(X, 2)) + (r * Val(msf1.TextMatrix(i, 6)) / 100)
           X = vta_facturacion2.msf1.Rows
        Else
           X = X + 1
        End If
      Wend
    End If
    
  Next i
  
Else
   'factura c
   s = 0
  v = 0
  t = 0
  For i = 1 To msf1.Rows - 1
     renglon = Val(msf1.TextMatrix(i, 0))
     If renglon > 0 Then
      
      r = Val(msf1.TextMatrix(i, 7))
      R2 = Val(msf1.TextMatrix(i, 8))
      's = s + r
      't = t + (R2 * Val(msf1.TextMatrix(i, 3)))
      't = t + (r * Val(msf1.TextMatrix(i, 6)) / 100)
            'agrega en composicion de iva
      X = 1
      While X < vta_facturacion2.msf1.Rows
        If Val(vta_facturacion2.msf1.TextMatrix(X, 0)) = Val(msf1.TextMatrix(i, 6)) Then
           vta_facturacion2.msf1.TextMatrix(X, 1) = Val(vta_facturacion2.msf1.TextMatrix(X, 1)) + r
           vta_facturacion2.msf1.TextMatrix(X, 2) = 0
           X = vta_facturacion2.msf1.Rows
        Else
           X = X + 1
        End If
      Wend
    End If
    
  Next i

 End If
End If
  
d = 0
X = 1
While X < vta_facturacion2.msf1.Rows
  d1 = 0
  If Val(t_porcdescuento) > 0 Then
     d1 = Format(Val(vta_facturacion2.msf1.TextMatrix(X, 1)) * Val(t_porcdescuento) / 100, "#####0.00")
     d = d + d1
  End If
  vta_facturacion2.msf1.TextMatrix(X, 1) = Format$(Val(vta_facturacion2.msf1.TextMatrix(X, 1)) - d1, "#####0.00")
  If t_letra <> "C" Then
    vta_facturacion2.msf1.TextMatrix(X, 2) = Format$(Val(vta_facturacion2.msf1.TextMatrix(X, 1)) * Val(vta_facturacion2.msf1.TextMatrix(X, 0)) / 100, "#####0.00")
  Else
    vta_facturacion2.msf1.TextMatrix(X, 2) = 0
  End If
    
  X = X + 1
Wend
 
  
  't_subtotal = s
  't_iva = t - s
  
  
  t_descuento = Format$(d, "#######.00")
  
  
  Call sacatotales
  Call sacaperc
  Call sacatotales
 

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
  
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter3(Me, 21)
End If


End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_clientes(c_prov)
c_prov.ListIndex = 0
c_tipoop.ListIndex = 0
c_transferencia.ListIndex = 0

Call carga_SUCURSALES(c_sucursal)



Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & glo.sucursal & " and  [id_tipocomp] <= 40 order by descripcion"
rs.Open q, cn1
Call llena_combo(rs, "descripcion", "id_tipocomp", c_tipocomp, True)
Set rs = Nothing




c_tipocomp.ListIndex = buscaindice(c_tipocomp, 1)

Frame14.Visible = False
Frame11.Visible = True

Set rs = New ADODB.Recordset
q = "select * from vta_05 order by [denominacion]"
rs.Open q, cn1
Call llena_combo(rs, "denominacion", "id_vendedor", c_vend, True)
Set rs = Nothing
c_vend.ListIndex = 0
Call armagrid
Call barraesag(Me)
Option1 = True
If para.moneda = "P" Then
  Option4 = True
Else
  Option3 = True
End If
t_sucursal = Format$(para.punto_venta_usuario, "0000")
Load vta_facturacion1
Load vta_selremitos
Load vta_selcomp
Load vta_facturacion2
Load ABM_COMP_COMPRA2

Frame11.Visible = False

Call carga_actividades(c_actividad)

Load vta_clientes
Load vta_formapago
vta_clientes.limpia
gcuit = "0"

Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)
If cl_fiscal.idmodelo <> 24 Then
  epson1.PortNumber = cl_fiscal.puerto
Else
  Set Fiscaltf = New Driver
  Fiscaltf.Modelo = cMODELO
  Fiscaltf.puerto = cPUERTO
  Fiscaltf.baudios = cBAUDIOS
End If
Set cl_fiscal = Nothing


Set rs = New ADODB.Recordset
q = "select [tipo_control_limite_credito] from g0 where [sucursal] = 0"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  t_cl = rs("tipo_control_limite_credito")
Else
  t_cl = False
End If
Set rs = Nothing


c_sucursal.ListIndex = buscaindice(c_sucursal, para.punto_venta_usuario)

c_transferencia.Visible = False
Label15.Visible = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload vta_facturacion1
Unload vta_facturacion2
   Unload vta_selremitos
    Unload vta_clientes
    Unload vta_formapago
    Unload vta_selcomp
    Unload ABM_COMP_COMPRA2
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F3] Descipcion extra - [F5] Saca Renglon - [F7] Costo - [F9] Graba "
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
   gen_descextra.t_modulo = "F"
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
 Call CALCULATOTALES
End If

If KeyCode = vbKeyF7 Then
 If msf1.Row > 0 Then
   g = InputBox$("Costo Total s/ Iva", "Costo Item", msf1.TextMatrix(msf1.Row, 10))
   If g <> "" Then
       msf1.TextMatrix(msf1.Row, 10) = Format$(Val(g), "######0.00")
   End If
 End If
 Call CALCULATOTALES
End If

If KeyCode = vbKeyF9 Then
  Call CALCULATOTALES
  Call sacatotales
  Call sacaperc
  Call sacatotales
  Call renumera
  Frame2.Enabled = True
  btnacepta.Enabled = True
  c_vend.SetFocus
End If

If KeyCode = vbKeyInsert Then
   vta_facturacion1.t_renglon = ""
   vta_facturacion1.t_cantidad = ""
   vta_facturacion1.t_pu = ""
   vta_facturacion1.t_importe = ""
   If msf1.Rows - 1 < cantlineas Then
     vta_facturacion1.Show
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
  Call bloquea_comp
  numint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(t_sucursal)
  cl_compvta.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
  cl_compvta.letra = t_letra
  cl_compvta.numcomp = Val(t_numcomp)
  abreviatura = cl_compvta.abreviatura
  
  If ch_plan.Value = 0 Then
      ubicacionctacte = cl_compvta.ctacte
  Else
      ubicacionctacte = "N"
  End If
     
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
      
      If EXISTE = "N" Then
        cl_compvta.ACTUALIZA_NUMERADOR
      End If
      
      If Option4 = True Then
        moneda = "P"
      Else
        moneda = "D"
      End If
      
      
      Set rs = New ADODB.Recordset
      q = "select id_actividad, alicuota_ib, cuenta_contable_venta from g8 where [id_actividad] = " & c_actividad.ItemData(c_actividad.ListIndex)
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
      
        
      'Set cl_cli = New Clientes
      'cl_cli.carga (c_prov.ItemData(c_prov.ListIndex))
              
      tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
       
      If c_prov.ListIndex = 0 Then
        idcli = 1
      Else
        idcli = c_prov.ItemData(c_prov.ListIndex)
      End If
      
      
      If Check3 Then
        T2 = 0
      Else
        T2 = Val(T_total2)
      End If
      
      
        'COMPROBATES ASOCIADOS A nc LO GRABO EN CHOFER2
       compasocnc = ""
       If c_tipocomp.ItemData(c_tipocomp.ListIndex) >= 2 And c_tipocomp.ItemData(c_tipocomp.ListIndex) <= 3 Then  ' solo nc/nd
          F = vta_selcomp.msf1.Rows - 1
          
          For i = 1 To F
            If vta_selcomp.msf1.TextMatrix(i, 0) = "**" Then
               compasocnc = vta_selcomp.msf1.TextMatrix(i, 1) & "-" & vta_selcomp.msf1.TextMatrix(i, 2)
               i = F
            End If
          Next i
      End If
      
      
      
      cn1.BeginTrans
       
       QUERY = ""
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], " & _
" [dominio_acoplado02], [SALDO_IMPAGO02], [num_z], [cae], [cae_vence], [tipo_op], [descuento],[numint_asociado])"



QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & _
", " & idcli & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotal) & ", " & Val(t_nograbado) & ", " & Val(t_iva) & ", " & Val(t_total) & _
", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & ubicacionctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & t_observaciones & _
" ', " & Val(t_cotizacion) & ", " & T2 & ", '" & moneda & "', " & c_vend.ItemData(c_vend.ListIndex) & ", '" & cl_compvta.venta & "', '" & contado & "', " & Val(t_perc)

QUERY2 = ", 0, " & Val(t_perciva) & ", " & codact & ", " & Val(t_alicuotaib) & ", " & Val(t_alicuotaperciva) & ", " & Check1 & ", '" & t_fechavto & "', 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & _
", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & Left$(vta_clientes.t_localidad, 50) & "', " & tiporespiva & ", '" & compasocnc & "', ' ', ' ', " & ssi & ", " & para.z_actual & ", '" & t_cae & "', '" & Format(t_cae_vence, "@@@@/@@/@@") & "', " & c_tipoop.ListIndex + 1 & ", " & Val(t_descuento) & ",0)"

                                                                                                                                                                                                                                                            
cn1.Execute QUERY & QUERY2
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
      
        
        If cl_compvta.STOCK <> "N" Then
           QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo], [id_cliente])"
           QUERY = QUERY & " VALUES ('" & t_fecha & "', " & Val(msf1.TextMatrix(i, 1)) & ", " & msf1.TextMatrix(i, 3) & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.abreviatura & t_letra & Format$(t_sucursal, "0000") & "-" & Format$(t_numcomp, "00000000") & "', '" & Left$(c_prov, 50) & _
           "', " & numint & ",'V', " & idcli & ")"
           cn1.Execute QUERY
          
           If cl_compvta.STOCK = "E" Then
             c = Val(msf1.TextMatrix(i, 3))
             COSTOINV = COSTOINV + (costo * Val(msf1.TextMatrix(i, 3)))
           Else
             c = -Val(msf1.TextMatrix(i, 3))
             COSTOINV = COSTOINV - (costo * Val(msf1.TextMatrix(i, 3)))
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

  Else
    'grabo desc extra
    QUERY = "INSERT INTO vta_015([num_int], [RENGLON], [desc_ext], [cant_lineas])"
    QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i - 1, 0)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 3)) & ")"
    cn1.Execute QUERY
  End If


Next i
      
      
      'actualizo tasa de iva
      If cl_compvta.grabado <> "N" Then
       If verificatasaunica Then
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
          QUERY = QUERY & " VALUES (" & numint & ", " & Val(vta_facturacion.msf1.TextMatrix(1, 6)) & ", " & Val(t_iva) & ", " & Val(t_subtotal) & ", " & tiporespiva & ", " & cuentaact & ")"
          cn1.Execute QUERY
       Else
        For i = 1 To 7
        If Val(vta_facturacion2.msf1.TextMatrix(i, 1)) > 0 Then
          QUERY = "INSERT INTO vta_09([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta09])"
          QUERY = QUERY & " VALUES (" & numint & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 0)) & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 2)) & ", " & Val(vta_facturacion2.msf1.TextMatrix(i, 1)) & ", " & tiporespiva & ", " & cuentaact & ")"
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
        'graba fortma de pago
        Call grabaformapago
        
     End If
      

    'contabilidad
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
         If Option3 = True Then
           'fact en dolares
           tot = Val(T_total2)
           m = Val(t_cotizacion)
         Else
           tot = Val(t_total)
           m = 1
         End If
         
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Ventas] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'V', " & numint & ", " & tot & ", " & tot & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_prov), 50) & "')"
         cn1.Execute QUERY
      
         
         
          If Option1 = True Then
           'ingresa deudores
           cta = para.cuenta_deudores
           ic = 1
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & tot & ", '" & dcta & "')"
           cn1.Execute QUERY
           ic = ic + 1
          Else
           'ingresa forma de pago
            ic = 1
            For i = 1 To vta_formapago.msf2.Rows - 1
               cta = Val(vta_formapago.msf2.TextMatrix(i, 9))
               im = Format(Val(vta_formapago.msf2.TextMatrix(i, 6)) * m, "######0.00")
               dcta = vta_formapago.msf2.TextMatrix(i, 3)
               QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
               QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & im & ", '" & dcta & "')"
               'MsgBox (QUERY)
               cn1.Execute QUERY
               ic = ic + 1
            Next i
         End If
         

         
         If Val(t_nograbado) > 0 Then
           'cuenta nogbra
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_conceptos_nograbados & ", '" & u2 & "', " & Format(Val(t_nograbado) * m, "#####0.00") & ", 'No Grabado')"
           cn1.Execute QUERY
           ic = ic + 1
           
         End If
                   
         If Val(t_perc) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_perc_IB & ", '" & u2 & "', " & Format(Val(t_perc) * m, "####0.00") & ", 'Perc. IB')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
          
          If Val(t_perciva) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_perc_iva & ", '" & u2 & "', " & Format(Val(t_perciva) * m, "####0.00") & ", 'Perc. IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         If Val(t_iva) > 0 And cl_compvta.grabado <> "N" Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_ventas & ", '" & u2 & "', " & Format(Val(t_iva) * m, "#####0.00") & ", 'IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         'contrapartida
         
         If cl_compvta.grabado <> "N" Then
           importe = Val(t_subtotal) * m
         Else
           importe = Val(t_total) * m
         End If
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cuentaact & ", '" & u2 & "', " & Format(importe, "######0.00") & ", '" & "Ventas" & "')"
         cn1.Execute QUERY
         ic = ic + 1
      
      
      End If
      
      
      If COSTOINV <> 0 Then
         If COSTOINV > 0 Then
           u1 = "H"
           u2 = "D"
         Else
           u2 = "H"
           u1 = "D"
           COSTOINV = -COSTOINV
         End If
         tot = COSTOINV
         If cl_compvta.contabilidad = "N" Then
          'realizo asiento de costo aunque el doc. no mueva contabilidad
          numintcgr = saca_ultnumero_int_comp("G")

                 
          QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
          QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Ventas] " & cl_compvta.abreviatura & " " & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 'V', " & numint & ", " & tot & ", " & tot & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_prov), 50) & "')"
          cn1.Execute QUERY
        
         End If
         
         ic = 1
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_inventario & ", '" & u2 & "', " & Format(COSTOINV * m, "#####0.00") & ", 'Inventario')"
         cn1.Execute QUERY
         ic = ic + 1
                           
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_costo & ", '" & u1 & "', " & Format(COSTOINV * m, "######0.00") & ", '" & "Costo Merc." & "')"
         cn1.Execute QUERY
      End If
      
     End If
     
      Set rs = Nothing
      Set cl_compvta = Nothing
      Set cl_cli = Nothing

      
      
      
      'actualizo remitos
     'If Val(vta_selremitos.t_seleccionados) > 0 Then
        For i = 1 To vta_selremitos.msf1.Rows - 1
          If vta_selremitos.msf1.TextMatrix(i, 0) = "**" Then
             QUERY = "INSERT INTO vta_08([id_factura], [id_remito])"
             QUERY = QUERY & " VALUES (" & numint & ", " & Val(vta_selremitos.msf1.TextMatrix(i, 4)) & ")"
             cn1.Execute QUERY
          End If
        Next i
     
      
      
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Emitir Factura/NC/ND NI:" & numint & "', " & para.id_usuario & ", 'V', " & numint & ", '" & Now & "', '[" & c_tipocomp.ItemData(c_tipocomp.ListIndex) & "] " & t_letra & " " & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 12, " & idcli & ")"
  
     cn1.Execute QUERY

      
      cn1.CommitTrans
      
      
     
        Call actualizaremitos2
     'End If
      
      
      If glo.sucursalf <> Val(t_sucursal) Then
          J = MsgBox("Confirma Impresion del Comprobante", 4)
          If J = 6 Then
             Set cl_compvta = New comprobantes_venta
             cl_compvta.cargar2 (numint)
             cl_compvta.imprimir
          End If
      End If
      
      
      
      If ch_plan.Value = 1 Then
      
            'arma plan de pago
            Call generaplancuotas(numint)
        
      
      End If
           
      
      Call INICIALIZA2(Me)
      Call armagrid
      
      
      
      'Set cl_compvta = New comprobantes_venta
      'cl_compvta.web (numint)
      
      
      Call libera_comp
      
      
      
      
      
      
      c_prov.SetFocus
      Frame2.Enabled = False
      t_sucursal = Format$(c_sucursal, "0000")
      vta_formapago.armagrid2
      Frame11.Visible = False
      vta_selremitos.armagrid
      
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos y vuelva a repetir la operacion")
  


  


End Sub

Sub generaplancuotas(ni21)
  
For i = 0 To Val(t_cantcuotas) - 1
  nint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(t_sucursal)
  cl_compvta.actual (251) 'cuotas
  cl_compvta.letra = t_letra
  numcomp = Val(RTrim$(t_numcomp) + Format$(i, "00"))
  cl_compvta.numcomp = numcomp  'numero nota venta + numero cuota
  abreviatura = cl_compvta.abreviatura
  ubicacionctacte = cl_compvta.ctacte
  ep = "N"
  cp = "0000-00000000"
  contado = "N"
  If Option4 = True Then
      ssi = Val(t_total)
      moneda = "P"
      cm = Val(t_valorcuota)
      cotram = Val(t_valorcuota) / Val(t_cotizacion)
   Else
      ssi = Val(T_total2)
      moneda = "D"
      cm = Val(t_valorcuota)
      cotram = Val(t_valorcuota) * Val(t_cotizacion)
 
  End If
  
  
  
  
  codact = 0
  alicuotaib = 0
  cuentaact = para.cuenta_ventas
  tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
  idcli = c_prov.ItemData(c_prov.ListIndex)
     
     
  fecha = DateValue(t_fechacuota1) + (i * 30)
  observaciones = Format$(c_tipocomp.ItemData(c_tipocomp.ListIndex), "000-") & Format$(t_sucursal, "0000") & "-" & Format$(t_numcomp, "00000000") & "/" & Format$(i, "00")
  
  
      If Check3 Then
        T2 = 0
      Else
        T2 = cotram
      End If
  
  
  
  
  cn1.BeginTrans
       
       QUERY = ""
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], " & _
" [dominio_acoplado02], [SALDO_IMPAGO02], [num_z], [cae], [cae_vence], [tipo_op], [descuento], [numint_asociado])"



QUERY = QUERY & " VALUES (" & nint & ", " & Val(t_sucursal) & ", " & numcomp & ", '" & t_letra & "', 251 " & _
", " & idcli & ", '" & fecha & "', " & para.id_usuario & ", " & cm & ", 0, 0," & cm & _
", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & observaciones & _
" ', " & Val(t_cotizacion) & ", " & T2 & ", '" & moneda & "', " & c_vend.ItemData(c_vend.ListIndex) & ", '" & cl_compvta.venta & "', '" & contado & "', " & Val(0)

QUERY2 = ", 0, " & Val(0) & ", " & codact & ", " & Val(0) & ", " & Val(0) & ", " & Check1 & ", '" & fecha & "', 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & _
", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & Left$(vta_clientes.t_localidad, 50) & "', " & tiporespiva & ", ' ', ' ', ' ', " & cm & ", " & para.z_actual & ", ' ', '" & fecha & "', " & c_tipoop.ListIndex + 1 & ", " & Val(0) & ", " & ni21 & ")"

'MsgBox (QUERY & QUERY2)
                                                                                                                                                                                                                                                            
cn1.Execute QUERY & QUERY2

cn1.CommitTrans

Set cl_cli = Nothing


Next i


Call generapagare(ni21)

End Sub

Sub generapagare(ni21)
  'cada plan de cuotas tiene asociado un pagare
  
  'vta_02
    'descuento --> tiene interes mensual
    'total_bultos --> cantidad cuotas
    'valor_declarado --> importe por cuota
   
  nint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(t_sucursal)
  cl_compvta.actual (401) 'pagare
  cl_compvta.letra = t_letra
  cl_compvta.SACANUMCOMP
  numcomp = cl_compvta.numcomp
  abreviatura = cl_compvta.abreviatura
  ubicacionctacte = cl_compvta.ctacte
  ep = "P"
  cp = "0000-00000000"
  contado = "S"
  If Option4 = True Then
      ssi = Val(t_total)
      moneda = "P"
      cotram = Val(T_total2)
    Else
      ssi = Val(T_total2)
      moneda = "D"
      cotram = Val(t_total)
    
  End If
  
  codact = 0
  alicuotaib = 0
  cuentaact = para.cuenta_ventas
  tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
  idcli = c_prov.ItemData(c_prov.ListIndex)
  fecha = DateValue(t_fecha)
  observaciones = "Pagare " & Format$(c_tipocomp.ItemData(c_tipocomp.ListIndex), "000-") & Format$(t_sucursal, "0000") & "-" & Format$(t_numcomp, "00000000")
  
      If Check3 Then
        T2 = 0
      Else
        T2 = cotram
      End If
  
  cm = ssi
  
  
  cn1.BeginTrans
       
       QUERY = ""
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], " & _
" [dominio_acoplado02], [SALDO_IMPAGO02], [num_z], [cae], [cae_vence], [tipo_op], [descuento], [numint_asociado])"



QUERY = QUERY & " VALUES (" & nint & ", " & Val(t_sucursal) & ", " & numcomp & ", '" & t_letra & "', 401 " & _
", " & idcli & ", '" & fecha & "', " & para.id_usuario & ", " & cm & ", 0, 0," & cm & _
", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & observaciones & _
" ', " & Val(t_cotizacion) & ", " & T2 & ", '" & moneda & "', " & c_vend.ItemData(c_vend.ListIndex) & ", '" & cl_compvta.venta & "', '" & contado & "', " & Val(0)

QUERY2 = ", 0, " & Val(0) & ", " & codact & ", " & Val(0) & ", " & Val(0) & ", " & Check1 & ", '" & fecha & "'," & Val(t_cantcuotas) & "," & Val(t_valorcuota) & ", ' ', ' ', ' ', 0, " & Val(c_sucursal) & _
", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & Left$(vta_clientes.t_localidad, 50) & "', " & tiporespiva & ", ' ', ' ', ' ', " & cm & ", " & para.z_actual & ", ' ', '" & fecha & "', " & c_tipoop.ListIndex + 1 & ", " & Val(t_interes_cuota) & ", " & ni21 & ")"

'MsgBox (QUERY2)
                                                                                                                                                                                                                                                            
cn1.Execute QUERY & QUERY2

cn1.CommitTrans







Set cl_cli = Nothing



J = MsgBox("Los comprobantes con Plan de Cuotas generan automáticamente un Pagaré respaldatorio del Plan de Pago, ¿desea imprimirlo?", 4)
If J = 6 Then
  'imprime pagare
  
  Set cl_compvta = New comprobantes_venta
  cl_compvta.cargar2 (nint)
  If cl_compvta.numint > 0 Then
     cl_compvta.imprimir
  End If
  Set cl_compvta = Nothing
End If


End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
   If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
     vta_facturacion1.t_renglon = msf1.Row
     vta_facturacion1.t_basico = msf1.TextMatrix(msf1.Row, 1)
     vta_facturacion1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
     vta_facturacion1.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
     vta_facturacion1.t_unidad = msf1.TextMatrix(msf1.Row, 4)
     vta_facturacion1.t_pu = msf1.TextMatrix(msf1.Row, 5)
     vta_facturacion1.t_importe = msf1.TextMatrix(msf1.Row, 7)
     vta_facturacion1.Show
   Else
     Load gen_descextra
     gen_descextra.Text1 = msf1.TextMatrix(msf1.Row, 2)
     gen_descextra.t_modulo = "F"
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
Call CALCULATOTALES
End Sub

Private Sub Option1_Click()
Command8.Enabled = False

Me.BackColor = &HE0E0E0

End Sub

Private Sub Option1_GotFocus()
Call keyform(Me, "D")

End Sub

Private Sub Option1_LostFocus()
Call keyform(Me, "A")

End Sub

Private Sub Option2_Click()
Command8.Enabled = True
Me.BackColor = &HFFC0C0
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

Private Sub t_alicuotaib_LostFocus()
If Val(t_alicuotaib) < 0 Then
  t_alicuotaib = "0.00"
End If
Call sacaperc
End Sub

Private Sub t_cantcuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call calculacuotas
  t_valorcuota.SetFocus
End If

End Sub

Sub calculacuotas()
If IsNull(t_cantcuotas) Then
  t_cantcvuotas = 1
End If

If Val(t_cantcuotas) < 1 Then
    t_cantcuotas = 1
End If
  
If Val(t_cantcuotas) > Val(t_cuotas_sininteres) Then
  'aplica interes
  interes = Format((Val(t_cantcuotas) * Val(t_interes_cuota) * Val(t_total)) / 100, "#######0.00")
Else
  interes = 0
End If
  
t_valorcuota = Format$((Val(t_total) + interes) / Val(t_cantcuotas), "######0.00")

End Sub

Private Sub t_cantcuotas_LostFocus()
If Val(t_cantcuotas) < 1 Then
  t_cantcuotas = 1
End If

Call calculacuotas

End Sub

Private Sub t_cotizacion_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub t_cotizacion_LostFocus()
If Val(t_cotizacion) <= 0 Then
   t_cotizacion = 1
End If
Call mensaje
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
c_actividad.ListIndex = buscaindice(c_actividad, sacaactividadsucursal(Val(t_sucursal)))
Call iniciacomp

'Call verifica_fechacorte(t_fecha)
End Sub


Private Sub t_fechacuota1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_cantcuotas.SetFocus
End If

End Sub

Private Sub t_fechacuota1_LostFocus()
If Not IsDate(t_fechacuota1) Then
  t_fechacuota1 = Format$(Now, "dd/mm/yyyy")
Else
  t_fechacuota1 = Format$(t_fechacuota1, "dd/mm/yyyy")
End If

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


If KeyAscii = 13 Then
  Call solonum(KeyAscii, 0)
  If c_tipocomp.ItemData(c_tipocomp.ListIndex) < 30 Or c_tipocomp.ItemData(c_tipocomp.ListIndex) > 32 Then
    t_fecha.SetFocus
  End If
End If
End Sub

Private Sub t_numcomp_LostFocus()
If IsNumeric(t_numcomp) Then
   t_numcomp = Format$(t_numcomp, "00000000")
   
    Call carga3
   'EXISTE = "N"
   
   
Else
  t_numcomp.SetFocus
End If
End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub

Private Sub t_perc_LostFocus()
Call sacatotales

End Sub

Private Sub t_porcdescuento_LostFocus()
  If Val(t_porcdescuento) > 0 Then
    Set rs = New ADODB.Recordset
    q = "select * from g0 where [sucursal] = 0"
    rs.Open q, cn1
    d1 = rs("descuento1")
    Set rs = Nothing
      
    If para.id_grupo_modulo_actual <= 7 Then
       If Val(t_porcdescuento) > d1 Then
           MsgBox ("El máximo descuento permitido es " & d1 & "%")
           t_porcdescuento = 0
       End If
    End If
  End If
  Call CALCULATOTALES
  Call sacatotales
  Call sacaperc
  Call sacatotales
End Sub

Private Sub t_subtotal_LostFocus()
Call sacatotales
End Sub
Sub sacatotales()
'busco iva
 
     X = 1
     iva = 0
     neto = 0
    While X < vta_facturacion2.msf1.Rows
      iva = iva + Val(vta_facturacion2.msf1.TextMatrix(X, 2))
      neto = neto + Val(vta_facturacion2.msf1.TextMatrix(X, 1))
      X = X + 1
    Wend


'busco perc
If ABM_COMP_COMPRA2.msf1.Rows > 1 Then
  t = 0
  For i = 1 To ABM_COMP_COMPRA2.msf1.Rows - 1
    t = t + Val(ABM_COMP_COMPRA2.msf1.TextMatrix(i, 3))
  Next i
  t_perc = Format$(t, "######0.00")
Else
  t_perc = Format$(0, "######0.00")
End If


t_descuento = Format$(Val(t_descuento), "######0.00")
t_subtotal2 = Format$(neto + Val(t_descuento), "######0.00")
t_nograbado = Format$(Val(t_nograbado), "######0.00")
't_perc = Format$(Val(t_perc), "######0.00")
t_subtotal = Format$(neto, "######0.00")
t_iva = Format$(iva, "######0.00")
t_perciva = Format$(Val(t_perciva), "######0.00")
t_total = Format$(Val(t_subtotal) + Val(t_nograbado) + Val(t_perc) + Val(t_iva) + Val(t_perciva), "######0.00")
If Option4 = True Then
 If Val(t_cotizacion) < 1 Then
   t_cotizacion = 1
 End If
 T_total2 = Format$(Val(t_total) / Val(t_cotizacion), "#####0.00")
Else
  T_total2 = Format$(Val(t_total) * Val(t_cotizacion), "#####0.00")
End If

If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 36 Then
   t_valorcuota = Format$(Val(t_total) / Val(t_cantcuotas), "######0.00")
End If

End Sub
Sub sacaperc()
If Option3 = True Then
   s = Val(t_subtotal) * Val(t_cotizacion)
 Else
   s = Val(t_subtotal)
 End If

q = "select * from i_01 where [id_impuesto] = 1"
'MsgBox (q)
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
   q = "select alicuota_i, alicuota_n, importe_minimo_sujeto_ret, retencion-minima from  i_01, i_02 where i_01.[id_impuesto] = i_02.[id_impuesto] and i_01.[id_impuesto] = 2"
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
'espere.Show
'espere.Label1 = "Inicializando Comprobante....."
'espere.Refresh
Call inicia
'Unload espere


End Sub

Private Sub t_total_LostFocus()
t_total = Format$(t_total, "######0.00")
End Sub

Private Sub T_total2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Option2 = True Then
      Command8.Enabled = True
   Else
      Command8.Enabled = False
 End If
 btnacepta.Enabled = True
 btnacepta.SetFocus
End If

End Sub



Private Sub t_valorcuota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  t_cantcuotas.SetFocus
End If
End Sub

Private Sub t_valorcuota_LostFocus()
If Val(t_valorcuota) <= 0 Then
   MsgBox ("Vertifique el importe de la cuota")
End If
End Sub

Private Sub UpDown1_DownClick()
If Val(t_cantcuotas) > 1 Then
  t_cantcuotas = Val(t_cantcuotas) - 1
  Call calculacuotas
End If
End Sub

Private Sub UpDown1_UpClick()
  t_cantcuotas = Val(t_cantcuotas) + 1
  Call calculacuotas
End Sub
