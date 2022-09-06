VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_listaprecios2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTA DE PRECIOS"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8895
   ScaleWidth      =   12120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consultas"
      Height          =   735
      Left            =   240
      TabIndex        =   83
      Top             =   7440
      Width           =   4095
      Begin VB.CommandButton Command8 
         Caption         =   "Ultimas otras Compras y Ventas"
         Height          =   375
         Left            =   1440
         TabIndex        =   85
         ToolTipText     =   "Calcula los nuevois valores segun la estructura de costo donde: costoreal= (precio compra - dto compra + flete compra)"
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Stock"
         Height          =   375
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox t_linea 
      Height          =   375
      Left            =   7200
      TabIndex        =   82
      Text            =   "Text1"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contadores"
      Height          =   1455
      Left            =   6720
      TabIndex        =   63
      Top             =   1080
      Width           =   1815
      Begin VB.TextBox t_pedidos 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   68
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox t_oc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   66
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox t_stock 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   64
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label29 
         BackColor       =   &H00800080&
         Caption         =   "Pedidos Int. Pendientes"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label28 
         BackColor       =   &H00800080&
         Caption         =   "en O.C."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label27 
         BackColor       =   &H00800080&
         Caption         =   "Stock "
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ult. Actualizacion"
      Height          =   1815
      Left            =   9960
      TabIndex        =   60
      Top             =   5520
      Width           =   1695
      Begin VB.TextBox t_fechaactuc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   49
         TabIndex        =   73
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox t_fechaactu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   49
         TabIndex        =   61
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Precio Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Precio Venta"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Ultima Operacion"
      Height          =   1095
      Left            =   240
      TabIndex        =   54
      Top             =   6240
      Width           =   9375
      Begin VB.TextBox t_cotizultcom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8280
         MaxLength       =   70
         TabIndex        =   103
         ToolTipText     =   "Si el comprobante es de tipo ""A"" el precio de venta es ""sin iva"", sino es ""precio final"""
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox t_ultimacompra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   70
         TabIndex        =   80
         ToolTipText     =   "Si el comprobante es de tipo ""A"" el precio de venta es ""sin iva"", sino es ""precio final"""
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox t_ultvta 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   70
         TabIndex        =   55
         ToolTipText     =   "Si el comprobante es de tipo ""A"" el precio de venta es ""sin iva"", sino es ""precio final"""
         Top             =   240
         Width           =   7815
      End
      Begin VB.Label Label39 
         BackColor       =   &H00800080&
         Caption         =   "Cotiz U$s Ult.Com"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   102
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackColor       =   &H00800080&
         Caption         =   "Ultima Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackColor       =   &H00800080&
         Caption         =   "Ultima Venta"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   240
      TabIndex        =   51
      Top             =   5520
      Width           =   7815
      Begin VB.TextBox t_observaciones 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   49
         TabIndex        =   52
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label23 
         BackColor       =   &H00800080&
         Caption         =   "Observaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Varios"
      Height          =   2055
      Left            =   8640
      TabIndex        =   40
      Top             =   1080
      Width           =   3255
      Begin VB.TextBox t_moneda 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   49
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox t_tipo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   47
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox t_envase 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox t_stockminimo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   43
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox c_unidad 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   41
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[Pesos] - [Dolares]"
         Height          =   255
         Left            =   1560
         TabIndex        =   72
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[Prod] - [Mat. Prima]"
         Height          =   255
         Left            =   1560
         TabIndex        =   71
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00800080&
         Caption         =   "Moneda"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label21 
         BackColor       =   &H00800080&
         Caption         =   "Tipo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label20 
         BackColor       =   &H00800080&
         Caption         =   "Envase"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label19 
         BackColor       =   &H00800080&
         Caption         =   "Stock Min."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label18 
         BackColor       =   &H00800080&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   11655
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Redondeo p/ precio venta final"
         Height          =   615
         Left            =   480
         TabIndex        =   96
         Top             =   1560
         Width           =   8655
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "10/100  Ej. 8.20"
            Height          =   255
            Left            =   3360
            TabIndex        =   101
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "25/100  Ej. 8.25"
            Height          =   255
            Left            =   5160
            TabIndex        =   100
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "50/100  Ej. 8.50"
            Height          =   255
            Left            =   6840
            TabIndex        =   99
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Entero Ej. 8.00 "
            Height          =   255
            Left            =   1680
            TabIndex        =   98
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sin redondeo"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox t_dtocompra2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   94
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_sugerido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   93
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Precio Final Sugerido"
         Height          =   255
         Left            =   9360
         TabIndex        =   92
         ToolTipText     =   "Sugiere un precio de Venta final segun la estructura de costo"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.ComboBox c_tasaib 
         Height          =   315
         Left            =   7440
         TabIndex        =   88
         Text            =   "Combo1"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Calculo de P&F a PU"
         Height          =   375
         Left            =   9360
         TabIndex        =   75
         ToolTipText     =   "Calcula los nuevois valores segun la estructura de costo donde: costoreal= (precio compra - dto compra + flete compra)"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Calculo de P&U a PF"
         Height          =   375
         Left            =   4440
         TabIndex        =   70
         ToolTipText     =   "Calcula los nuevois valores segun la estructura de costo donde: costoreal= (precio compra - dto compra + flete compra)"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Calcula Precio Unitario"
         Height          =   375
         Left            =   2640
         TabIndex        =   59
         ToolTipText     =   "Calcula los nuevois valores segun la estructura de costo donde: costoreal= (precio compra - dto compra + flete compra)"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Calcula Estructura de Costos"
         Height          =   375
         Left            =   480
         TabIndex        =   58
         ToolTipText     =   "Calcula los nuevois valores segun la estructura de costo donde: costoreal= (precio compra - dto compra + flete compra)"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox t_final 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   10320
         MaxLength       =   10
         TabIndex        =   39
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_impuesto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_tasaimpint 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   37
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox c_iva 
         Height          =   315
         Left            =   7440
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_pu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   35
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_utilidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   34
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_costo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_fletecompra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_dtocompra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   31
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_preciocompra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   405
         Left            =   480
         MaxLength       =   10
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "% Dto Compra 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   95
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label36 
         BackColor       =   &H00800000&
         Caption         =   "Tasa IB"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   89
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Precio Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10320
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Impuesto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9360
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00800000&
         Caption         =   "% Imp. Int."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8520
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00800000&
         Caption         =   "Tasa Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7440
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "P.U. s/Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00800000&
         Caption         =   "% Utilidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Costo Real"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "% Flete"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3360
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "% Dto Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Precio Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   9120
      TabIndex        =   17
      Top             =   7440
      Width           =   2655
      Begin VB.CommandButton Command4 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   1440
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Grabar(F9)"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   11655
      Begin VB.TextBox t_medida 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10440
         MaxLength       =   34
         TabIndex        =   108
         ToolTipText     =   "[M]manual - [A] Automatica"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox t_color 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8760
         MaxLength       =   24
         TabIndex        =   106
         ToolTipText     =   "[M]manual - [A] Automatica"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox t_talle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   104
         ToolTipText     =   "[M]manual - [A] Automatica"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox t_tipocarga 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   86
         ToolTipText     =   "[M]manual - [A] Automatica"
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   7680
         TabIndex        =   78
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox t_textocentral 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9360
         MaxLength       =   19
         TabIndex        =   76
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox t_basico 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox t_codbarra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_detalle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   11
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label42 
         BackColor       =   &H00800000&
         Caption         =   "Medida"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9840
         TabIndex        =   109
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label41 
         BackColor       =   &H00800000&
         Caption         =   "Color"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8280
         TabIndex        =   107
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label40 
         BackColor       =   &H00800000&
         Caption         =   "Talle"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6960
         TabIndex        =   105
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label35 
         BackColor       =   &H00800000&
         Caption         =   "Carga Tique"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   87
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label34 
         BackColor       =   &H00800000&
         Caption         =   "Vigente"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6960
         TabIndex        =   79
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label33 
         BackColor       =   &H00800000&
         Caption         =   "Abreviatura"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8280
         TabIndex        =   77
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Cod. Barra"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   6375
      Begin VB.TextBox t_idprodprov 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   90
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox c_prov 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.ComboBox c_marca 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   960
         Width           =   4575
      End
      Begin VB.ComboBox c_depto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   600
         Width           =   4575
      End
      Begin VB.ComboBox c_grupo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label37 
         BackColor       =   &H00800080&
         Caption         =   "Cod. Prod. Prov."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800080&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00800080&
         Caption         =   "Marca"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800080&
         Caption         =   "Departamento"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Grupo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8535
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   635
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
            TextSave        =   "04/09/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "07:30 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_listaprecios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub Command1_Click()
Call calcula
End Sub
Sub calcula()
d = Val(t_preciocompra) * Val(t_dtocompra) / 100
n = Val(t_preciocompra) - Val(d)
d2 = n * Val(t_dtocompra2) / 100
n = n - d2
F = n * Val(t_fletecompra) / 100
n2 = F + n
t_costo = Format$(n2, "#####0.000")




End Sub

Private Sub Command2_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual >= 7 Then
   Call graba
End If
End Sub
Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   
   'On Error GoTo ERRORGRABA
       
   QUERY = "update a2 set  [Descripcion]='" & t_detalle & "' , [id_proveedor]=" & c_prov.ItemData(c_prov.ListIndex) & " , [id_unidad]=" & c_unidad.ItemData(c_unidad.ListIndex) & _
   " , [envase]=" & Val(t_envase) & " , [id_grupo]=" & c_grupo.ItemData(c_grupo.ListIndex) & " , [precio_ult_compra]=" & Val(t_preciocompra) & " , [pu]=" & Val(t_pu) & _
   " , [cod_tasaiva]=" & c_iva.ItemData(c_iva.ListIndex) & " , [stock_minimo]= " & Val(t_stockminimo) & " , [id_marca]=" & c_marca.ItemData(c_marca.ListIndex) & " , [id_departamento]=" & _
   c_depto.ItemData(c_depto.ListIndex) & " , [porc_utilidad]=" & Val(t_utilidad) & " , [costoreal]=" & Val(t_costo) & " , [flete_compra]=" & Val(t_fletecompra) & " , [dto_compra]=" & _
   Val(t_dtocompra) & " , [cod_barra]='" & RTrim$(t_codbarra) & "' , [precio_final]=" & Val(t_final) & " , [tasa_imp_interno]=" & Val(t_tasaimpint) & " , [tipo_producto]='" & t_tipo & _
   "' , [moneda]='" & t_moneda & "' , [impuesto]=" & Val(t_impuesto) & " , [observaciones]='" & t_observaciones & "' , [vigente]= " & Check1 & " , [tipo_carga_tique]='" & t_tipocarga & _
   "' , [id_tasaib]=" & c_tasaib.ItemData(c_tasaib.ListIndex) & " , [id_prod_prov]='" & RTrim$(UCase(t_idprodprov)) & "' , [dto_compra2]=" & Val(t_dtocompra2) & ", [dolar_ult_compra]=" & Val(t_cotizultcom) & _
   " , [talle]= '" & t_talle & " ', [color]= '" & t_color & " ', [medida]= '" & t_medida & " '"
      
      
      
      If Val(t_preciocompra) <> Val(t_preciocompra.Tag) Or Val(t_costo) <> Val(t_costo.Tag) Or Val(t_cotizultcom) <> Val(t_cotizultcom.Tag) Then
          QUERY = QUERY & " , [fecha_ult_compra]='" & Format$(Now, "dd/mm/yyyy") & "' "
      End If
      
      If Val(t_final) <> Val(t_final.Tag) Then
          QUERY = QUERY & " , [fecha_actu_precio_venta]='" & Format$(Now, "dd/mm/yyyy") & "' "
      End If
          
     QUERY = QUERY & " , [stock]=" & Val(t_stock) & " , [requeridos]=" & Val(t_pedidos) & " , [pedidos]=" & Val(t_oc) & " , [texto_central]='" & RTrim$(t_textocentral) & " '"
     QUERY = QUERY & " where [id_producto]= " & Val(t_basico)
     cn1.BeginTrans
     cn1.Execute QUERY
     
     
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Modifica Precio Producto: " & t_basico & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', '" & Left$(t_detalle, 49) & "', 20, " & 0 & ")"
     cn1.Execute QUERY
     
     If Val(t_stock.Tag) <> Val(t_stock) Then
       'genero mov stock para ajustarlo
       
               Set cl_stock = New STOCK
               cl_stock.sacastock (Val(t_basico))
                 
               If cl_stock.stock_movimientos > Val(t_stock) Then
                   'salida
                    cantajuste = cl_stock.stock_movimientos - Val(t_stock)
                    ubicaajuste = "S"
               Else
                    'entrada
                    cantajuste = Val(t_stock) - cl_stock.stock_movimientos
                    ubicaajuste = "E"
               
               End If
                        
              Set cl_stock = Nothing
       
       
               If cantajuste <> 0 Then
                     Set rss1 = New ADODB.Recordset
                     q = "Select * from g0 where sucursal=0"
                     rss1.Open q, cn1, adOpenDynamic, adLockOptimistic
                     numcomp = rss1("ult_num_ajuste_stock") + 1
                     rss1("ult_num_ajuste_stock") = numcomp
                     rss1.Update
                     Set rss1 = Nothing
            
                     QUERY = "INSERT INTO stk_02([fecha], [letra], [num_comprobante], [id_usuario], [detalle], [sucursal], [tipo_comprobante], [id_proveedor], [id_obra])"
                     QUERY = QUERY & " VALUES ('" & Format$(Now, "dd/mm/yyyy") & "', 'X', " & numcomp & ", " & para.id_usuario & ", 'Ajuste stock auto lista precios', 0, 1, 1,1)"
                     cn1.Execute QUERY
            
                     qr = "SELECT @@IDENTITY AS NewID"
                     Set rss2 = cn1.Execute(qr)
                     numint = rss2.Fields("NewID").Value
                     idcli = 0
              
                    QUERY = "INSERT INTO stk_03([num_int], [RENGLON], [id_producto], [descripcion], [unidad], [detalle], [cantidad], [ubicacion])"
                    QUERY = QUERY & " VALUES (" & numint & ", " & 1 & ", " & Val(t_basico) & ", '" & t_detalle & "', ' ', 'Ajuste auto lista precios', " & cantajuste & " , '" & ubicaajuste & "')"
                    cn1.Execute QUERY
            
                   QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo], [id_cliente])"
                   QUERY = QUERY & " VALUES ('" & Format$(Now, "dd/mm/yyyy") & "', " & Val(t_basico) & ", " & cantajuste & ", '" & ubicaajuste & "', 'Mov.Int.Stk " & Format$(numint, "00000000") & _
                   "', 'Ajuste auto listaprecios', " & numint & ", 'S', " & idcli & ")"
                   cn1.Execute QUERY
              End If
                 
     End If
     
     
     
     
     cn1.CommitTrans
      
      t_stock.Tag = t_stock
   
     'actualizo renglon
     Select Case para.tipolistaprecios
     Case Is = 1
      vta_listaprecios.msf1.TextMatrix(Val(t_linea), 1) = t_detalle
      vta_listaprecios.msf1.TextMatrix(Val(t_linea), 2) = Format$(Val(t_final), "#####0.00")
      vta_listaprecios.msf1.TextMatrix(Val(t_linea), 5) = Format$(Val(c_iva), "###0.00") & "%"
     Case Is = 2
      vta_listaprecios_2.msf1.TextMatrix(Val(t_linea), 1) = t_detalle
      vta_listaprecios_2.msf1.TextMatrix(Val(t_linea), 2) = Format$(Val(t_final), "#####0.00")
      vta_listaprecios_2.msf1.TextMatrix(Val(t_linea), 5) = Format$(Val(c_iva), "###0.00") & "%"
     End Select
     
   
   Me.Hide
    
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub

Private Sub Command3_Click()
d = Format$(Val(t_preciocompra) * Val(t_dtocompra) / 100, "#####0.000")
n = Val(t_preciocompra) - Val(d)
F = n * Val(t_fletecompra) / 100
n2 = F + n
pu = n2 + (n2 * Val(t_utilidad) / 100)
i = d * Val(t_tasaimpint) / 100
fi = Format$(pu + (pu * Val(c_iva) / 100) + i, "######0.00")
If Option1 = True Then
    'dos decimales
     fi = Format(fi, "######0.00")
Else
   If Option2 = True Then
      'entero
       fi = Format(fi, "######0")
   Else
     If Option3 = True Then
       '0.25
       fi = redondeanum(Format(fi, "#######0.00"), 1)
     Else
      If Option4 = True Then
       fi = Format(fi, "######0.0")
      Else
       fi = redondeanum(Format(fi, "######0.00"), 2)
      End If
     End If
   End If
End If


t_sugerido = Format(fi, "#######0.00")
End Sub

Private Sub Command4_Click()
Me.Hide
End Sub

Private Sub Command5_Click()
If Val(t_costo) > 0 Then
  t_pu = Format$(Val(t_costo) + (Val(t_costo) * Val(t_utilidad) / 100), "#####0.000")
  t_impuesto = Format$(Val(t_preciocompra) * Val(t_tasaimpint) / 100, "####0.0000")
End If
't_final = Format$(Val(t_pu) + (Val(t_pu) * Val(c_iva) / 100) + Val(t_impuesto), "######0.00")
t_preciocompra = Format$(Val(t_preciocompra), "#####0.000")
t_dtocompra = Format$(Val(t_dtocompra), "#####0.000")
t_fletecompra = Format$(Val(t_fletecompra), "#####0.000")
t_tasaimpint = Format$(Val(t_tasaimpint), "#####0.000")

End Sub

Private Sub Command6_Click()
t_impuesto = Format$(Val(t_preciocompra) * Val(t_tasaimpint) / 100, "####0.0000")
pf = Val(t_pu) + (Val(t_pu) * Val(c_iva) / 100) + Val(t_impuesto)
If Option1 = True Then
    'dos decimales
     pf = Format(pf, "######0.00")
Else
   If Option2 = True Then
      'entero
       pf = Format(pf, "######0")
   Else
     If Option3 = True Then
       pf = redondeanum(Format(pf, "#######0.00"), 1)
     Else
       pf = redondeanum(Format(pf, "#######0.00"), 2)
     End If
   End If
End If

t_final = Format(pf, "#######0.00")

End Sub

Private Sub Command7_Click()
t_impuesto = Format$(Val(t_preciocompra) * Val(t_tasaimpint) / 100, "####0.0000")
psi = Val(t_final) - Val(t_impuesto)
t_pu = Format$((psi / (1 + (Val(c_iva) / 100))) + Val(t_impuesto), "######0.00")

End Sub

Private Sub Command8_Click()
    vta_listaprecios4.Show
    vta_listaprecios4.t_idprod = t_basico
    vta_listaprecios4.t_prod = t_detalle
End Sub

Private Sub Command9_Click()
On Error GoTo err1
    stk_movprod.Show
    stk_movprod.t_id = t_basico
    stk_movprod.t_prod = t_detalle

Exit Sub
err1:
 'Call errormod
End Sub

Private Sub Form_Activate()
't_final.SetFocus
End Sub

Sub actualiza()
t_preciocompra.Tag = Val(t_preciocompra)
t_costo.Tag = Val(t_costo)
t_final.Tag = Val(t_final)
t_cotizultcom.Tag = Val(t_cotizultcom)
t_stock.Tag = Val(t_stock)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
 Call graba
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Me.Hide
End If

End Sub

Private Sub Form_Load()
  Call carga_grupos(c_grupo)
  c_grupo.ListIndex = 0
  Call carga_deptos_venta(c_depto)
  c_depto.ListIndex = 0
  Call carga_marcas(c_marca)
  c_marca.ListIndex = 0
  Call carga_proveedores(c_prov)
  c_prov.ListIndex = 0
  Call carga_unidad(c_unidad)
  c_unidad.ListIndex = 0
  Call carga_tasaiva(c_iva)
  Call carga_tasaib(c_tasaib)
  Option5 = True
  
  Call carga_redondeo
End Sub

  

Sub carga_redondeo()
Select Case para.tiporedondeo
Case Is = 0
  Option1 = True
Case Is = 1
  Option4 = True
Case Is = 2
 Option3 = True
Case Is = 3
 Option5 = True
Case Is = 4
 Option2 = True
Case Else
 Option1 = True
End Select

End Sub





Private Sub t_basico_GotFocus()
t_basico = ""
End Sub

Private Sub t_detalle_LostFocus()
If t_detalle = "" Then
  t_detalle = "Null"
End If
End Sub

Private Sub t_moneda_LostFocus()
t_moneda = Format$(t_moneda, ">@")
If t_moneda <> "P" And t_moneda <> "D" Then
  t_moneda = "P"
End If
End Sub

Private Sub t_tipo_LostFocus()
t_tipo = Format$(t_tipo, ">@")
If t_tipo <> "M" And t_tipo <> "P" Then
  t_tipo = "P"
End If
End Sub

Private Sub t_tipocarga_LostFocus()
t_tipocarga = Format$(t_tipocarga, ">@")
If t_tipocarga <> "M" And t_tipocarga <> "A" Then
  t_tipocarga = "M"
End If
End Sub
