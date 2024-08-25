VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form ABM_COMP_COMPRA 
   BackColor       =   &H00E0E0E0&
   Caption         =   "COMPROBANTES DE COMPRA"
   ClientHeight    =   9435
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   18165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   18165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Ordenes de Compra Asociadas"
      Height          =   255
      Left            =   15480
      TabIndex        =   64
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   240
      TabIndex        =   61
      Top             =   8280
      Width           =   15015
      Begin VB.Label Label18 
         Caption         =   "Label18"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   9015
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   15360
      TabIndex        =   59
      Top             =   1920
      Width           =   2655
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Comprobnate Electronico"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.TextBox t_ni 
      Height          =   375
      Left            =   5520
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   15360
      TabIndex        =   54
      Top             =   1440
      Width           =   2655
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Moneda Unica"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Percepciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   50
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   15360
      TabIndex        =   46
      Top             =   0
      Width           =   975
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000A&
         Caption         =   "$"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000A&
         Caption         =   "U$s"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   16320
      TabIndex        =   39
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Command4 
         Caption         =   "Forma Pago Ctdo"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Contado "
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   15360
      TabIndex        =   36
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales del Comprobante"
      Height          =   2295
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   15015
      Begin VB.CommandButton Command7 
         Caption         =   "Iva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   66
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox c_cliente 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   5655
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   8640
         Picture         =   "Proc003A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox t_dolares 
         Alignment       =   1  'Right Justify
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
         Left            =   12960
         MaxLength       =   21
         TabIndex        =   20
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox t_obs 
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
         MaxLength       =   79
         TabIndex        =   14
         Text            =   """ "" "
         Top             =   960
         Width           =   5655
      End
      Begin VB.ComboBox c_ib 
         Height          =   315
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   5655
      End
      Begin VB.ComboBox c_ret 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Width           =   5655
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Text            =   "c_cuenta"
         Top             =   240
         Width           =   6855
      End
      Begin VB.TextBox t_total 
         Alignment       =   1  'Right Justify
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
         Left            =   10200
         MaxLength       =   21
         TabIndex        =   19
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox t_iva 
         Alignment       =   1  'Right Justify
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
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   21
         TabIndex        =   18
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox t_perc 
         Alignment       =   1  'Right Justify
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
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   21
         TabIndex        =   17
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox t_nograbado 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         MaxLength       =   21
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox t_subtotal 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         MaxLength       =   21
         TabIndex        =   15
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Stock por Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00400040&
         Caption         =   "Total U$s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   12960
         TabIndex        =   49
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Observaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7560
         TabIndex        =   43
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Conc. Ret.IB"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7560
         TabIndex        =   42
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Conc. Ret.Gan."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta Contable"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10200
         TabIndex        =   33
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "No Grabado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "Subtotal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   2535
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   6165
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   15135
      Begin VB.ComboBox c_zona 
         Height          =   315
         ItemData        =   "Proc003A.frx":030A
         Left            =   9960
         List            =   "Proc003A.frx":0314
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox t_fechavto 
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
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   12600
         Picture         =   "Proc003A.frx":0328
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox t_cotiz 
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
         Left            =   9960
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   13200
         Picture         =   "Proc003A.frx":069A
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   720
         Width           =   1095
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
         Height          =   405
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox c_tipocomp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox t_fecha 
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox t_numoc 
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
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1200
         Width           =   3255
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
         Height          =   405
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox c_prov 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   1
         Text            =   "c_prov"
         Top             =   720
         Width           =   10215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Zona:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8400
         TabIndex        =   58
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Vto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   52
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cotizacion Dolar:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8400
         TabIndex        =   45
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
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
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Proveedor:"
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
      Left            =   16320
      TabIndex        =   22
      Top             =   7800
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Proc003A.frx":079F
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
         Picture         =   "Proc003A.frx":1021
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
      Height          =   405
      Left            =   0
      TabIndex        =   21
      Top             =   9030
      Width           =   18165
      _ExtentX        =   32041
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   20285
            MinWidth        =   20285
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "24/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:18 p.m."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   15600
      TabIndex        =   63
      Top             =   3480
      Width           =   2175
   End
End
Attribute VB_Name = "ABM_COMP_COMPRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Dim numint As Long
Dim totalcompexiste As Double
Dim nicompexiste As Long
Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   T_TOTAL = ""
   
   
  
End Sub
Sub sacatotales2()
  s = 0
  v = 0
  For i = 1 To msf1.Rows - 1
      r = msf1.TextMatrix(i, 7)
      s = s + r
      v = v + (r * msf1.TextMatrix(i, 5) / 100)
  Next i
  t_subtotal = s
  t_iva = v
  Call sacatotales

End Sub
Sub estadoformapago()

If Option1 = True Then
  'ctacte
   c = &HC000&
  Me.Caption = "COMPROBANTES DE COMPRA         ***** CUENTA CORRIENTE *****"
  Label19.BackColor = c
  Label19 = "CUENTA CORRIENTE"
 
  
Else
  'contado
  c = &HFF&
  Me.Caption = "COMPROBANTES DE COMPRA         ##### CONTADO #####"
  Label19.BackColor = c
  Label19 = "CONTADO"
  
End If
Label1.BackColor = c
Label2.BackColor = c
Label3.BackColor = c
Label5.BackColor = c
  
End Sub
Sub carga()
  If t_funcion = "D" Then
    Call limpia
  End If
  
  com_formapago.armagrid2
  
  Set cl_comp = New COMPROBANTES
  Call cl_comp.cargar(c_tipocomp.ItemData(c_tipocomp.ListIndex), t_letra, Val(t_sucursal), Val(t_numoc), c_prov.ItemData(c_prov.ListIndex))
  If cl_comp.numint <> 0 Then
   k = MsgBox("El numero de Comprobante existe para este Proveedor, presione SI para modificar o NO para ingresar otro comprobante con el mismo numero", 4)
   If k = 6 Then
     EXISTE = "S"
     t_ni = cl_comp.numint
     Set rs = New ADODB.Recordset
     q = "select * from a6 where [num_int] = " & cl_comp.numint & " order by [renglon]"
     t_fecha = cl_comp.fecha
     t_fechavto = cl_comp.fechavto
     t_cotiz = cl_comp.cotizacion
     c_prov.ListIndex = buscaindice(c_prov, cl_comp.idproveedor)
     rs.Open q, cn1
     While Not rs.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs("id_producto"), "00000") & Chr(9) & rs("detalle") & Chr(9) & rs("cantidad") & Chr(9) & Format$(rs("pu"), "######0.00") & Chr(9) & rs("tasa_iva") & Chr(9) & rs("descuento") & Chr(9) & rs("importe") & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & rs("pusindto") & Chr(9) & rs("unidad06") & Chr(9) & rs("exportacion") & Chr(9) & rs("envase")
       rs.MoveNext
     Wend
     Set rs = Nothing
     c_ret.ListIndex = buscaindice(c_ret, cl_comp.idcodretgan)
     c_cuenta.ListIndex = buscaindice(c_cuenta, cl_comp.idcuenta)
     t_subtotal = Format$(cl_comp.subtotal, "######0.00")
     t_nograbado = Format$(cl_comp.nograbado, "######0.00")
     t_perc = Format$(cl_comp.percep, "######0.00")
     t_iva = Format$(cl_comp.iva, "######0.00")
     T_TOTAL = Format$(cl_comp.total, "######0.00")
     totalcompexiste = Val(T_TOTAL)
     nicompexiste = cl_comp.numint
     C_CLIENTE.ListIndex = buscaindice(C_CLIENTE, cl_comp.idcliente)
     
     
     
     If cl_comp.moneda = "P" Then
       Option4 = True
     Else
       Option3 = True
     End If
      
     If cl_comp.contado = "N" Then
       Option1 = True
     Else
       Option2 = True
     End If
     
     'cargo percepciones
     Set rs = New ADODB.Recordset
     q = "select * from a13, a12 where a13.[id_percepcion] = a12.[id_percepcion] and [num_int] = " & cl_comp.numint
     rs.Open q, cn1
     ABM_COMP_COMPRA2.msf1.clear
     i = 1
     While Not rs.EOF
       ABM_COMP_COMPRA2.msf1.AddItem i & Chr$(9) & rs("a13.id_percepcion") & Chr$(9) & rs("descripcion") & Chr$(9) & rs("importe") & Chr$(9) & rs("a13.id_cuenta")
       rs.MoveNext
       i = i + 1
     Wend
     Set rs = Nothing
     
     
     'cargo iva
     Call ABM_COMP_COMPRA5.armagrid
     Set rs = New ADODB.Recordset
     q = "select * from a23 where [num_int] = " & cl_comp.numint
     rs.Open q, cn1
     
     i = 1
     While Not rs.EOF
       ABM_COMP_COMPRA5.msf1.AddItem rs("tipo_iva") & Chr$(9) & rs("tasa_iva") & Chr$(9) & rs("neto") & Chr$(9) & rs("iva")
       rs.MoveNext
       i = i + 1
     Wend
     Set rs = Nothing
     Call ABM_COMP_COMPRA5.sacatotales
           
           
           
  Else
     EXISTE = "N"
     t_cotiz = para.cotizacion
     t_ni = 0
     gen_path.t_path = ""
     Call ABM_COMP_COMPRA5.armagrid
   
   End If
  Else
     EXISTE = "N"
     t_ni = 0
     t_cotiz = para.cotizacion
     Call ABM_COMP_COMPRA5.armagrid
  End If
  Set cl_comp = Nothing
End Sub

Private Sub btnacepta_Click()
Call estadoformapago
If verifica Then
If Option2 = True Then
 'contado
   If estadocaja(t_fecha) = "A" Then
    If Val(com_formapago.t_diferencia) = 0 And com_formapago.msf2.Rows > 1 Then
      If c_tipocomp.ItemData(c_tipocomp.ListIndex) <> 5 Then
        abm_comp_compra3.Show
      Else
        abm_COMP_COMPRA4.Show
      End If
    Else
       If com_formapago.msf2.Rows <= 1 Then
           J = MsgBox("No ha ingresado forma de pago, acepta pago total en Efectivo", 4)
           If J = 6 Then
              'pone forma de pago efectivo
              com_formapago.msf2.AddItem "001" & Chr(9) & 1 & Chr(9) & "-" & Chr(9) & "Efectivo $" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & Format$(Val(T_TOTAL), "######0.00") & Chr(9) & Format$(t_fecha, "DD/MM/YYYY") & Chr(9) & "" & Chr(9) & para.cuenta_caja
              If c_tipocomp.ItemData(c_tipocomp.ListIndex) <> 5 Then
                abm_comp_compra3.Show
              Else
                abm_COMP_COMPRA4.Show
              End If
              
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
     If c_tipocomp.ItemData(c_tipocomp.ListIndex) <> 5 Then
        abm_comp_compra3.Show
      Else
        abm_COMP_COMPRA4.Show
      End If
  Else
    MsgBox ("El Proveedor Manual solo puede utilizarse para comprobantes de contado")
  End If
End If
End If


End Sub
Function verifica() As Boolean
v = True
Set rs = New ADODB.Recordset
q = "select * from a15 where [num_int_comp] = " & nicompexiste
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  If Val(T_TOTAL) <> totalcompexiste Then
    MsgBox ("El comrpobante tiene asiganadas OP y no pùede modifcarse el importe total del mismo")
    v = False
  End If

End If
Set rs = Nothing

If c_prov.ItemData(c_prov.ListIndex) = 1 Then
  If Option1 = True Then
     MsgBox ("Solo se puede cargar con proveedor <CONTADO> comprobantes en contado")
     v = False
  End If
End If

verifica = v
End Function
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
msf1.Cols = 15
msf1.ColWidth(0) = 400
msf1.ColWidth(1) = 1200
msf1.ColWidth(2) = 6200
msf1.ColWidth(3) = 1400
msf1.ColWidth(4) = 1400
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 900
msf1.ColWidth(7) = 2000
msf1.ColWidth(8) = 1200
msf1.ColWidth(9) = 1400
msf1.ColWidth(10) = 1400
msf1.ColWidth(11) = 1400
msf1.ColWidth(12) = 1000
msf1.ColWidth(13) = 1000
msf1.ColWidth(14) = 900

msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "P.U."
msf1.TextMatrix(0, 5) = "% Iva"
msf1.TextMatrix(0, 6) = "% Dto"
msf1.TextMatrix(0, 7) = "Importe"
msf1.TextMatrix(0, 8) = "Num.Ref."
msf1.TextMatrix(0, 9) = "PU.Ult.Comp."
msf1.TextMatrix(0, 10) = "Fec.Ult.Comp"
msf1.TextMatrix(0, 11) = "P.U. s/Dto"
msf1.TextMatrix(0, 12) = "Unidad"
msf1.TextMatrix(0, 13) = "Reint.Exp."
msf1.TextMatrix(0, 14) = "Envase"
End Sub







Private Sub c_alicuota_LostFocus()
Call barraesag(Me)
Call sacatotales
End Sub

Private Sub C_cliente_LostFocus()
If C_CLIENTE.ListIndex < 0 Then
  C_CLIENTE.ListIndex = 0
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

Private Sub c_ib_LostFocus()
If c_ib.ListIndex < 0 Then
  c_ib.ListIndex = 0
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


com_proveedor.t_id = c_prov.ItemData(c_prov.ListIndex)
com_proveedor.carga
If c_prov.ItemData(c_prov.ListIndex) = 1 Then
      Option2 = True
      com_proveedor.Show
Else
    Call iniciaprov
End If
  
End Sub
Sub iniciaprov()
  Set cl_prov = New proveedores
  cl_prov.carga (c_prov.ItemData(c_prov.ListIndex))
  If cl_prov.idprov <> 0 Then
    espere.Show
    espere.Label1 = "Comprobante proveedor..."
    espere.Refresh
    t_letra = com_proveedor.t_letra
    c_ret.ListIndex = buscaindice(c_ret, cl_prov.idcodretgan)
    c_ib.ListIndex = buscaindice(c_ib, cl_prov.idcodretib)
    c_cuenta.ListIndex = buscaindice(c_cuenta, cl_prov.idcuenta)
    If buscafacturaapocrifa(Val(com_proveedor.t_cuit)) Then
      MsgBox ("ATENCION!!! El cuit del proveedor se encuentra en el registro de FACTURAS APOCRIFAS entregado por el AFIP")
    End If
    Unload espere
   End If
   Set cl_prov = Nothing
End Sub
Private Sub c_ret_LostFocus()
If c_ret.ListIndex < 0 Then
  c_ret.ListIndex = 0
End If
End Sub

Private Sub c_tipocomp_LostFocus()
If c_tipocomp.ListIndex < 0 Then
  If Val(c_tipocomp) > 0 Then
    c_tipocomp.ListIndex = buscaindice(c_tipocomp, Val(c_tipocomp))
  Else
    c_tipocomp.ListIndex = 0
  End If
End If

End Sub

Private Sub c_zona_LostFocus()
If c_zona.ListIndex < 0 Then
   c_zona.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
ABM_PROv.Show
End Sub

Private Sub Command1_LostFocus()
c_prov.clear
Call carga_proveedores(c_prov)
c_prov.ListIndex = 0
End Sub


Private Sub Command2_Click()
  ABM_COMP_COMPRA2.t_modulo = "C"
  ABM_COMP_COMPRA2.Show
End Sub

Private Sub Command3_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Command4_Click()
  com_formapago.Show
  com_formapago.T_TOTAL = T_TOTAL

End Sub

Private Sub Command5_Click()
com_proveedor.t_id = c_prov.ItemData(c_prov.ListIndex)
com_proveedor.carga
com_proveedor.Show

End Sub




Private Sub Command6_Click()

com_seloc.carga
com_seloc.Show
End Sub

Private Sub Command7_Click()
  ABM_COMP_COMPRA5.Show

End Sub

Private Sub Form_Activate()
If para.cuenta_sel > 0 Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_sel)
  para.cuenta_sel = 0
End If


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
   tm = tm & " **" & com_proveedor.t_cli & "**"
Else
    tm = tm & " **" & c_prov & "**"
End If
Label18 = UCase$(tm)
Frame9.Visible = True

End Sub

Sub grabaformapago()
 
  
  For i = 1 To com_formapago.msf2.Rows - 1
        QUERY = "INSERT INTO a7([num_int], [secuencia], [id_formapago], [formapago], [num_ch], [detalle_banco], [sucursal], [titular], [importe], [fecha_dif], [num_int_fp])"
        QUERY = QUERY & " VALUES (" & numint & ", " & i & ", " & Val(com_formapago.msf2.TextMatrix(i, 0)) & ", '" & Left$(com_formapago.msf2.TextMatrix(i, 1), 20) & "', " & Val(com_formapago.msf2.TextMatrix(i, 2)) & ", '" & Left$(com_formapago.msf2.TextMatrix(i, 3), 25) & "', '" & Left$(com_formapago.msf2.TextMatrix(i, 4), 25) & "', '" & Left$(com_formapago.msf2.TextMatrix(i, 5), 25) & "', " & Val(com_formapago.msf2.TextMatrix(i, 6)) & ", '" & _
        com_formapago.msf2.TextMatrix(i, 7) & "', " & Val(com_formapago.msf2.TextMatrix(i, 8)) & ")"
        cn1.Execute QUERY

        
       If Val(com_formapago.msf2.TextMatrix(i, 0)) = 3 Then 'ch. terceros
        Set rs = New ADODB.Recordset
        q = "select [estado], [destino], [num_int_op], [fecha_salida], [tipo_salida] from cyb_03 where [num_interno] = " & Val(com_formapago.msf2.TextMatrix(i, 8))
        rs.MaxRecords = 1
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
          rs("estado") = "P"
          rs("destino") = Left$(c_prov, 50)
          rs("num_int_op") = numint
          rs("fecha_salida") = t_fecha
          rs("tipo_salida") = "P"
          numintch = Val(com_formapago.msf2.TextMatrix(i, 8))
          rs.Update
         Else
          numintch = 0
        End If
        Set rs = Nothing
       Else
        numintch = 0
       End If
      
      
      If Val(com_formapago.msf2.TextMatrix(i, 0)) = 4 Then 'Transf. o db. Automatico
                
        'emito debito
        QUERY = "INSERT INTO cyb_04([id_banco], [fecha], [importe], [id_tipomov], [fecha_dif], [ubicacion], [entro], [fecha_acreed], [num_comp], [detalle], [modulo], [num_mov_int], [id_tipodbcr], [num_mov_int_compras])"
        QUERY = QUERY & " VALUES (" & Val(com_formapago.msf2.TextMatrix(i, 5)) & ", '" & com_formapago.msf2.TextMatrix(i, 7) & "', " & Val(com_formapago.msf2.TextMatrix(i, 6)) & ", 90, '" & com_formapago.msf2.TextMatrix(i, 7) & "', 'D', 'N', '" & com_formapago.msf2.TextMatrix(i, 7) & "', " & Val(com_formapago.msf2.TextMatrix(i, 2)) & ", '" & t_letra & Format$(Val(t_suc), "0000") & "-" & Format$(Val(t_numoc), "00000000") & " " & Left$(c_prov, 20) & "', 'C', " & numint & ", 1, 0)"
        cn1.Execute QUERY
        
       
       End If
      
      
      If Val(com_formapago.msf2.TextMatrix(i, 0)) >= 50 Then 'ch. propios
         
        Set rs = New ADODB.Recordset
        q = "select [estado], [fecha_dif], [fecha_emision], [destino], [importe], [num_int_op] from cyb_02 where [id_banco] = " & Val(com_formapago.msf2.TextMatrix(i, 0)) & " and [num_cheque] = " & Val(com_formapago.msf2.TextMatrix(i, 2))
        rs.MaxRecords = 1
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
          If rs("estado") = "P" Then
             rs("estado") = "E"
             rs("fecha_dif") = com_formapago.msf2.TextMatrix(i, 7)
             rs("fecha_emision") = t_fecha
             rs("destino") = c_prov
             rs("importe") = Val(com_formapago.msf2.TextMatrix(i, 6))
             rs("num_int_op") = numint
             rs.Update
             
          Else
             MsgBox ("Error al asignar ch. propio")
          End If
        End If
        Set rs = Nothing
       
        
       
        
        
        'emito ch.
        QUERY = "INSERT INTO cyb_04([id_banco], [fecha], [importe], [id_tipomov], [fecha_dif], [ubicacion], [entro], [fecha_acreed], [num_comp], [detalle], [modulo], [num_mov_int], [id_tipodbcr], [num_mov_int_compras])"
        QUERY = QUERY & " VALUES (" & Val(com_formapago.msf2.TextMatrix(i, 0)) & ", '" & t_fecha & "', " & Val(com_formapago.msf2.TextMatrix(i, 6)) & ", 1, '" & com_formapago.msf2.TextMatrix(i, 7) & "', 'D', 'N', '" & t_fecha & "', " & Val(com_formapago.msf2.TextMatrix(i, 2)) & ", 'FC" & Format$(Val(t_numoc), "00000000") & " " & Left$(c_prov, 28) & "', 'C', " & numint & ", 1, 0)"
        cn1.Execute QUERY
        
       
       End If
 
      
      'la cuenta contrapartida es una sola
      'If msf1.Rows > 1 Then
            'utilizo la cuenta del primer comprobante
      '      cuentacontra = Val(msf1.TextMatrix(1, 7))
      
      'Else
            'utilizo la cuenta de compras varias
      '      cuentacontra = para.cuenta_compras_varias
      'End If
      cuentacontra = c_cuenta.ItemData(c_cuenta.ListIndex)
      
      
      Set rs = New ADODB.Recordset
      q = "select [caja] from cyb_01 where [id_forma_pago] = " & Val(com_formapago.msf2.TextMatrix(i, 0))
      rs.MaxRecords = 1
      rs.Open q, cn1
      If Not rs.BOF And Not rs.EOF Then
        If rs("caja") = "S" Then
                    
          'grabo mov caja
           QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
           QUERY = QUERY & " VALUES (" & Val(com_formapago.msf2.TextMatrix(i, 9)) & ", " & cuentacontra & ", '" & Left$(com_proveedor.t_cli, 49) & " ', " & Val(com_formapago.msf2.TextMatrix(i, 6)) & ", 'H', '" & t_fecha & "', " & numint & ", 'C', 'FC. " & Format$(t_sucursal, "0000") & "-" & Format$(t_numoc, "00000000") & "', " & Val(com_formapago.msf2.TextMatrix(i, 0)) & ", " & numintch & ", " & para.id_usuario & ")"
           cn1.Execute QUERY
        End If
      End If
      Set rs = Nothing
     
      Next i

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call TabEnter2(Me, 20)
End If


End Sub

Private Sub Form_Load()

Call INICIALIZA2(Me)
Frame9.Visible = False
Call carga_proveedores(c_prov)
c_prov.ListIndex = 0
Call carga_tipocomp(c_tipocomp)
c_tipocomp.ListIndex = buscaindice(c_tipocomp, 1)
Call carga_impuestos(c_ret, 217) 'ret gan
c_ret.ListIndex = 0

Call carga_impuestos(c_ib, 50) 'ret ib
c_ib.ListIndex = 0

Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "Sin Imputacion", 0
c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_inventario)

Call carga_clientes(C_CLIENTE)
C_CLIENTE.AddItem "<Todos>", 0
C_CLIENTE.ListIndex = 0



c_cuenta.AddItem "Sin Imputacion", 0
Call armagrid
Call barraesag(Me)

Call estadoformapago

Load abm_COMP_COMPRA1
Load ABM_COMP_COMPRA2
Load ver_PROD_oc
Option1 = True
If para.moneda = "P" Then
  Option4 = True
Else
  Option3 = True
End If


Call barraesag(Me)




para.cuenta_sel = 0
Check1 = 0
Load com_formapago

c_zona.ListIndex = 0

Load com_seloc
com_seloc.limpia

'Load ABM_COMP_COMPRA5

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_COMP_COMPRA1
Unload ABM_COMP_COMPRA2
Unload ver_PROD_oc
Unload com_formapago
Unload com_seloc
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Continua - [F3] Pendientes "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  ver_PROD_oc.c_prov.ListIndex = buscaindice(ver_PROD_oc.c_prov, c_prov.ItemData(c_prov.ListIndex))
  ver_PROD_oc.Show
End If


If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If



If KeyCode = vbKeyF9 Then
  If msf1.Rows > 1 Then
     Call sacatotales2
  End If
  Frame2.Enabled = True
  c_cuenta.SetFocus
  Call estadoformapago
End If

If KeyCode = vbKeyInsert Then
   abm_COMP_COMPRA1.t_renglon = ""
   abm_COMP_COMPRA1.t_cantidad = ""
   abm_COMP_COMPRA1.t_pu = ""
   abm_COMP_COMPRA1.t_importe = ""
   abm_COMP_COMPRA1.t_ref = ""

   abm_COMP_COMPRA1.Show
End If


End Sub

Sub graba(z As Integer)
'z determina el tipo de actualizacion
'1 actualiza precio de venta y de compra
'2 actualiza solo precio compra
'3 no actualiza nada


If EXISTE = "S" Then
  Set cl_comp = New COMPROBANTES
  Call cl_comp.cargar2(Val(t_ni))
  If cl_comp.numint <> 0 Then
    ssi = cl_comp.saldoimpago
    ep = cl_comp.estado_pago
    cl_comp.borrar
    cp = cl_comp.num_op
    
  Else
    ssi = Val(T_TOTAL)
    ep = "N"
    cp = "0000-00000000"
  End If
  Set cl_comp = Nothing
Else
 'no existe
  If Option4 = True Then
     ssi = Val(T_TOTAL)
  Else
     ssi = Val(t_dolares)
  End If
  
  If Option1 = True Then
    ep = "N"
    cp = "0000-00000000"
    ctdo = "N"
  Else
     ep = "S"
     cp = "ctdo"
     ssi = 0
     ctdo = "S"
  End If

End If

   ' nueva
      
     ' On Error GoTo ERRORGRABA
      numint = saca_ultnumero_int_comp("C")
      
      Set cl_comp = New COMPROBANTES
      cl_comp.actual (c_tipocomp.ItemData(c_tipocomp.ListIndex))
      
      If c_cuenta.ListIndex = 0 Then
            c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_compras_varias)
      End If
      
      If Option1 = True Then
        cc = cl_comp.ctacte
        ctdo = "N"
      Else
        cc = "N"
        ctdo = "S"
      End If
       
      
      t_obs = RTrim$(t_obs) & " "
      
      If Option4 = True Then
        moneda = "P"
      Else
        moneda = "D"
      End If
      
      
      If Check1 = 0 Then
        tom = Val(t_dolares)
      Else
        tom = 0
      End If
      
      
      Set cl_prov = New proveedores
      cl_prov.carga (c_prov.ItemData(c_prov.ListIndex))
      
      If C_CLIENTE.ListIndex > 0 Then
        idcli = C_CLIENTE.ItemData(C_CLIENTE.ListIndex)
      Else
        idcli = 0
      End If
      
      cn1.BeginTrans
      
      QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], " & _
" [no_grabado], [percep_ret], [iva], [total], [fecha_prob_entrega], [fecha_recepcion], [estado], [id_codretgan], [id_cuenta], [stock], [ctacte], [grabado], " & _
" [estado_pago], [num_op], [id_codretib], [obs], [condiciones], [info_contacto], [moneda], [cotiz_dolar], [contado], [TOTAL_D], [monto_suj_ret], " & _
"[alicuota_ret], [ret_mes], [pagos_realizados], [pago_actual], [minimo_no_imp], [fecha_vto], [COMPRA], [saldo_impago], [zona], [cuit05], [proveedor05], [id_cliente])"
      
 QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numoc) & ", '" & t_letra & "', " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & _
 ", " & c_prov.ItemData(c_prov.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(t_subtotal) & ", " & Val(t_nograbado) & ", " & Val(t_perc) & ", " & Val(t_iva) & _
 ", " & Val(T_TOTAL) & ", '" & Format$(Now, "dd/mm/yyyy") & "', '" & t_fecha & "', 'A', " & c_ret.ItemData(c_ret.ListIndex) & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & _
 ", '" & cl_comp.STOCK & "', '" & cc & "', '" & cl_comp.grabado & "', '" & ep & "', '" & cp & "', " & c_ib.ItemData(c_ib.ListIndex) & ", '" & t_obs & "', ' ', ' ', '" & moneda & "', " & _
 Val(t_cotiz) & ", '" & ctdo & "', " & tom & ", 0, 0, 0, 0, 0, 0, '" & t_fechavto & "', '" & cl_comp.compra & "', " & ssi & ", " & c_zona.ListIndex + 1 & ", " & Val(com_proveedor.t_cuit) & ", '" & Left$(Trim$(com_proveedor.t_cli), 50) & _
 "', " & idcli & ")"
' MsgBox (QUERY)
 cn1.Execute QUERY
   
   Set cl_prov = Nothing
      
   
  
  'actualiza tabla iva
  If ABM_COMP_COMPRA5.msf1.Rows > 1 Then
    'actualizo deste abm_comp_compra5
     i = 1
     While i < ABM_COMP_COMPRA5.msf1.Rows
      QUERY = "insert into a23([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta23])"
      QUERY = QUERY & " VALUES (" & numint & ", " & Val(ABM_COMP_COMPRA5.msf1.TextMatrix(i, 1)) & ", " & Val(Val(ABM_COMP_COMPRA5.msf1.TextMatrix(i, 3))) & ", " & Val(ABM_COMP_COMPRA5.msf1.TextMatrix(i, 2)) & ", " & Val(ABM_COMP_COMPRA5.msf1.TextMatrix(i, 0)) & ", " & para.cuenta_iva_compras & ")"
      cn1.Execute QUERY
      i = i + 1
    Wend
        
   
  Else
    'actualizo directamente desde formulario
    QUERY = "insert into a23([num_int], [tasa_iva], [iva], [neto], [tipo_iva], [id_cuenta23])"
    QUERY = QUERY & " VALUES (" & numint & ", 0, " & Val(t_iva) & ", " & Val(t_subtotal) & ", 3, " & para.cuenta_iva_compras & ")"
    cn1.Execute QUERY
  
  End If
  Call ABM_COMP_COMPRA5.armagrid  'limpio tabla de ivas
  
  If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 5 Then
    'grabo subsidio a los combustibles
    QUERY = "insert into a19([num_int], [litros], [pu_impuesto_int], [importe], [ubicacion], [detalle], [fecha])"
    QUERY = QUERY & " VALUES (" & numint & ", " & Val(abm_COMP_COMPRA4.t_cantidad) & ", " & Val(abm_COMP_COMPRA4.t_pu) & ", " & Val(abm_COMP_COMPRA4.t_importe) & ", 'D', '" & Left$(c_prov, 50) & "', '" & t_fecha & "')"
    cn1.Execute QUERY
  End If
  
      
  If c_prov.ItemData(c_prov.ListIndex) = 1 Then
    'grabo datos de clientes contado
    QUERY = "insert into a22([num_int], [proveedor22], [direccion22], [cuit22], [id_tipoiva22], [localidad22])"
    QUERY = QUERY & " VALUES (" & numint & ", '" & Left$(com_proveedor.t_cli, 49) & " ', '" & Left$(com_proveedor.t_direccion, 49) & " ', " & Val(com_proveedor.t_cuit) & ", " & com_proveedor.c_iva.ItemData(com_proveedor.c_iva.ListIndex) & ", '" & Left$(com_proveedor.t_localidad, 49) & " ')"
    cn1.Execute QUERY
  End If
      
      
           
      ingreso_productos = 0
      COSTOINV = 0
      
      'busco si se tiene asociado una oc.
      confirmareception = 4 'NO
      If msf1.Rows > 1 Then
         confirmareception = MsgBox("Realiza Recepcion de Mercaderia(S/N)", 4) '6 = si , 4 = no
      End If
      
      
      For i = 1 To msf1.Rows - 1
        'tipoprecio actulaizacion
        Set rs2 = New ADODB.Recordset
        q = "select * from g0 where [sucursal] = 0"
        rs2.Open q, cn1
        If rs2("tipo_actu_pu_compra") = "C" Then
          pu02 = Val(msf1.TextMatrix(i, 4))
        Else
          pu02 = Val(msf1.TextMatrix(i, 11))
        End If
        Set rs2 = Nothing
        
        
        If Val(msf1.TextMatrix(i, 1)) > 1 Then
          Set cl_prod = New productos
          cl_prod.cargar (Val(msf1.TextMatrix(i, 1)))
          costo = cl_prod.costoreal
          Set cl_prod = Nothing
        Else
          costo = 0
        End If
        
        
        'referencias
        If Val(msf1.TextMatrix(i, 8)) > 0 Then 'nro ref.
            Set rs2 = New ADODB.Recordset
            q = "select * from pro_04 where [num_referencia] = " & Val(msf1.TextMatrix(i, 8))
            rs2.MaxRecords = 1
            rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
            If Not rs2.EOF And Not rs2.BOF Then
             If confirmareception = 6 Then
               TOTALR = rs2("total_recibido") + Val(msf1.TextMatrix(i, 3))
               If TOTALR >= rs2("total_PEDIDO") Then
                  estado = "C"
               Else
                  estado = "I"
               End If
               rs2("total_recibido") = TOTALR
               rs2("estado_oc") = estado
             End If
             If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 1 Then
                rs2("total_facturado") = rs2("total_facturado") + Val(msf1.TextMatrix(i, 3))
             End If
             rs2.Update
            End If
           
           nr = Val(msf1.TextMatrix(i, 8))
        Else
            'creo una entrada por producto para seguirlo por el sistema
            'num_referencia auto
            Set rs2 = New ADODB.Recordset
            q = "select * from pro_04"
            rs2.MaxRecords = 1
            rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
            rs2.AddNew
            rs2("id_producto") = Val(msf1.TextMatrix(i, 1))
            rs2("detalle") = msf1.TextMatrix(i, 2)
            rs2("total_pedido") = 0
            rs2("total_oc") = 0
            If confirmareception = 6 Then
               rs2("total_recibido") = Val(msf1.TextMatrix(i, 3))
            Else
               rs2("total_recibido") = 0
            End If
            rs2("estado_pedido") = "C"  'O.C Completada
            rs2("estado_oc") = "C"
            rs2("fecha") = t_fecha
            rs2("id_usuario") = para.id_usuario
            rs2("observaciones") = RTrim$(msf1.TextMatrix(i, 4)) & " "
            rs2("fecha_esperado") = t_fecha
            rs2("id_obra") = 1
            If c_tipocomp.ItemData(c_tipocomp.ListIndex) = 1 Then
                rs2("total_facturado") = Val(msf1.TextMatrix(i, 3))
            Else
                rs2("total_facturado") = 0
            End If
            rs2.Update
            nr = rs2("num_referencia")
            Set rs2 = Nothing
        End If
        
        q = "select * from pro_05 where [num_referencia] = " & nr
        Set rs2 = New ADODB.Recordset
        rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs2.EOF And Not rs2.BOF Then
           rs2.MoveLast
           s = rs2("secuencia") + 1
        Else
           s = 1
        End If
        Set rs2 = Nothing
        QUERY = "INSERT INTO pro_05([num_referencia], [secuencia], [modulo], [num_int], [cantidad], [tipo_comprobante], [fecha])"
        QUERY = QUERY & " VALUES (" & nr & ", " & s & ", 'C', " & numint & ", " & Val(msf1.TextMatrix(i, 3)) & ", 3, '" & t_fecha & "')"
        cn1.Execute QUERY
                 

        QUERY = "INSERT INTO a6([num_int], [RENGLON], [id_producto], [detalle], [cantidad], [pu], [importe], [envase], [bultos],[id_requisicion],[estado],[tasa_iva], [num_int_ITEM], [descuento], [pusindto], [unidad06])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 3)) & ", " & Val(msf1.TextMatrix(i, 4)) & ", " & Val(msf1.TextMatrix(i, 7)) & ", " & Val(msf1.TextMatrix(i, 14)) & ", 0, 0,'R', " & Val(msf1.TextMatrix(i, 5)) & ", " & nr & ", " & Val(msf1.TextMatrix(i, 6)) & ", " & Val(msf1.TextMatrix(i, 11)) & ", '" & msf1.TextMatrix(i, 12) & "')"
        cn1.Execute QUERY
      
               
        If Val(msf1.TextMatrix(i, 1)) > 1 Then
         
         If cl_comp.STOCK <> "N" Then
           QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo], [id_cliente])"
           QUERY = QUERY & " VALUES ('" & t_fecha & "', " & Val(msf1.TextMatrix(i, 1)) & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & cl_comp.STOCK & "', '" & cl_comp.abreviatura & t_letra & Format$(t_sucursal, "0000") & "-" & Format$(t_numoc, "00000000") & _
           "', '" & c_prov & "', " & numint & ",'C', " & idcli & ")"
           cn1.Execute QUERY
          
           If cl_comp.STOCK = "E" Then
              s = Val(msf1.TextMatrix(i, 3))
              COSTOINV = COSTOINV + (costo * s)
              If EXISTE = "N" Then
                Set rs = New ADODB.Recordset
                q = "select [pedidos] from a2 where [id_producto] = " & Val(msf1.TextMatrix(i, 1))
                rs.MaxRecords = 1
                rs.Open q, cn1
                If Not rs.EOF And Not rs.BOF Then
                  pqq = rs("pedidos") - Val(msf1.TextMatrix(i, 3))
                  If pqq < 0 Then
                     pqq = 0
                  End If
                  qp = ", [pedidos]= " & pqq
                 Else
                  qp = ""
                 End If
              Else
               qp = ""
              End If
           Else
              COSTOINV = COSTOINV - (costo * s)
              s = -Val(msf1.TextMatrix(i, 3))
              qp = ""
           End If
           
           
           QUERY = "update a2 set [stock] = [stock] + " & s & qp & " where [id_producto] = " & Val(msf1.TextMatrix(i, 1))
           
           cn1.Execute QUERY
          
          
          
          
          End If
      
          If cl_comp.compra <> "N" Then
           ultcom = t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numoc), "00000000") & " | " & Left$(c_prov, 20) & " | " & t_fecha & " | " & Format$(pu02, "#####0.00")
           QUERY = "update a2 set  [ultima_compra]='" & Left$(ultcom, 50) & "'"
           QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
           cn1.Execute QUERY
          End If
        
          If z < 3 Then  'actualizo precio de compra o de venta
            
            Set rs = New ADODB.Recordset
            q = "select * from a2 where [id_producto] = " & Val(msf1.TextMatrix(i, 1))
            rs.Open q, cn1
            If Not rs.EOF And Not rs.BOF Then
                 'saco datos para calculo
                  descuento = rs("dto_compra")
                  flete = rs("flete_compra")
                  tasaimpuesto = rs("tasa_imp_interno")
                  utilidad = rs("porc_utilidad")
                  
                  
                  'calculo estructura costo
                  d = Format$(pu02 * descuento / 100, "#####0.000")
                  n = pu02 - Val(d)
                  F = n * flete / 100
                  n2 = F + n
                  costo = Format$(n2, "#####0.000")
                  QUERY = "update a2 set  [precio_ult_compra]=" & pu02 & " , [fecha_ult_compra]='" & t_fecha & "' , [costoreal]=" & Val(costo) & " , [id_proveedor_ult_compra]=" & c_prov.ItemData(c_prov.ListIndex) & ", [num_int_ult_compra] = " & numint & ", [dolar_ult_compra]=" & Val(t_cotiz)
                  QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
                  cn1.Execute QUERY

                   
                  If z = 1 Then
                    'precio unitario
                    If Val(costo) > 0 Then
                     pu = Format$(Val(costo) + (Val(costo) * utilidad / 100), "#####0.000")
                     impuesto = Format$(pu02 * tasaimpuesto / 100, "####0.0000")
                     final = Format$(Val(pu) + (Val(pu) * Val(msf1.TextMatrix(i, 5)) / 100) + Val(impuesto), "######0.00")
                     QUERY = "update a2 set  [pu]=" & Val(pu) & " , [fecha_actu_precio_venta]='" & t_fecha & "' , [precio_final]=" & Val(final)
                     QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
                     cn1.Execute QUERY
                    End If
                  End If
               End If
               Set rs = Nothing
          End If
        End If
      Next i
      
      
      If Val(t_perc) > 0 Then
        For i = 1 To ABM_COMP_COMPRA2.msf1.Rows - 1
          QUERY = "INSERT INTO a13([num_int], [secuencia], [id_percepcion], [importe], [id_cuenta], [cod_regimen])"
          QUERY = QUERY & " VALUES (" & numint & ", " & i & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 1) & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 3) & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 4) & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 6) & ")"
          cn1.Execute QUERY
        Next i
      End If
      
      If Option2 = True Then
        Call grabaformapago
      End If
      
     If Generaasientosauto Then
      If cl_comp.contabilidad <> "N" Then
         'grabo asiento
         numintcgr = saca_ultnumero_int_comp("G")

         
         u1 = cl_comp.contabilidad
          
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         
         If Option4 = True Then
            m5 = 1
         Else
           m5 = Val(t_cotiz)
         End If
         
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Compras] " & cl_comp.abreviatura & " " & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numoc), "00000000") & "', 'C', " & numint & ", " & Val(T_TOTAL) * m5 & ", " & Val(T_TOTAL) * m5 & ", " & para.id_usuario & ", '" & Left$(RTrim$(c_prov), 50) & "')"
         cn1.Execute QUERY
      
         ic = 1
        
         
         'cuenta madre ctacte o caja
          If Option1 = True Then
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
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u1 & "', " & Format(Val(T_TOTAL) * m5, "######0.00") & ", '" & dcta & "')"
           cn1.Execute QUERY
           ic = ic + 1
          Else
           'ingresa forma de pago
            For i = 1 To com_formapago.msf2.Rows - 1
               cta = Val(com_formapago.msf2.TextMatrix(i, 9))
               Set rs = New ADODB.Recordset
               q = "select * from c_01 where [id_cuenta] = " & cta
               rs.Open q, cn1
               If Not rs.EOF And Not rs.BOF Then
                 dcta = rs("descripcion")
               Else
                  dcta = "Cuenta Inexistente"
               End If
               Set rs = Nothing
               
               im = Format(Val(com_formapago.msf2.TextMatrix(i, 6)) * m5, "######0.00")
               dcta = com_formapago.msf2.TextMatrix(i, 3)
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
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & u2 & "', " & Val(t_nograbado) * m5 & ", 'No Grabado')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
                   
         If Val(t_perc) > 0 Then
           'cuenta perc
           For i = 1 To ABM_COMP_COMPRA2.msf1.Rows - 1
              QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
              QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & ABM_COMP_COMPRA2.msf1.TextMatrix(i, 4) & ", '" & u2 & "', " & Val(ABM_COMP_COMPRA2.msf1.TextMatrix(i, 3)) * m5 & ", 'Perc.')"
              cn1.Execute QUERY
              ic = ic + 1
           Next i
         End If
          
         If Val(t_iva) > 0 Then
           'cuenta perc
           QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
           QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & para.cuenta_iva_compras & ", '" & u2 & "', " & Val(t_iva) * m5 & ", 'IVA')"
           cn1.Execute QUERY
           ic = ic + 1
         End If
         
         'contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & u2 & "', " & Val(t_subtotal) * m5 & ", '" & c_cuenta & "')"
         cn1.Execute QUERY
      
       End If
          
      End If
      
      
      'actualizo esto ordenes de compra
       If com_seloc.msf1.Rows > 1 Then
        For i = 1 To com_seloc.msf1.Rows - 1
          If com_seloc.msf1.TextMatrix(i, 0) = "**" Then
             QUERY = "update a5 set  [estado]='F'"
             QUERY = QUERY & " where [num_int]= " & Val(com_seloc.msf1.TextMatrix(i, 4))
             cn1.Execute QUERY
          End If
        Next i
      End If
      
      
      
      
      'log
       nc = "(" & c_tipocomp.ItemData(c_tipocomp.ListIndex) & ")" & t_letra & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numoc), "00000000")
      
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Ingreso Comprobante Compra:" & numint & "', " & para.id_usuario & ", 'C', " & numint & ", '" & Now & "', '[" & nc & "', 101, " & c_prov.ItemData(c_prov.ListIndex) & ")"
  
     cn1.Execute QUERY
      
      cn1.CommitTrans
      Set rs = Nothing
      
      
      
      
      Call INICIALIZA2(Me)
      Label18 = ""
      Frame9.Visible = False
      Call armagrid
      ABM_COMP_COMPRA2.armagrid
      c_tipocomp.SetFocus

     If Check2 = 1 Then
       Load gen_path
       gen_path.t_id = numint
       gen_path.t_modulo = "Compras"
       gen_path.t_origen = "C"
       gen_path.t_path = ""
       gen_path.Show
     End If


Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Sub actureq(r As Long)
'actualizo requisiciones
          p = 0
          Set rs = New ADODB.Recordset
          q = "select * from a3 where [id_renglon] = " & r
          rs.Open q, cn1, adOpenStatic, adLockOptimistic
          If Not rs.EOF And Not rs.BOF Then
             ca = rs("cantidad_requisicion") - rs("cant_pedida") 'falta pedir
             If cantidadp >= ca Then
                p = ca
                rs("cant_pedida") = rs("cant_pedida") + p
                cantidadp = cantidadp - p
                rs("estado") = "P"
                
             Else
                p = cantidadp
                rs("cant_pedida") = rs("cant_pedida") + p
                cantidadp = 0
             End If
             rs("fecha_prob_entrega") = t_fechaprob
             rs.Update
           End If
           cp = rs("id_producto")
           Set rs = Nothing
           
           Set cl_prod = New productos
           Call cl_prod.actualizar(cp, 0, -p, p)
           Set cl_prod = Nothing
End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If msf1.Row > 0 Then
    abm_COMP_COMPRA1.t_renglon = msf1.Row
    abm_COMP_COMPRA1.t_basico = msf1.TextMatrix(msf1.Row, 1)
    abm_COMP_COMPRA1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    abm_COMP_COMPRA1.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
    abm_COMP_COMPRA1.t_pu = msf1.TextMatrix(msf1.Row, 4)
    abm_COMP_COMPRA1.t_dto = msf1.TextMatrix(msf1.Row, 6)
    abm_COMP_COMPRA1.t_importe = msf1.TextMatrix(msf1.Row, 7)
    abm_COMP_COMPRA1.t_ref = Val(msf1.TextMatrix(msf1.Row, 8))
    abm_COMP_COMPRA1.c_tasa.ListIndex = buscaindice(abm_COMP_COMPRA1.c_tasa, Val(msf1.TextMatrix(msf1.Row, 5)))
    abm_COMP_COMPRA1.t_unidad = msf1.TextMatrix(msf1.Row, 12)
    abm_COMP_COMPRA1.t_envase = msf1.TextMatrix(msf1.Row, 14)
  
    abm_COMP_COMPRA1.Show
    End If
End If


 



End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True
End Sub

Private Sub Option1_Click()
Call estadoformapago
End Sub

Private Sub Option1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub Option1_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub Option2_Click()
Call estadoformapago
End Sub

Private Sub Option2_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub Option2_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub Option3_Click()
Label16 = "Total $"
End Sub

Private Sub Option4_Click()
Label16 = "Total U$s"
End Sub

Private Sub t_cotiz_LostFocus()
If Val(t_cotiz) < 1 Then
   t_cotiz = para.cotizacion
End If
End Sub

Private Sub t_dolares_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If

End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
If verificaperiodo(t_fecha) = "C" Then
   MsgBox ("El periodo para el cual se deseas ingresar el comprobante esta CERRADO!!!!!")
   t_fecha.SetFocus
   t_fecha = ""
End If

End Sub


Private Sub t_fechavto_LostFocus()
If Not IsDate(t_fechavto) Then
  t_fechavto = Format$(t_fecha, "dd/mm/yyyy")
Else
  t_fechavto = Format$(t_fechavto, "dd/mm/yyyy")
End If
End Sub

Private Sub t_iva_LostFocus()
Call sacatotales
End Sub

Private Sub t_letra_LostFocus()
  t_letra = com_proveedor.t_letra


End Sub

Private Sub t_numoc_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_numoc_LostFocus()
If IsNumeric(t_numoc) Then
   t_numoc = Format$(t_numoc, "00000000")
   Call carga
End If
Call mensaje

End Sub

Private Sub t_obs_LostFocus()
t_obs = RTrim$(t_obs) & " "
End Sub


Private Sub t_perc_GotFocus()
StatusBar1.Panels(2) = "<INS> Ingresa Percepciones -  [ENTER] ACEPTA"
End Sub

Private Sub t_perc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyins Then
  ABM_COMP_COMPRA2.Show
End If
End Sub

Private Sub t_subtotal_LostFocus()
If msf1.Rows > 1 Then
 Call sacatotales
End If
End Sub
Sub sacatotales()
t_subtotal = Format$(Val(t_subtotal), "######0.00")
t_nograbado = Format$(Val(t_nograbado), "######0.00")
If ABM_COMP_COMPRA2.msf1.Rows > 1 Then
  t = 0
  For i = 1 To ABM_COMP_COMPRA2.msf1.Rows - 1
    t = t + Val(ABM_COMP_COMPRA2.msf1.TextMatrix(i, 3))
  Next i
  t_perc = Format$(t, "######0.00")
Else
  t_perc = Format$(0, "######0.00")
End If

If ABM_COMP_COMPRA5.msf1.Rows > 1 Then
  t = 0
  For i = 1 To ABM_COMP_COMPRA5.msf1.Rows - 1
    t = t + Val(ABM_COMP_COMPRA5.msf1.TextMatrix(i, 3))
  Next i
  t_iva = Format$(t, "######0.00")
Else
  t_iva = Format$(Val(t_subtotal) * Val(c_alicuota) / 100, "######0.00")
End If

T_TOTAL = Format$(Val(t_subtotal) + Val(t_nograbado) + Val(t_perc) + Val(t_iva), "######0.00")
If Val(t_cotiz) < 1 Then
   t_cotiz = 1
End If
If Option4 = True Then
  t_dolares = Format$(Val(T_TOTAL) / Val(t_cotiz), "#####0.00")
Else
  t_dolares = Format$(Val(T_TOTAL) * Val(t_cotiz), "#####0.00")
End If
End Sub

Private Sub t_sucursal_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_sucursal_LostFocus()
If t_sucursal = "" Then
  t_sucursal = "0001"
Else
  t_sucursal = Format$(t_sucursal, "0000")
  
End If
End Sub

Private Sub t_total_LostFocus()
Call sacatotales
Call estadoformapago
End Sub

