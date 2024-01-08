VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_prod1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODUCTOS"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8580
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   37
      Top             =   120
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
      Begin VB.Label Label10 
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
      Height          =   8175
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_medida 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5160
         MaxLength       =   20
         TabIndex        =   23
         Top             =   6480
         Width           =   2775
      End
      Begin VB.TextBox t_color 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox t_talle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox t_dtocompra2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5400
         MaxLength       =   13
         TabIndex        =   12
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox c_tasaib 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox t_abreviatura 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   19
         TabIndex        =   25
         ToolTipText     =   "Abreviatura - texto central etiquetas - (max. 19 caracteres)"
         Top             =   7200
         Width           =   2775
      End
      Begin VB.TextBox t_tipocarga 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7440
         MaxLength       =   1
         TabIndex        =   27
         ToolTipText     =   "[M] Manual - [A] Automatica(Solo en Tiques Fiscales)"
         Top             =   7200
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   6000
         Picture         =   "Arch002.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   6000
         Picture         =   "Arch002.frx":0105
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   6000
         Picture         =   "Arch002.frx":020A
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   255
         Left            =   6000
         Picture         =   "Arch002.frx":030F
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox t_observaciones 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   24
         Top             =   6840
         Width           =   5775
      End
      Begin VB.TextBox t_moneda 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   28
         ToolTipText     =   "[P] Pesos - [D] Dolares "
         Top             =   7560
         Width           =   495
      End
      Begin VB.TextBox t_tipo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   26
         ToolTipText     =   "[P] Productos - [M] Materias Primas"
         Top             =   7200
         Width           =   495
      End
      Begin VB.TextBox t_stockminimo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   22
         Top             =   6480
         Width           =   1335
      End
      Begin VB.TextBox t_codbarra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5640
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox c_marca 
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Text            =   "c_marca"
         Top             =   2880
         Width           =   3615
      End
      Begin VB.ComboBox c_depto 
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Text            =   "c_depto"
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox t_final 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   21
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox t_impuesto 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5160
         MaxLength       =   13
         TabIndex        =   20
         ToolTipText     =   "Se calcula sobre el precio de compra"
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox t_tasaimpint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   19
         Top             =   5760
         Width           =   1335
      End
      Begin VB.ComboBox c_iva 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox t_pu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   17
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox t_utilidad 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   15
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox t_costo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   14
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox t_fletecompra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   13
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox t_dtocompra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   11
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox t_preciocompra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   10
         Top             =   3240
         Width           =   1335
      End
      Begin VB.ComboBox c_grupo 
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Text            =   "c_grupo"
         Top             =   2160
         Width           =   3615
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Text            =   "c_prov"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox t_envase 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox c_unidad 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   1
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Medida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   71
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   70
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Talle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   69
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "% Desc. Compra2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   23
         Left            =   3720
         TabIndex        =   68
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Tasa IB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   22
         Left            =   3840
         TabIndex        =   67
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Abreviatura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   21
         Left            =   480
         TabIndex        =   66
         Top             =   7200
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   20
         Left            =   6720
         TabIndex        =   65
         Top             =   7200
         Width           =   615
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   19
         Left            =   480
         TabIndex        =   60
         Top             =   6840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   18
         Left            =   480
         TabIndex        =   59
         Top             =   7560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Tipo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   17
         Left            =   5160
         TabIndex        =   58
         Top             =   7200
         Width           =   735
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Stock Minimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   16
         Left            =   480
         TabIndex        =   57
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cod. Barra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   56
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   15
         Left            =   480
         TabIndex        =   55
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   14
         Left            =   480
         TabIndex        =   54
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Precio Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   480
         TabIndex        =   53
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Impuesto Int."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   3840
         TabIndex        =   52
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Tasa Imp. Int."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   480
         TabIndex        =   51
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Tasa Iva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   480
         TabIndex        =   50
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "PU s/Iva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   480
         TabIndex        =   49
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "% Utilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   480
         TabIndex        =   48
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Costo Real"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   47
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "% Flete Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   480
         TabIndex        =   46
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "% Desc. Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   480
         TabIndex        =   45
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Precio Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   44
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   43
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   42
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Envase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   41
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   40
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "id. Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   30
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "Arch002.frx":0414
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Arch002.frx":0C96
         Style           =   1  'Graphical
         TabIndex        =   31
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
      TabIndex        =   29
      Top             =   8325
      Width           =   11910
      _ExtentX        =   21008
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
            TextSave        =   "29/12/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:25 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "abm_prod1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String



Private Sub btnacepta_Click()
Call graba
End Sub

Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   

   
   'On Error GoTo ERRORGRABA
       
   If IsNull(t_codbarra) Then
     t_codbarra = 0
   End If
   
   If IsNull(t_talle) Then
     t_talle = " "
   End If
   
   If IsNull(t_color) Then
     t_color = " "
   End If
        
   If IsNull(t_medida) Then
     t_medida = " "
   End If
        
  
   Select Case t_funcion
      
   Case "A"


QUERY = "INSERT INTO a2([descripcion], [id_grupo], [id_proveedor], [precio_ult_compra], [fecha_ult_compra], [id_proveedor_ult_compra], [pu], [cod_tasaiva], [id_unidad], [envase], [stock], " & _
" [requeridos], [pedidos], [stock_minimo], [id_marca], [id_departamento], [porc_utilidad], [costoreal], [flete_compra], [dto_compra], [cod_barra], [precio_final], [tasa_imp_interno], " & _
" [tipo_producto], [moneda], [impuesto], [observaciones], [ultima_compra], [ultima_venta], [fecha_actu_precio_venta],  [vigente], [tipo_carga_tique], [texto_central], [id_tasaib], [emite_etiqueta], [dto_compra2], " & _
" [id_prod_prov], [talle], [color],[medida], [dolar_ult_compra], [num_int_ult_compra], [percibe_5329])"

QUERY = QUERY & " VALUES ('" & t_descripcion & "', " & c_grupo.ItemData(c_grupo.ListIndex) & ", " & c_prov.ItemData(c_prov.ListIndex) & ", " & Val(t_preciocompra) & _
", '01/01/2005', 1, " & Val(t_pu) & ", " & c_iva.ItemData(c_iva.ListIndex) & ", " & c_unidad.ItemData(c_unidad.ListIndex) & ", " & Val(t_envase) & ", 0, 0, 0, " & Val(t_stockminimo) & _
", " & c_marca.ItemData(c_marca.ListIndex) & ", " & c_depto.ItemData(c_depto.ListIndex) & ", " & Val(t_utilidad) & ", " & Val(t_costo) & ", " & Val(t_fletecompra) & ", " & _
Val(t_dtocompra) & ", '" & RTrim$(t_codbarra) & "', " & Val(t_final) & ", " & Val(t_tasaimpint) & ", '" & t_tipo & "', '" & t_moneda & "', " & t_impuesto & ", '" & t_observaciones & _
"', 'Sin Compras', 'Sin Ventas', '01/01/2006',  1, '" & t_tipocarga & "', '" & RTrim$(t_abreviatura) & " ', " & c_tasaib.ItemData(c_tasaib.ListIndex) & ", 'S', " & Val(t_dtocompra2) & _
", 0, '" & t_talle & "', '" & t_color & "','" & t_medida & "',1,0,'N')"
       
       
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      
 qr = "SELECT @@IDENTITY AS NewID"
 Set rs = cn1.Execute(qr)
 numintprod = rs.Fields("NewID").Value
     
     MsgBox ("Nuevo Basico de producto generado: " & Format$(numintprod, "00000"))
     
      
   
   Case "M"
   
   QUERY = "update a2 set  [Descripcion]='" & t_descripcion & "' , [id_proveedor]=" & c_prov.ItemData(c_prov.ListIndex) & " , [id_unidad]=" & c_unidad.ItemData(c_unidad.ListIndex) & _
" , [envase]=" & Val(t_envase) & " , [id_grupo]=" & c_grupo.ItemData(c_grupo.ListIndex) & " , [precio_ult_compra]=" & Val(t_preciocompra) & " , [pu]=" & Val(t_pu) & " , [cod_tasaiva]=" & _
c_iva.ItemData(c_iva.ListIndex) & " , [stock_minimo]= " & Val(t_stockminimo) & " , [id_marca]=" & c_marca.ItemData(c_marca.ListIndex) & " , [id_departamento]=" & _
c_depto.ItemData(c_depto.ListIndex) & " , [porc_utilidad]=" & Val(t_utilidad) & " , [costoreal]=" & Val(t_costo) & " , [flete_compra]=" & Val(t_fletecompra) & " , [dto_compra]=" & _
Val(t_dtocompra) & " , [cod_barra]='" & RTrim$(t_codbarra) & "' , [precio_final]=" & Val(t_final) & " , [tasa_imp_interno]=" & Val(t_tasaimpint) & " , [tipo_producto]='" & t_tipo & _
 "' , [moneda]='" & t_moneda & "' , [impuesto]=" & Val(t_impuesto) & " , [observaciones]='" & t_observaciones & "' , [texto_central]='" & RTrim$(t_abreviatura) & " ', [id_tasaib]=" & _
 c_tasaib.ItemData(c_tasaib.ListIndex) & " , [talle]='" & t_talle & "', [color]='" & t_color & "', [medida]='" & t_medida & "'"
 
QUERY = QUERY & " where [id_producto]= " & Val(t_id)
   

      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      
   Case "B"
      
      q = "select * from vta_03 where [id_producto] = " & Val(t_id)
      Set rs = New ADODB.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        MsgBox ("El producto esta ingresado en comprobantes de Venta. No se puede Eliminar")
      Else
        QUERY = "DELETE FROM a2 WHERE [id_producto] = " & Val(t_id)
        cn1.BeginTrans
        cn1.Execute QUERY
        cn1.CommitTrans
      End If
      Set rs = Nothing
   
   End Select
   
   ABM_PROD.DataGrid1.Refresh
   ABM_PROD.Show
   Me.Hide
    
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub



Private Sub c_iva_LostFocus()
Call calcula
End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
  End If
End If

End Sub

Private Sub c_tasaib_LostFocus()
If c_tasaib.ListIndex < 0 Then
  c_tasaib.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
ABM_grupos.Show
End Sub

Private Sub Command1_LostFocus()
Call carga_grupos(c_grupo)

End Sub

Private Sub Command2_Click()
ABM_deptoS.Show

End Sub

Private Sub Command2_LostFocus()
Call carga_deptos_venta(c_depto)
End Sub
Sub deptos()

End Sub

Private Sub Command3_Click()
ABM_marcas.Show
End Sub

Private Sub Command3_LostFocus()
Call carga_marcas(c_marca)
End Sub

Private Sub Command4_Click()
ABM_PROv.Show
End Sub

Private Sub Command4_LostFocus()
Call carga_proveedores(c_prov)
End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  If t_funcion = "A" Then
    t_tipo = "P"
    t_moneda = para.moneda
    t_envase = 1
    Call buscavalor(c_iva, para.tasageneral)
    
  End If
  t_codbarra.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   Case Is = vbKeyF9
     Call graba
         
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 28)
  
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_unidad(c_unidad)
Call carga_proveedores(c_prov)
Call carga_tasaiva(c_iva)
Call carga_grupos(c_grupo)
Call carga_deptos_venta(c_depto)
Call carga_marcas(c_marca)
Call carga_tasaib(c_tasaib)



End Sub


Private Sub t_costo_LostFocus()
Call calcula
End Sub

Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub



Private Sub t_dtocompra_LostFocus()
Call calcula
End Sub

Private Sub t_dtocompra2_LostFocus()
Call calcula
End Sub

Private Sub t_envase_LostFocus()
If Val(t_envase) < 1 Then
  t_envase = 1
End If
End Sub

Private Sub t_final_GotFocus()
If Val(t_final) <= 0 Then
  t_final = ""
End If
End Sub

Private Sub t_final_LostFocus()
t_final = Format$(Val(t_final), "####0.00")
Call calcula3
End Sub

Private Sub t_fletecompra_LostFocus()
Call calcula
End Sub

Private Sub t_impuesto_LostFocus()
Call calcula2
End Sub

Private Sub t_moneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
End If
End Sub

Private Sub t_moneda_LostFocus()
If t_moneda <> "P" And t_moneda <> "D" Then
  t_moneda = para.moneda
End If

End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub

Private Sub t_preciocompra_LostFocus()
Call calcula
End Sub
Sub calcula()
d = Format$(Val(t_preciocompra) * Val(t_dtocompra) / 100, "#####0.000")
n = Val(t_preciocompra) - Val(d)
d2 = Format$(n * Val(t_dtocompra2) / 100, "#####0.000")
n = n - Val(d2)

F = n * Val(t_fletecompra) / 100
n2 = F + n
t_costo = Format$(n2, "#####0.000")
t_pu = Format$(Val(t_costo) + (Val(t_costo) * Val(t_utilidad) / 100), "#####0.000")
t_preciocompra = Format$(Val(t_preciocompra), "#####0.000")
t_dtocompra = Format$(Val(t_dtocompra), "#####0.000")
t_dtocompra2 = Format$(Val(t_dtocompra2), "#####0.000")

t_fletecompra = Format$(Val(t_fletecompra), "#####0.000")


End Sub
Sub calcula2()
t_impuesto = Format$(Val(t_preciocompra) * Val(t_tasaimpint) / 100, "####0.0000")
t_final = Format$(Val(t_pu) + (Val(t_pu) * Val(c_iva) / 100) + Val(t_impuesto), "######0.00")
t_tasaimpint = Format$(Val(t_tasaimpint), "#####0.000")

End Sub
Sub calcula3()
t_pu = Format$(Val(t_final) / (1 + (Val(c_iva) / 100)), "#####0.00")


End Sub
Private Sub t_pu_GotFocus()
If Val(t_pu) <= 0 Then
  t_pu = ""
End If
End Sub

Private Sub t_pu_LostFocus()
If Val(t_pu) <= 0 Then
  Call calcula
Else
  t_pu = Format$(Val(t_pu), "####0.00")
End If
 t_final = Format$(Val(t_pu) * (1 + (Val(c_iva) / 100)), "####0.00")
End Sub

Private Sub t_stockminimo_LostFocus()
If Val(t_stockminimo) <= 0 Then
  t_stockminimo = 1
End If
  
End Sub

Private Sub t_tasaimpint_LostFocus()
Call calcula2
End Sub

Private Sub t_tipo_LostFocus()
If t_tipo <> "P" And t_tipo <> "M" Then
  t_tipo = "P"
End If
End Sub

Private Sub t_tipocarga_LostFocus()
t_tipocarga = Format$(t_tipocarga, ">@")
If t_tipocarga <> "M" And t_tipocarga <> "A" Then
  t_tipocarga = "M"
End If

End Sub

Private Sub t_utilidad_LostFocus()
Call calcula
End Sub
