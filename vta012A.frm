VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_config_comp1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONFIGURA COMPROBANTES VENTA"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPORTANTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   9240
      TabIndex        =   68
      Top             =   1320
      Width           =   2655
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1080
         Picture         =   "vta012A.frx":0000
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   $"vta012A.frx":030A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1935
         Index           =   22
         Left            =   120
         TabIndex        =   69
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "¡Atencion!"
      Height          =   1695
      Left            =   120
      TabIndex        =   60
      Top             =   6600
      Width           =   9015
      Begin VB.Label Label19 
         Caption         =   "Registra Contabilidad(Debe o Haber) se debe establecer para la cuenta madre de la Forma de Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1320
         Width           =   8055
      End
      Begin VB.Label Label17 
         Caption         =   "Los comprobantes que no utilicen letras A, B, C utilizaran el numerador de la letra A"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   960
         Width           =   8775
      End
      Begin VB.Label Label16 
         Caption         =   "Los comprobantes entre 41 y 49 tendran el mismo numero.(Remitos y devlouciones)"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   960
         Width           =   8775
      End
      Begin VB.Label Label13 
         Caption         =   "Los comprobantes que no son PROPIOS tendran los numeradores en cero."
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   720
         Width           =   8775
      End
      Begin VB.Label Label12 
         Caption         =   $"vta012A.frx":039B
         Height          =   495
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   30
      Top             =   120
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   31
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
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   6375
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Width           =   9015
      Begin VB.TextBox t_ie 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   21
         Top             =   5880
         Width           =   375
      End
      Begin VB.TextBox t_sucursal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   70
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_formato2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   20
         Top             =   5880
         Width           =   375
      End
      Begin VB.TextBox t_ib 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   16
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox t_moneda 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   19
         Top             =   5520
         Width           =   375
      End
      Begin VB.TextBox t_items 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   18
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox t_formato 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   17
         Top             =   4800
         Width           =   375
      End
      Begin VB.TextBox t_propio 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox t_contab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   15
         Top             =   4080
         Width           =   375
      End
      Begin VB.TextBox t_venta 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   14
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox t_iva 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   13
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox t_ctacte 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   12
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox t_stock 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   11
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox t_copiasE 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7200
         MaxLength       =   8
         TabIndex        =   10
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox t_copiasC 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5640
         MaxLength       =   8
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox t_copiasB 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   8
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox t_copiasA 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox t_ultimoE 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7200
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox t_ultimoc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5640
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox t_ultimoB 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox t_ultimoA 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox t_abrevia 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   0
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Imprime Desc. Extra"
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
         Left            =   3360
         TabIndex        =   75
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label20 
         Caption         =   "[N] No Imprime   [S] Imprime "
         Height          =   255
         Left            =   5880
         TabIndex        =   74
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "[N] No Imprime   [G] Imprime "
         Height          =   255
         Left            =   2640
         TabIndex        =   72
         Top             =   4800
         Width           =   3375
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Sucursal:"
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
         Left            =   5160
         TabIndex        =   71
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Formato:"
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
         Index           =   21
         Left            =   120
         TabIndex        =   65
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Informe IB"
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
         Index           =   20
         Left            =   120
         TabIndex        =   64
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "[S] Suma   [R] Resta  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   63
         Top             =   4440
         Width           =   3375
      End
      Begin VB.Label Label11 
         Caption         =   "[P] Solo Pesos   [D] Solo Dolares  [A] Ambas"
         Height          =   255
         Left            =   2640
         TabIndex        =   59
         Top             =   5520
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra en Cuenta:"
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
         Left            =   120
         TabIndex        =   58
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cant. Items:"
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
         Index           =   18
         Left            =   120
         TabIndex        =   57
         Top             =   5160
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Tipo Impresion:"
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
         Index           =   17
         Left            =   120
         TabIndex        =   56
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "[S] Si   [N] No (Numerado por Terceros)"
         Height          =   255
         Left            =   2640
         TabIndex        =   55
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Comprobante Propio:"
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
         Height          =   495
         Index           =   16
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "[D] Debe   [H] Haber  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   53
         Top             =   4080
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Contabilidad:"
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
         Left            =   120
         TabIndex        =   52
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "[S] Suma   [R] Resta  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   51
         Top             =   3720
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Informe Vta.:"
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
         Left            =   120
         TabIndex        =   50
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "[S] Suma   [R] Resta  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   49
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "[D] Debe   [H] Haber  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "[E] Entrada   [S] Salida  [N] No Registra"
         Height          =   255
         Left            =   2640
         TabIndex        =   47
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Iva Vta como"
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
         Left            =   120
         TabIndex        =   46
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra CtaCte como:"
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
         Index           =   12
         Left            =   120
         TabIndex        =   45
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Registra Stock como:"
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
         Left            =   120
         TabIndex        =   44
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "E"
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
         Index           =   10
         Left            =   6840
         TabIndex        =   43
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "C"
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
         Index           =   9
         Left            =   5280
         TabIndex        =   42
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "B"
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
         Index           =   8
         Left            =   3720
         TabIndex        =   41
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "A"
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
         Index           =   7
         Left            =   2160
         TabIndex        =   40
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Cantidad Copias"
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
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "E"
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
         Index           =   5
         Left            =   6840
         TabIndex        =   38
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "C"
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
         Index           =   4
         Left            =   5280
         TabIndex        =   37
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "B"
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
         Index           =   2
         Left            =   3720
         TabIndex        =   36
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "A"
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
         Index           =   1
         Left            =   2160
         TabIndex        =   35
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Ultimo Num. Usado"
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
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   1935
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
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Id."
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
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Comprobante"
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
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   23
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "vta012A.frx":0437
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta012A.frx":0CB9
         Style           =   1  'Graphical
         TabIndex        =   24
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
      TabIndex        =   22
      Top             =   8265
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
            TextSave        =   "16/09/2022"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:30 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_config_comp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Private cambio As String


Private Sub btnacepta_Click()
Call graba
End Sub

Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   'On Error GoTo ERRORGRABA
    
   Select Case t_funcion
   
   Case "M"
      cn1.BeginTrans
      If para.numeracion_comun_Fact_nc = "S" And Val(t_id) <= 3 Then
        For i = 1 To 10
         QUERY = "update vta_06 set  [ult_num_A]= " & Val(t_ultimoA) & " , [ult_num_b]= " & Val(t_ultimoB) & " , [ult_num_c]= " & Val(t_ultimoc) & " , [ult_num_e]= " & Val(t_ultimoE)
         QUERY = QUERY & " where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp]= " & i
         
          cn1.Execute QUERY
         
        Next i
      
      End If
      
      If Val(t_id) >= 41 And Val(t_id) <= 49 Then
      
       For i = 41 To 49
         QUERY = "update vta_06 set  [ult_num_A]= " & Val(t_ultimoA) & " , [ult_num_b]= " & Val(t_ultimoB) & " , [ult_num_c]= " & Val(t_ultimoc) & " , [ult_num_e]= " & Val(t_ultimoE)
         QUERY = QUERY & " where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = " & i
          cn1.Execute QUERY
         
       Next i
      End If
      
      QUERY = "update vta_06 set  [Descripcion]='" & t_descripcion & "' , [abreviatura]='" & t_abrevia & "' , [ult_num_A]= " & Val(t_ultimoA) & " , [ult_num_b]= " & Val(t_ultimoB) & _
" , [ult_num_c]= " & Val(t_ultimoc) & " , [ult_num_e]= " & Val(t_ultimoE) & " , [cant_copias_a]= " & Val(t_copiasA) & " , [cant_copias_b]= " & Val(t_copiasB) & _
 " , [cant_copias_c]= " & Val(t_copiasC) & " , [cant_copias_e]= " & Val(t_copiasE) & " , [stock]='" & t_stock & "' , [ctacte]='" & t_ctacte & "' , [iva]='" & t_iva & _
"' , [tipo_impresora]='" & t_formato & "' , [moneda]='" & t_moneda & "' , [venta]='" & t_venta & "' , [contabilidad]='" & t_contab & "' , [propio]='" & t_propio & _
 "' , [cant_lineas]= " & Val(t_items) & " , [ib]='" & t_ib & "' , [formato]='" & t_formato2 & "' , [imprime_desc_extra]='" & t_ie & "'"
 QUERY = QUERY & " where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = " & Val(t_id)
 cn1.Execute QUERY
       
   
   
      'forma linea de texto para informar cambio
       t = t_propio & "-" & t_ultimoA & "-" & t_ultimoB & "-" & t_ultimoc & "-" & t_ultimoE & " [" & t_stock & t_ctacte & t_iva & t_venta & t_contab & t_ib & t_formato & t_items & t_moneda & t_formato2 & "]"
      
      QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
      QUERY = QUERY & " VALUES ('Configura Comp: " & t_id & " sucursal: " & t_sucursal & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', '" & Left$(t, 50) & "', 8, 0)"

      cn1.Execute QUERY
       cn1.CommitTrans
    
   
   End Select
   
   vta_config_comp.DataGrid1.Refresh
   vta_config_comp.Show
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



Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  c_cuenta.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  t_descripcion.SetFocus
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
    Call TabEnter2(Me, 21)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)

End Sub

Private Sub t_comision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If

End Sub

Private Sub t_comision_LostFocus()
If t_comision = "" Then
  t_comision = 0
End If

End Sub

Private Sub t_contab_GotFocus()
cambio = t_contab
End Sub

Private Sub t_contab_LostFocus()
t_contab = Format$(t_contab, ">@")
If t_contab <> "D" And t_contab <> "H" And t_contab <> "N" Then
   t_contab = cambio
End If

End Sub

Private Sub t_ctacte_GotFocus()
cambio = t_ctacte
End Sub

Private Sub t_ctacte_LostFocus()
t_ctacte = Format$(t_ctacte, ">@")
If t_ctacte <> "D" And t_ctacte <> "H" And t_ctacte <> "N" Then
   t_ctacte = cambio
End If

End Sub

Private Sub t_descripcion_LostFocus()
If t_descripcion = "" Then
  t_descripcion = "Null"
End If
End Sub


Private Sub t_formato_LostFocus()
If t_formato <> "" Then
  t_formato = Format$(t_formato, ">@")
Else
  t_formato = "G"
End If
End Sub

Private Sub t_formato2_LostFocus()
If t_formato2 = "" Then
  t_formato2 = "1"
End If
End Sub

Private Sub t_ib_GotFocus()
cambio = t_ib
End Sub

Private Sub t_ib_LostFocus()
t_ib = Format$(t_ib, ">@")
If t_ib <> "S" And t_ib <> "R" And t_ib <> "N" Then
   t_ib = cambio
End If

End Sub

Private Sub t_ie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_ie_LostFocus()
t_ie = Format$(t_ie, ">@")
If t_ie <> "S" And t_ie <> "N" Then
  t_ie = "N"
End If
End Sub

Private Sub t_iva_GotFocus()
cambio = t_iva
End Sub

Private Sub t_iva_LostFocus()
t_iva = Format$(t_iva, ">@")
If t_iva <> "S" And t_iva <> "R" And t_iva <> "N" Then
   t_iva = cambio
End If

End Sub

Private Sub t_moneda_GotFocus()
cambio = t_moneda

End Sub

Private Sub t_moneda_LostFocus()
t_moneda = Format$(t_moneda, ">@")
If t_moneda <> "P" And t_moneda <> "D" And t_moneda <> "A" Then
   t_moneda = cambio
End If
End Sub

Private Sub t_propio_GotFocus()
cambio = t_propio
End Sub

Private Sub t_propio_LostFocus()
t_propio = Format$(t_propio, ">@")
If t_propio <> "S" And t_propio <> "N" Then
  t_propio = cambio
End If

If t_propio = "N" Then
 t_ultimoA = 0
 t_ultimoB = 0
 t_ultimoc = 0
 t_ultimoE = 0
End If

End Sub

Private Sub t_stock_GotFocus()
cambio = t_stock
End Sub

Private Sub t_stock_LostFocus()
t_stock = Format$(t_stock, ">@")
If t_stock <> "E" And t_propio <> "S" And t_propio <> "N" Then
   t_stock = cambio
End If

End Sub

Private Sub t_venta_GotFocus()
cambio = t_venta
End Sub

Private Sub t_venta_LostFocus()
t_venta = Format$(t_venta, ">@")
If t_venta <> "S" And t_venta <> "R" And t_venta <> "N" Then
   t_venta = cambio
End If
End Sub

Private Sub Text1_Change()

End Sub
