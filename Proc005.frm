VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form op 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "ORDEN DE PAGO A PROVEDORES"
   ClientHeight    =   8490
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.TextBox t_cuit 
      Height          =   375
      Left            =   11160
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox c_zona 
      Height          =   315
      Left            =   8760
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculo Manual Ret."
      Height          =   495
      Left            =   10320
      TabIndex        =   56
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales"
      Height          =   1695
      Left            =   120
      TabIndex        =   42
      Top             =   6120
      Width           =   9975
      Begin VB.TextBox t_netoretib 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         MaxLength       =   14
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox c_tiporetg 
         Height          =   315
         Left            =   3360
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox t_neto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox fdolar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox t_total 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8520
         MaxLength       =   14
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_totald 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8520
         MaxLength       =   14
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox t_retib 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8880
         MaxLength       =   14
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox t_alicuotaretib 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8280
         MaxLength       =   5
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox t_pago 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox retencion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         MaxLength       =   14
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Neto sujeto a Ret. IB:"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3360
         TabIndex        =   63
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Neto sujeto a Ret. Gananacias:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Cotiz. U$s de Ajuste:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "%"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   8640
         TabIndex        =   48
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Total a Pagar:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   8
         Left            =   6840
         TabIndex        =   47
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Total U$s"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   46
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ret. IB s/ neto comp. sel."
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   5
         Left            =   6840
         TabIndex        =   45
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Total $ a  Aplicar en  cuenta prov.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ret. Gan. s/  neto comp. seleccionados:"
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   20
         Left            =   3360
         TabIndex        =   43
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   41
      Top             =   240
      Width           =   4215
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   240
      TabIndex        =   35
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   8640
         Picture         =   "Proc005.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   8280
         Picture         =   "Proc005.frx":0105
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1200
         Width           =   255
      End
      Begin VB.ComboBox denominACION 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   6135
      End
      Begin VB.TextBox t_numop 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox sucursal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Nº O.Pago:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retenciones"
      Height          =   1695
      Left            =   9720
      TabIndex        =   26
      Top             =   120
      Width           =   2175
      Begin VB.TextBox T_CALCULARETGAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox t_CALCULARETIB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox T_TASARG 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox t_crg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox t_big 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox t_minimoretib 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Retiene IB/GAN"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo/tasa"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Base Imp.Gan."
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Min.IB"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comprobantes a Aplicar"
      Height          =   1695
      Left            =   240
      TabIndex        =   23
      Top             =   1800
      Width           =   10695
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   1335
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "A los comprobantes que tengan Cod. ret.IB = 0  no se le calculará dicho impuesto"
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   2355
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Forma de Pago"
      Height          =   2535
      Left            =   240
      TabIndex        =   22
      Top             =   3480
      Width           =   11655
      Begin VB.TextBox t_aingresar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox t_diferencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox t_ingresado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid msf2 
         Height          =   1815
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   3201
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "A Ingresar:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   60
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "Diferencia:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   29
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Ingresado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   25
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   10320
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "O.P. U$s"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Confirma 
         Appearance      =   0  'Flat
         Caption         =   "Confirma"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Salir 
         Appearance      =   0  'Flat
         Caption         =   "Salir"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox Detalle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   79
      TabIndex        =   15
      Top             =   7920
      Width           =   5655
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
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
            TextSave        =   "19/12/2017"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:26 a.m."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Zona:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   61
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "En concepto  de:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   7920
      Width           =   1815
   End
End
Attribute VB_Name = "op"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim EXISTE As String
Dim phoy As Double
Dim crg As Integer
Dim pagomes As Double
Dim retmes As Double
Dim impnosujret As Double
Dim rg_alicuota As Double
Dim excedente As Double
Dim gnumintop As Long

'FIXIT: Declare 'cc' and 'ig' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Private Function calcularetgan(ByVal cc, ByVal ig) As Double
If T_CALCULARETGAN = "S" Then
 q = "select * from i_01 where [id_impuesto] = 217"
 Set rs2 = New ADODB.Recordset
 rs2.Open q, cn1
 If Not rs2.EOF And Not rs2.BOF Then
   If rs2("calcula") = "S" Then
      retmin = rs2("retencion-minima")
      impmin = rs2("importe_minimo_sujeto_ret")
      If Val(t_neto) >= impmin Then
        'cc es cod concepto gan ret
        'ig inscripto en ganancias S o N
        'pago del comprobante actual
        'saco pago mes y ret mes
        m = Val(Mid$(t_fecha, 4, 2))
        a = Val(Mid$(t_fecha, 7, 4))
        retmes = 0
        pagomes = 0
        excedente = 0
        insr = 0
        rg_alicuota = 0
        impnosujret = 0

        Set rs = New ADODB.Recordset
        q = "select * from ret_01 where [id_proveedor] = " & denominACION.ItemData(denominACION.ListIndex) & " and [id_retgan] = " & cc & " and [mes] = " & m & " and [año] = " & a
        rs.Open q, cn1
        If Not rs.BOF And Not rs.EOF Then
            retmes = rs("ret_mes")
            pagomes = rs("pagos_mes")
        Else
            retmes = 0
            pagomes = 0
        End If
        Set rs = Nothing


        q = "select * from i_02 where [id_impuesto] = 217 and [id_concepto] = " & cc
        Set rs = New ADODB.Recordset
        rs.Open q, cn1
        If Not rs.BOF And Not rs.EOF Then
           If ig = "S" Then
             insr = rs("importe_noretenido")
             usatabla = rs("porescala_i")
             rg_alicuota = rs("alicuota_i")
           Else
             insr = rs("importe_noretenido_n")
             usatabla = rs("porescala_n")
             rg_alicuota = rs("alicuota_n")
           End If
           conceptoret = rs("concepto")
       Else
           insr = 0
           usatabla = "N"
           rg_alicuota = 0
           conceptoret = "No se detalla"
           MsgBox ("error al generar retencion de ganancias")
       End If
       Set rs = Nothing
       impnosujret = insr
       excedente = pagomes + phoy - insr
       t_ret = 0
       T_TASARG = rg_alicuota
       If usatabla = "N" Then
           t_ret = (excedente * rg_alicuota / 100) - retmes
       Else
           'por tabla
          Set rs = New ADODB.Recordset
        q = "select * from i_03 where [id_impuesto] = 217 order by [secuencia]"
        rs.Open q, cn1, adOpenStatic, adLockReadOnly
        While Not rs.EOF
            If excedente >= rs("minimo") And excedente <= rs("maximo") Then
             'ENCONTRE LA RETENCION
                R1 = rs("importe_retenido")
                R2 = (excedente - rs("sobre_EXcEDENTE")) * (rs("porcentaje_extra") / 100)
                t_ret = R1 + R2 - retmes
                rs.MoveLast
            End If
            rs.MoveNext
        Wend
       End If
    Else
      t_ret = 0
    End If
   Else
    t_ret = 0
   End If
  Else
   t_ret = 0
  End If
Else
  t_ret = 0
End If
calcularetgan = t_ret

End Function


Private Sub carga_comp_pendiente()
op1.armagrid
ic = Space$(10)
id = Space$(10)
n2 = Space$(10)
id = Space$(10)
iss = Space$(10)
QUERY = "select [ctacte], [moneda], [total], [subtotal], [total_d], [letra], [sucursal], [num_comprobante], [fecha], [id_codretgan], [id_tipocomp], [num_int], [id_codretib], [id_cuenta], [cotiz_dolar], [saldo_impago] from a5 where  [id_proveedor] = " & denominACION.ItemData(denominACION.ListIndex) & " and [estado_pago] = 'N' and  [ctacte] <> 'N' and [contado] = 'N' and id_tipocomp <= 30"
QUERY = QUERY & " order by fecha"
Set rs = New ADODB.Recordset
rs.Open QUERY, cn1
While Not rs.EOF
   If rs("ctacte") = "D" Then
     If rs("moneda") = "P" Then
       RSet ic = Format$(-rs("total"), "######0.00")
       RSet n2 = Format$(-rs("subtotal"), "######0.00")
       RSet id = Format$(-rs("total_d"), "######0.00")
       RSet iss = Format$(-rs("saldo_impago"), "######0.00")
     Else
       RSet ic = Format$(-rs("total_d"), "######0.00")
       RSet n2 = Format$(-rs("subtotal") * rs("cotiz_dolar"), "######0.00")
       RSet id = Format$(-rs("total"), "######0.00")
        RSet iss = Format$(-rs("saldo_impago"), "######0.00")
     End If
   Else
     If rs("moneda") = "P" Then
       RSet ic = Format$(rs("total"), "######0.00")
       RSet n2 = Format$(rs("subtotal"), "######0.00")
       RSet id = Format$(rs("total_d"), "######0.00")
       RSet iss = Format$(rs("saldo_impago"), "######0.00")
     Else
       RSet ic = Format$(rs("total_d"), "######0.00")
       RSet n2 = Format$(rs("subtotal") * rs("cotiz_dolar"), "######0.00")
       RSet id = Format$(rs("total"), "######0.00")
       RSet iss = Format$(rs("saldo_impago"), "######0.00")
     End If
   End If
   tc = Format$(rs("letra"), "@")
   sc = Format$(rs("sucursal"), "0000")
   nc = Format$(rs("num_comprobante"), "00000000")
   fc = Format$(rs("fecha"), "dd/mm/yyyy")
   cr = Format$(rs("id_codretgan"), "000")
   cc = "(" & Format$(rs("id_tipocomp"), "000") & ")"
   ni = Format$(rs("num_int"), "00000000")
   crib = Format$(rs("id_codretib"), "00")
   cta = Format$(rs("id_cuenta"), "000000")
   op1.msf1.AddItem "" & Chr$(9) & fc & Chr$(9) & cc & tc & " " & sc & "-" & nc & Chr$(9) & ic & Chr$(9) & cr & Chr$(9) & ni & Chr$(9) & n2 & Chr$(9) & id & Chr$(9) & cta & Chr$(9) & iss & Chr$(9) & Chr$(9) & Chr$(9) & crib
   rs.MoveNext
Wend
Set rs = Nothing
End Sub




Private Sub c_tiporetg_LostFocus()
If c_tiporetg.ListIndex < 0 Then
  c_tiporetg.ListIndex = 0
End If
t_crg = c_tiporetg.ItemData(c_tiporetg.ListIndex)
t_hoy = 0
Call pi50(Val(t_crg))
End Sub

Private Sub c_zona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If c_zona.ListIndex < 0 Then
     c_zona.ListIndex = 0
   End If
   Detalle = RTrim$(Detalle) & " "
   If Val(t_total) <> Val(t_ingresado) Then
     MsgBox ("El importe ngresado No coinciden con el total de la O.P.")
   Else
     Command3.Enabled = True
     confirma.Enabled = True
     confirma.SetFocus
   End If
End If

End Sub

Private Sub Command1_Click()
calcula_ret.c_prov.ListIndex = buscaindice(calcula_ret.c_prov, denominACION.ItemData(denominACION.ListIndex))
Call carga_impuestos(calcula_ret.c_concepto, calcula_ret.c_impuesto.ItemData(calcula_ret.c_impuesto.ListIndex))
calcula_ret.c_concepto.ListIndex = buscaindice(calcula_ret.c_concepto, calcula_ret.c_impuesto.ListIndex)
calcula_ret.t_fecha = t_fecha
calcula_ret.Show
End Sub

Private Sub Command2_Click()
ABM_PROv.Show

End Sub

Private Sub Command3_Click()
op2.Show
End Sub

Private Sub Command5_Click()
com_proveedor.t_id = denominACION.ItemData(denominACION.ListIndex)
com_proveedor.carga
com_proveedor.Show

End Sub

Private Sub Confirma_Click()
 
If verificaperiodog(t_fecha) = "A" Then
 J = MsgBox("Confirma Operacion para Orden de Pago", 4)
 If J = 6 Then
  If verificaperiodo(t_fecha) = "C" Then
   MsgBox ("El periodo para el cual se deseas ingresar el comprobante esta CERRADO!!!!!")
  Else
    q = "select * from a5 where  [num_comprobante] = " & Val(t_numop) & " and [sucursal] = " & Val(sucursal) & " and [id_tipocomp] = 50"
    Set rso = New ADODB.Recordset
    rso.MaxRecords = 1
    rso.Open q, cn1
    If Not rso.BOF And Not rso.EOF Then
      MsgBox ("El Numero de OP ya existe en el sistema")
    Else
      espere.Show
      espere.Label1 = "Espere.... Grabando Orden de Pago"
      espere.Refresh
      EXISTE = "N"
      Call graba
      Unload espere
      t = MsgBox("Imprime Orden de pago", 4)
      If t = 6 Then
        Set cl_comp = New COMPROBANTES
        cl_comp.cargar2 (gnumintop)
        If cl_comp.numint > 0 Then
          cl_comp.imprimir
        End If
        Set cl_comp = Nothing
      End If

    
      If Val(retencion) > 0 Then
        Call grabaretencion
      End If
    
      If Val(t_retib) > 0 Then
        Call grabaretencionib
      End If
 
        End If
    Set rso = Nothing
       
    
    
    Call INICIALIZA2(Me)
    Call armagrid
    Call armagrid2
    t_numop.SetFocus

    
    Call pi3
  End If
 
 End If
Else
 MsgBox ("Periodo cerrado. Imposible grabar operacion")
End If
End Sub

Private Function controlacodret() As Integer
'devuelve EL COD. RET GAN SI SON TODOS IGUALES o 0 si son distintos
k = 1
ccr = 0
b = 0

While k <= msf1.Rows - 1
 If b = 0 Then
    codretgan = Val(msf1.TextMatrix(k, 3))
    ccr = codretgan
    b = 1
 Else
    If codretgan <> Val(msf1.TextMatrix(k, 3)) Then
       ccr = 0
       k = msf1.Rows
    End If
 End If
 k = k + 1
 Wend
 
controlacodret = ccr

End Function


Private Sub denominACION_LostFocus()
If denominACION.ListIndex < 0 Then
  denominACION.ListIndex = 0
End If
 Call inicia
  
  
End Sub
Sub inicia()
espere.Show
espere.Label1 = "Inicializando Comprobante....."
espere.Refresh
Set cl_prov = New proveedores
cl_prov.carga (denominACION.ItemData(denominACION.ListIndex))
If cl_prov.idprov > 0 Then
    t_tipoiva = cl_prov.codtipoiva
    t_minimoretib = para.minimo_retib
    t_CALCULARETIB = cl_prov.calcularetib
    T_CALCULARETGAN = cl_prov.calcularetgan
    t_cuit = cl_prov.CUIT
    Call carga_comp_pendiente
    'para.calcula_ret_ib = "S"
    
    If cl_prov.calcularetib = "S" And para.calcula_ret_ib = "S" Then
     Set cl_padronib = New padron_ib
     cl_padronib.inicia
     periodo = cl_padronib.fecha_desde & "-" & cl_padronib.fecha_hasta
     cl_padronib.cuit_texto = cl_prov.CUIT
     cl_padronib.buscar
     t_alicuotaretib = Format$(cl_padronib.tasa_retib, "##0.00")
     Select Case cl_padronib.estado_consulta
     Case Is = "OK"
       Label20 = "¡COMPROBANTE SUJETO A RETENCION IB! Consulta del Padron de IB Satistactoria"
     Case Is = "NO"
       Label20 = "¡ATENCION! El contribuyente NO se encuentra en el padron, se aplica la tasa Maxima de Ret/Perc. Verifique si corresponde"
     Case Is = "ER"
       Label20 = "¡CUIDADO! Numero de cuit con formato invalido. Padron NO consultado"
     End Select
     Frame11.Visible = True
     Frame11.Caption = "Validez: " & periodo
     
     If cl_padronib.estado_embargo = "OK" Then
        MsgBox ("PROVEEDOR EMBARGADO POR EL ARBA. Fecha Padron de Embargos " & cl_padronib.fecha_embargo)
        Label20 = "¡PROVEEDOR EMBARGADO POR ARBA!"
     End If
     
     Set cl_padronib = Nothing
     
     
   Else
     t_alicuotaretib = "0.00"
     t_retib = "0.00"
     gcuit = "0"
   End If
   Unload espere
Else
    t_tipoiva = 0
    t_cuit = 0
    Unload espere
    MsgBox ("Error en B.D. Proveedores")
    Unload Me
End If

End Sub

Private Sub fdolar_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub fdolar_LostFocus()
If Val(fdolar) <= 1 Then
   fdolar = "1.00"
End If
Call totales2
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 10
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 1900
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 800
msf1.ColWidth(4) = 1000
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 1000
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 1000

msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Comprobante"
msf1.TextMatrix(0, 2) = "Saldo Comp."
msf1.TextMatrix(0, 3) = "Ret.Gan"
msf1.TextMatrix(0, 4) = "Num.Int."
msf1.TextMatrix(0, 5) = "Neto Aplicar"
msf1.TextMatrix(0, 6) = "Total U$s"
msf1.TextMatrix(0, 7) = "Cuenta"
msf1.TextMatrix(0, 8) = "Total Aplicar"
msf1.TextMatrix(0, 9) = "Cod.Ret.IB"

End Sub

Sub armagrid2()
msf2.clear
msf2.Rows = 1
msf2.Cols = 10
msf2.ColWidth(0) = 600
msf2.ColWidth(1) = 1200
msf2.ColWidth(2) = 1200
msf2.ColWidth(3) = 2500
msf2.ColWidth(4) = 1700
msf2.ColWidth(5) = 1700
msf2.ColWidth(6) = 1000
msf2.ColWidth(7) = 1000
msf2.ColWidth(8) = 1000
msf2.ColWidth(9) = 1000

msf2.TextMatrix(0, 0) = "Cod."
msf2.TextMatrix(0, 1) = "Forma Pago"
msf2.TextMatrix(0, 2) = "Num.Cheque"
msf2.TextMatrix(0, 3) = "Detalle/Banco"
msf2.TextMatrix(0, 4) = "Sucursal"
msf2.TextMatrix(0, 5) = "Titular"
msf2.TextMatrix(0, 6) = "Importe"
msf2.TextMatrix(0, 7) = "Fecha Dif."
msf2.TextMatrix(0, 8) = "Num.Int."
msf2.TextMatrix(0, 9) = "Cuenta"


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 16)
  'Case Is = 27
  '      Me.Hide
End Select
End Sub

Sub graba()
 If EXISTE = "N" Then
   'op nueva
      
      'On Error GoTo ERRORGRABA
      numint = saca_ultnumero_int_comp("C")
      't_numop = Format$(saca_ultnumero_comp(50), "00000000")
      gnumintop = numint
      
      Set cl_comp = New COMPROBANTES
      cl_comp.actual (50)
      ctacte = cl_comp.ctacte
      
      'iva guarda datos de alicuota ib
      'subtoital guarda dato de neto sujeto a retenciones
      
      cn1.BeginTrans
      
QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], [iva], [no_grabado], [percep_ret], [total], [fecha_prob_entrega]," & _
"[fecha_recepcion], [estado], [ID_CODRETGAN], [ID_CUENTA], [STOCK], [CTACTE], [grabado], [estado_pago], [num_op], [obs], [ret_ib], [ret_gan], [condiciones], [info_contacto], [moneda], [cotiz_dolar], [contado], " & _
" [TOTAL_D], [monto_suj_ret], [alicuota_ret], [ret_mes], [pagos_realizados], [pago_actual], [minimo_no_imp], [fecha_vto], [zona], [cuit05], [proveedor05])"
QUERY = QUERY & " VALUES (" & numint & ", " & Val(sucursal) & ", " & Val(t_numop) & ", 'O', 50, " & denominACION.ItemData(denominACION.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & ", " & _
 Val(t_neto) & ", " & Val(t_alicuotaretib) & ", 0, 0, " & Val(t_total) & ", '" & t_fecha & "', '" & t_fecha & "', 'A', " & Val(t_crg) & ", 0, " & "'N', '" & ctacte & "', '" & cl_comp.grabado & "', 'X', '0000-00000000', '" & _
 Detalle & " ', " & Val(t_retib) & ", " & Val(retencion) & ", ' ', ' ', 'P', " & Val(fdolar) & ", 'N', " & Val(t_totald) & ", 0, 0, 0, 0, 0, 0, '" & t_fecha & "', " & c_zona.ListIndex + 1 & ", " & Val(t_cuit) & ", '" & Left$(denominACION, 50) & "')"

      cn1.Execute QUERY
      
      'actualiza contador
      QUERY = "update g2 set  [ult_num]=" & Val(t_numop)
      QUERY = QUERY & " where [id_tipo_comp]= 50"
      cn1.Execute QUERY
      
      
      'actualiza pagos
      
      For i = 1 To msf2.Rows - 1
        QUERY = "INSERT INTO a7([num_int], [secuencia], [id_formapago], [formapago], [num_ch], [detalle_banco], [sucursal], [titular], [importe], [fecha_dif], [num_int_fp])"
        QUERY = QUERY & " VALUES (" & numint & ", " & i & ", " & Val(msf2.TextMatrix(i, 0)) & ", '" & Left$(msf2.TextMatrix(i, 1), 20) & "', " & Val(msf2.TextMatrix(i, 2)) & ", '" & Left$(msf2.TextMatrix(i, 3), 25) & "', '" & Left$(msf2.TextMatrix(i, 4), 25) & "', '" & Left$(msf2.TextMatrix(i, 5), 25) & "', " & Val(msf2.TextMatrix(i, 6)) & ", '" & msf2.TextMatrix(i, 7) & "', " & Val(msf2.TextMatrix(i, 8)) & ")"
        cn1.Execute QUERY

        
       If Val(msf2.TextMatrix(i, 0)) = 3 Then 'ch. terceros
        Set rs = New ADODB.Recordset
        q = "select [estado], [destino], [num_int_op], [fecha_salida], [tipo_salida] from cyb_03 where [num_interno] = " & Val(msf2.TextMatrix(i, 8))
        rs.MaxRecords = 1
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
          rs("estado") = "P"
          rs("destino") = Left$(denominACION, 50)
          rs("num_int_op") = numint
          rs("fecha_salida") = t_fecha
          rs("tipo_salida") = "O"
          numintch = Val(msf2.TextMatrix(i, 8))
          rs.Update
         Else
          numintch = 0
        End If
        Set rs = Nothing
       Else
        numintch = 0
       End If
      
      
      
      If Val(msf2.TextMatrix(i, 0)) >= 50 Then 'ch. propios
         
        'emito ch.
        QUERY = "INSERT INTO cyb_04([id_banco], [fecha], [importe], [id_tipomov], [fecha_dif], [ubicacion], [entro], [fecha_acreed], [num_comp], [detalle], [modulo], [num_mov_int], [id_tipodbcr], [num_mov_int_compras])"
        QUERY = QUERY & " VALUES (" & Val(msf2.TextMatrix(i, 0)) & ", '" & t_fecha & "', " & Val(msf2.TextMatrix(i, 6)) & ", 1, '" & msf2.TextMatrix(i, 7) & "', 'D', 'N', '" & t_fecha & "', " & Val(msf2.TextMatrix(i, 2)) & ", 'Op." & Format$(Val(t_numop), "00000000") & " " & Left$(denominACION, 28) & "', 'C', " & numint & ", 1, 0)"
        cn1.Execute QUERY
        
        qr = "SELECT @@IDENTITY AS NewID"
        Set rs = cn1.Execute(qr)
        numintbcox = rs.Fields("NewID").Value

        
        
        Set rs = New ADODB.Recordset
        q = "select [estado], [fecha_dif], [fecha_emision], [destino], [importe], [num_int_op], [num_mov_banco] from cyb_02 where [id_banco] = " & Val(msf2.TextMatrix(i, 0)) & " and [num_cheque] = " & Val(msf2.TextMatrix(i, 2))
        rs.MaxRecords = 1
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
          If rs("estado") = "P" Then
             rs("estado") = "E"
             rs("fecha_dif") = msf2.TextMatrix(i, 7)
             rs("fecha_emision") = t_fecha
             rs("destino") = denominACION
             rs("importe") = Val(msf2.TextMatrix(i, 6))
             rs("num_int_op") = numint
             rs("num_mov_banco") = numintbcox
             rs.Update
             
          Else
             MsgBox ("Error al asignar ch. propio")
          End If
        End If
        Set rs = Nothing
       
       End If
 
 
      If Val(msf2.TextMatrix(i, 0)) = 4 Then 'Transf. o db. Automatico
                
        'emito debito
        QUERY = "INSERT INTO cyb_04([id_banco], [fecha], [importe], [id_tipomov], [fecha_dif], [ubicacion], [entro], [fecha_acreed], [num_comp], [detalle], [modulo], [num_mov_int], [id_tipodbcr])"
        QUERY = QUERY & " VALUES (" & Val(msf2.TextMatrix(i, 5)) & ", '" & msf2.TextMatrix(i, 7) & "', " & Val(msf2.TextMatrix(i, 6)) & ", 90, '" & msf2.TextMatrix(i, 7) & "', 'D', 'N', '" & msf2.TextMatrix(i, 7) & "', " & Val(msf2.TextMatrix(i, 2)) & ", 'Op." & Format$(Val(t_numop), "00000000") & " " & Left$(denominACION, 28) & "', 'C', " & numint & ", 1)"
        cn1.Execute QUERY
        
       
       End If
 
 
      
      'la cuenta concepto de caja
      If msf1.Rows > 1 Then
            'utilizo la cuenta del primer comprobante
            cuentacontra = Val(msf1.TextMatrix(1, 7))
      Else
            'utilizo la cuenta de compras varias
      '      cuentacontra = para.cuenta_compras_varias
             cuentacontra = para.cuenta_acreedores
      
      End If

      
      
      Set rs = New ADODB.Recordset
      q = "select [caja] from cyb_01 where [id_forma_pago] = " & Val(msf2.TextMatrix(i, 0))
      rs.MaxRecords = 1
      rs.Open q, cn1
      If Not rs.BOF And Not rs.EOF Then
        If rs("caja") = "S" Then
                    
          'grabo mov caja
           QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
           QUERY = QUERY & " VALUES (" & Val(msf2.TextMatrix(i, 9)) & ", " & cuentacontra & ", '" & Left$(denominACION, 49) & " ', " & Val(msf2.TextMatrix(i, 6)) & ", 'H', '" & t_fecha & "', " & numint & ", 'C', 'O.P. " & Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000") & "', " & Val(msf2.TextMatrix(i, 0)) & ", " & numintch & ", " & para.id_usuario & ")"
           cn1.Execute QUERY
        End If
      End If
      Set rs = Nothing
      
      Next i
     
      
      'actualiza comprobantes aplicados
      
      For i = 1 To msf1.Rows - 1
        Set rs = New ADODB.Recordset
        q = "select [num_int], [estado_pago], [num_op], [saldo_impago], [id_tipocomp] from a5 where [num_int] = " & Val(msf1.TextMatrix(i, 4))
        rs.MaxRecords = 1
        rs.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
          If rs("id_tipocomp") = 30 Then
            ssi = rs("saldo_impago") + Val(msf1.TextMatrix(i, 8))
          Else
             ssi = rs("saldo_impago") - Val(msf1.TextMatrix(i, 8))
          End If
          rs("saldo_impago") = ssi
          If ssi <= 0.1 Then
            rs("estado_pago") = "P"
          End If
          rs("num_op") = Format$(Val(sucursal), "0000") & "-" & Format$(t_numop, "00000000")
          rs.Update
        
          Set rs2 = New ADODB.Recordset
          q = "select * from a15"
          rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
          rs2.AddNew
            rs2("num_int_comp") = rs("num_int")
            rs2("num_int_op") = numint
            rs2("importe_pagado") = Val(msf1.TextMatrix(i, 8))
            rs2("saldo_comprobante") = Val(msf1.TextMatrix(i, 2))
          rs2.Update
          Set rs2 = Nothing
            
        
        End If
        Set rs = Nothing
      Next i
        
      
      'actualizo acumulador retencion gan.
      Set rs = New ADODB.Recordset
      q = "select * from ret_01 where [id_proveedor] = " & denominACION.ItemData(denominACION.ListIndex) & " and [id_retgan] = " & Val(t_crg) & " and [mes] = " & Val(Mid$(t_fecha, 4, 2)) & " and [año] = " & Val(Mid$(t_fecha, 7, 4))
      rs.MaxRecords = 1
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      If Not rs.BOF And Not rs.EOF Then
         
      Else
        rs.AddNew
        rs("id_proveedor") = denominACION.ItemData(denominACION.ListIndex)
        rs("id_retgan") = Val(t_crg)
        rs("mes") = Val(Mid$(t_fecha, 4, 2))
        rs("año") = Val(Mid$(t_fecha, 7, 4))
      End If
     
      rs("ret_mes") = rs("ret_mes") + Val(retencion)
      rs("pagos_mes") = rs("pagos_mes") + Val(t_big)

      rs.Update

      
      'contabilidad
      If Generaasientosauto Then
       If cl_comp.contabilidad <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")
         cta = para.cuenta_acreedores
         u1 = cl_comp.contabilidad
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         Set rs = New ADODB.Recordset
         q = "select [descripcion] from c_01 where [id_cuenta] = " & cta
         rs.MaxRecords = 1
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Pagos] " & cl_comp.abreviatura & " " & Format$(Val(sucursal), "0000") & "-" & Format$(Val(t_numop), "00000000") & "', 'C', " & numint & ", " & Val(t_total) & ", " & Val(t_total) & ", " & para.id_usuario & ", '" & Left$(RTrim$(denominACION), 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre acreedores
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & cta & ", '" & u1 & "', " & Val(t_total) & ", 'Op. Nro." & Format$(Val(sucursal), "0000") & "-" & Format$(Val(t_numop), "00000000") & "')"
         cn1.Execute QUERY
      
         'formas de pago
         ic = 2
         For i = 1 To msf2.Rows - 1
              d = Left$(RTrim$(msf2.TextMatrix(i, 3)), 35) & " " & msf2.TextMatrix(i, 2)
              cta = Val(msf2.TextMatrix(i, 9))
              QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
              QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u2 & "', " & Val(msf2.TextMatrix(i, 6)) & ", '" & d & "')"
              cn1.Execute QUERY
              ic = ic + 1
         Next i
       End If
      End If
      
      cn1.CommitTrans
        
Else
   MsgBox ("No se puede modificar O.P.")
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  


End Sub

Private Sub Form_Load()
denominACION.clear
Call carga_proveedores(denominACION)
denominACION.RemoveItem 0
denominACION.ListIndex = 0
Call carga_impuestos(c_tiporetg, 217)
c_tiporetg.ListIndex = 0
c_tiporetg.Visible = False
Load op1
Load op2
Load op_fp3
Load op_fp2
Load op_fp1
Load calcula_ret
Call carga_zonas(c_zona)
c_zona.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload op1
Unload op2
Unload op_fp3
Unload op_fp2
Unload op_fp1
Unload calcula_ret
End Sub

Sub grabaretencionib()
      'On Error GoTo errorretib
      numint = saca_ultnumero_int_comp("C")
      numret = Format$(saca_ultnumero_comp(96), "00000000")
      
      Set cl_comp = New COMPROBANTES
      cl_comp.actual (96)
      ctacte = cl_comp.ctacte
      
      nop = Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000")
      trd = Format$(Val(op2.t_retib), "#####0.00")
      cn1.BeginTrans
      QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], [iva], [no_grabado], [percep_ret], [total], [fecha_prob_entrega], " & _
      " [fecha_recepcion], [estado], [ID_CODRETGAN], [ID_CUENTA], [STOCK], [CTACTE], [grabado], [estado_pago], [num_op], [obs], [ret_ib], [ret_gan], [condiciones], [info_contacto], [moneda], [cotiz_dolar], " & _
      " [contado], [total_d], [monto_suj_ret], [alicuota_ret], [ret_mes], [pagos_realizados], [pago_actual], [minimo_no_imp], [saldo_impago], [zona], [cuit05], [proveedor05])"
      QUERY = QUERY & " VALUES (" & numint & ", " & Val(sucursal) & ", " & numret & ", 'X', 96, " & denominACION.ItemData(denominACION.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & ", 0, 0, 0, " & _
      Val(t_retib) & ", " & Val(t_retib) & ", '" & t_fecha & "', '" & t_fecha & "', 'A', 0, 0, " & "'N', '" & ctacte & "', '" & cl_comp.grabado & "', 'X', '" & Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000") & _
      "', '" & "Op. " & Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000") & "', " & Val(t_retib) & ", 0, ' ', ' ', 'P', " & Val(fdolar) & ", 'N', " & Val(trd) & ", " & Val(t_netoretib) & ", " & Val(t_alicuotaretib) & _
      ", 0, 0, 0, 0, 0, " & c_zona.ListIndex + 1 & ", " & Val(t_cuit) & ", '" & Left$(denominACION, 50) & "')"
      
     
      cn1.Execute QUERY
      
      
      For i = 1 To msf1.Rows - 1
        If msf1.TextMatrix(i, 9) <> "00" Then
         QUERY = "INSERT INTO a16([num_int_ret], [num_int_comp], [importe_imponible], [num_comprobante])"
         QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 4)) & ", " & Val(msf1.TextMatrix(i, 5)) & ", '" & Left$(msf1.TextMatrix(i, 1), 25) & "')"
         cn1.Execute QUERY
        End If
      Next i
      
      
     If Generaasientosauto Then
      If cl_comp.contabilidad <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")
         cta = para.cuenta_acreedores
         u1 = cl_comp.contabilidad
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         Set rs = New ADODB.Recordset
         q = "select [descripcion] from c_01 where [id_cuenta] = " & cta
         rs.MaxRecords = 1
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Compras] " & cl_comp.abreviatura & " " & Format$(Val(sucursal), "0000") & "-" & Format$(numret, "00000000") & "', 'C', " & numint & ", " & Val(t_retib) & ", " & Val(t_retib) & ", " & para.id_usuario & ", '" & Left$(RTrim$(denominACION), 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre acreedores
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & cta & ", '" & u1 & "', " & Val(t_retib) & ", 'Op. Nro." & nop & "')"
         cn1.Execute QUERY
      
         'cta ret
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 2, " & para.cuenta_retib & ", '" & u2 & "', " & Val(t_retib) & ", 'Op. Nro." & nop & "')"
         cn1.Execute QUERY
       End If
     End If
     cn1.CommitTrans
      
            
      J = MsgBox("Imprime Comrobante de Retencion Ing. Brutos", 4)
      If J = 6 Then
        Call imprimeretib(numint)
      End If
             
      Exit Sub
      
errorretib:
   MsgBox ("Error al generar Retencion IB")
   cn1.RollbackTrans
   Exit Sub

End Sub

Private Sub grabaretencion()
     ' On Error GoTo errorret
      numint = saca_ultnumero_int_comp("C")
      numret = Format$(saca_ultnumero_comp(95), "00000000")
      
      Set cl_comp = New COMPROBANTES
      cl_comp.actual (95)
      ctacte = cl_comp.ctacte
      
      nop = Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000")
      trd = Format$(Val(op2.t_retgan), "#####0.00")
      cn1.BeginTrans
      
      QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], [iva], [no_grabado], [percep_ret], [total], [fecha_prob_entrega], [fecha_recepcion], [estado], [ID_CODRETGAN], [ID_CUENTA], [STOCK], [CTACTE], [grabado], [estado_pago], [num_op], [obs], [ret_ib], [ret_gan], [condiciones], [info_contacto], [moneda], [cotiz_dolar], " & _
      "[contado], [total_d], [monto_suj_ret], [alicuota_ret], [ret_mes], [pagos_realizados], [pago_actual], [minimo_no_imp], [saldo_impago], [zona], [cuit05], [proveedor05])"
      QUERY = QUERY & " VALUES (" & numint & ", " & Val(sucursal) & ", " & numret & ", 'X', 95, " & denominACION.ItemData(denominACION.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & ", 0, 0, 0, " & Val(retencion) & ", " & Val(retencion) & ", '" & t_fecha & "', '" & t_fecha & "', 'A', " & Val(t_crg) & ", 0, " & "'N', '" & ctacte & "', '" & cl_comp.grabado & "', 'X', '" & nop & "', '" & "Op. " & _
      nop & "', 0, " & Val(retencion) & ", '" & Mid$(t_fecha, 4, 2) & "/" & Mid$(t_fecha, 7, 4) & "', ' ', 'P', " & Val(fdolar) & ", 'N', " & Val(trd) & ", " & excedente & ", " & Val(T_TASARG) & ", " & retmes & ", " & pagomes & ", " & phoy & ", " & impnosujret & ", 0, " & c_zona.ListIndex + 1 & ", " & Val(t_cuit) & ", '" & Left$(denominACION, 50) & "')"
     ' MsgBox (QUERY)
      cn1.Execute QUERY
      
      
      If Generaasientosauto Then
       If cl_comp.contabilidad <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")
         cta = para.cuenta_acreedores
         u1 = cl_comp.contabilidad
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
         Set rs = New ADODB.Recordset
         q = "select [descripcion] from c_01 where [id_cuenta] = " & cta
         rs.MaxRecords = 1
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Compras] " & cl_comp.abreviatura & " " & Format$(Val(sucursal), "0000") & "-" & Format$(numret, "00000000") & "', 'C', " & numint & ", " & Val(retencion) & ", " & Val(retencion) & ", " & para.id_usuario & ", '" & Left$(RTrim$(denominACION), 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre acreedores
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & cta & ", '" & u1 & "', " & Val(retencion) & ", 'Op. Nro." & nop & "')"
         cn1.Execute QUERY
      
         'cta ret
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 2, " & para.cuenta_retgan & ", '" & u2 & "', " & Val(retencion) & ", 'Op. Nro." & nop & "')"
         cn1.Execute QUERY
       End If
      End If
      
      cn1.CommitTrans
      
      
      J = MsgBox("Imprime Comrobante de Retencion Ganancia", 4)
      If J = 6 Then
        'Call IMPRIMERETG(numint)
        Set cl_comp = New COMPROBANTES
        cl_comp.cargar2 (numint)
        If cl_comp.numint > 0 Then
            cl_comp.imprimir
         End If
         Set cl_comp = Nothing
        End If
             
      Exit Sub
      
errorret:
   MsgBox ("Error al generar Retencion")
   cn1.RollbackTrans
   Exit Sub
     

End Sub



Private Sub IMPRIMERETG(ByVal n As Double)
'imprime retencion de ganancia
'n es el num-int


Dim gf1 As String
Set rs = New ADODB.Recordset
q = "select * from a5, a1, g1 where [num_int] = " & n & " and a5.[id_proveedor] = a1.[id_proveedor] and a5.[id_usuario] = g1.[id_usuario]"
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
     Set rs2 = New ADODB.Recordset
     q = "select [copias] from g2 where [id_tipo_comp] = " & rs("id_tipocomp")
     rs2.MaxRecords = 1
     rs2.Open q, cn1
     If Not rs2.BOF And Not rs2.EOF Then
        copias = rs2("copias")
     Else
        copias = 1
     End If
     Set rs2 = Nothing
     
     gf1 = "Courier New"
     ip = Space$(9)
     ret = Space$(9)
     c2 = InputBox("Cantidad de Copias", "Imprimir Retencion Ganancias", copias)
     If Val(c2) > 0 Then
      For h = 1 To Val(c2)
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Comprobante Retencion Ganancia - R.G.(AFIP) 830"
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Regimen de Retencion del IMPUESTO A LAS GANANCIAS"
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(50); "Nro.: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(3); Format$(rs("SUCURSAL"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(50); "Fecha:";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(3); rs("fecha")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Agente de Retencion:"; glo.nombrecli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Direccion..........:"; glo.direccioncli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "C.u.i.t............:"; glo.CUIT
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
 
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Contribuyente: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("denominacion"); "  ("; Format$(rs("a5.id_proveedor"), "00000"); ")    "
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Direccion....: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("direccion") & "  -" & rs("localidad") & "-   (" & rs("cp") & ")"
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "CUIT.........: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("cuit")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "N", 10)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Regimen: " & crg
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       RSet ip = Format$(pagomes, "####0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Pagos a la fecha   : " & ip
       RSet ip = Format$(phoy, "####0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Pago a Efectuar    : " & ip
       RSet ip = Format$(impnosujret, "####0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Minimo No Imponible: " & ip
       RSet ip = Format$(excedente, "####0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Monto Suj.a Reten. : " & ip
       RSet ip = Format$(Val(T_TASARG), "####0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Alicuota           : " & ip
       RSet ip = Format$(retmes, "####0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Ret. Anteriores    : " & ip
       RSet ip = Format$(rs("total"), "####0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "RETENCION          : " & ip
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Son pesos: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print convierte(rs("total"))
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Ingresa en DDJJ periodo: "; Mid$(rs("fecha"), 4, 2) & "/" & Mid$(rs("fecha"), 7, 4)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "N", 8)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Orden de Pago"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Fecha         Numero                  Importe      "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print t_fecha & "  " & Format$(sucursal, "0000") & "-" & Format$(t_numop, "00000000") & "             " & Format$(Val(t_total), "######0.00")
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "   ________________________                     ___________________________"
       Call definefont(gf1, "N", 9)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "       Por " & glo.nombrecli; Tab(80); " Por "; rs("denominacion")
        
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.NewPage
     Next h
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.EndDoc
  End If
End If
Set rs = Nothing

End Sub

Private Sub imprimeretib(ByVal n As Double)
'imprime retencion de iB
'n es el num-int
Set cl_comp = New COMPROBANTES
cl_comp.cargar2 (n)
cl_comp.imprimir
Set cl_comp = Nothing
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If


If KeyCode = vbKeyInsert Then
   op1.Show
End If

End Sub
Sub empieza()
t_pago = Format$(suma_msflexgrid(Me.msf1, 8), "######0.00")
t_neto = Format$(suma_msflexgrid(Me.msf1, 5), "######0.00")
t_totald = Format$(suma_msflexgrid(Me.msf1, 6), "######0.00")
If Val(t_totald) > 0 Then
   fdolar = Format$(Val(t_pago) / Val(t_totald), "####0.000")
Else
   fdolar = Format$(para.cotizacion, "####0.000")
End If
t_big = t_neto
End Sub
Sub pasaaplicacion()
  phoy = 0
  crg = controlacodret
  t_crg = crg
  If crg <> 0 Then
    c_tiporetg.Visible = False
    Call pi50(crg)
  Else
    c_tiporetg.Visible = True
    c_tiporetg.ListIndex = 0
    retencion = "0.00"
  End If
  
End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If cl_prov.inscriptogan <> "E" Then
  crg = controlacodret
  t_pago = Format$(suma_msflexgrid(Me.msf1, 2), "######0.00")
  If crg <> 0 Then
    Call pi50(crg)
  Else
    t_retencion = 0
  End If
 Else
   t_retencion = 0
 End If
 total = Format$(Val(t_pago) - Val(retencion))
 msf2.SetFocus
End If
End Sub

Private Sub msf1_LostFocus()
 Call empieza
 Call pasaaplicacion
 Call cargabaseimpib
 Call calcularetib
 Call totales2
 
  msf1.FocusRect = flexFocusLight
End Sub
Sub cargabaseimpib()
k = 1
netoretib = 0
      While k <= msf1.Rows - 1
       'If msf1.TextMatrix(k, 9) <> "00" Then
          netoretib = netoretib + Val(msf1.TextMatrix(k, 5))
       'End If
       k = k + 1
      Wend
 t_netoretib = Format$(netoretib, "######0.00")
End Sub
Private Sub msf2_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F1]Ch.Terc. - [F2]Ch.Propios - [F3]Otras - [F4] Transf/Db autom. - [ENTER] Continua - [F5] Saca Pago "
If msf2.Rows > 0 Then
  msf2.FocusRect = flexFocusNone
Else
  msf2.FocusRect = flexFocusLight
End If
t_ingresado = Format$(suma_msflexgrid(msf2, 6), "######0.00")
t_diferencia = Format$(Val(t_total) - Val(t_ingresado), "######0.00")


End Sub

Private Sub msf2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  Load op_fp3
  op_fp3.t_modulo = "O"
  op_fp3.Show
End If

If KeyCode = vbKeyF2 Then
 Load op_fp2
 op_fp2.t_modulo = "O"
  op_fp2.Show
End If

If KeyCode = vbKeyF1 Then
  Load op_fp1
  op_fp1.t_modulo = "O"
  op_fp1.Show
End If

If KeyCode = vbKeyF4 Then
  Load op_fp5
  op_fp5.t_modulo = "O"
  op_fp5.Show
End If

If KeyCode = vbKeyF9 Then
  If Val(t_total) <> Val(t_ingresado) Then
     MsgBox ("El total ingresado no coincide con el Total a Pagar")
  End If
  t_pago.SetFocus
End If


If KeyCode = vbKeyF5 Then
 If msf2.Rows > 2 Then
    msf2.RemoveItem (msf2.Row)
 Else
   Call armagrid2
 End If
End If

End Sub

Private Sub msf2_LostFocus()
Call barra(Me)
t_ingresado = suma_msflexgrid(msf2, 6)
msf2.FocusRect = flexFocusLight
End Sub




Private Sub pi3()
   Call INICIALIZA2(Me)
   sucursal = Format$(glo.sucursal, "0000")
   t_fecha = Format$(Now, "dd/mm/yyyy")
   Call armagrid
   Call armagrid2
   confirma.Enabled = False
   Command3.Enabled = False
   t_numop = Format$(saca_ultnumero_comp2(50), "00000000")

   
   
End Sub
Sub calcularetib()
t_retib = "0.00"
If t_CALCULARETIB = "S" Then
 q = "select * from i_01 where [id_impuesto] = 50"
 Set rs2 = New ADODB.Recordset
 rs2.Open q, cn1
 
 If Not rs2.EOF And Not rs2.BOF Then
   If rs2("calcula") = "S" Then
      retmin = rs2("retencion-minima")
      impmin = rs2("importe_minimo_sujeto_ret")
       
      If Val(t_netoretib) >= impmin Then
         tr = Format$(Val(t_netoretib) * Val(t_alicuotaretib) / 100, "####0.00")
         t_retib = Format$(tr, "######0.00")
      End If
      
   End If
 End If
 Set rs2 = Nothing
End If
End Sub
Private Sub pi50(ByVal c As Integer)
'calcula retencion ganancias
codretgan = c
'calculo de pago actual
phoy = Val(t_neto)




'calculo la retencion
r = calcularetgan(codretgan, cl_prov.inscriptogan)


Set rs = New ADODB.Recordset
q = " SELECT * FROM i_01 where [id_impuesto] = 217"
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
   If r > rs("retencion-minima") Then
      retencion = Format$(r, "#####0.00")
   Else
      retencion = "0.00"
   End If
Else
   retencion = "0.00"
End If



      


End Sub



Private Sub retencion_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)
End Sub

Private Sub retencion_LostFocus()
If Val(retencion) > 0 And Val(t_crg) > 0 Then
  Call totales2
Else
  If Val(retencion) > 0 Then
    MsgBox ("El concepto de la retencion no esta definido")
    c_tiporetg.SetFocus
  Else
    Call totales2
  End If
End If

End Sub

Private Sub Salir_Click()
   
   Unload Me

End Sub

Private Sub sucursal_GotFocus()
 Call pi3
End Sub

Private Sub sucursal_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_alicuotaretib_LostFocus()
If Val(t_alicuotaretib) < 0 Then
   t_alicuotaretib = "0.00"
End If
Call calcularetib
Call totales2
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
   t_numop.SetFocus
   t_fecha = ""
End If



End Sub

Private Sub t_neto_LostFocus()
If Val(t_neto) < 0 Then
   t_neto = "0.00"
End If
t_big = t_neto
If t_crg <> 0 Then
   Call pi50(crg)
Else
    retencion = "0.00"
End If
Call calcularetib
Call totales2
End Sub

Private Sub t_numop_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Unload Me
End If
End Sub

Private Sub t_numop_LostFocus()
  If t_numop = "" Then
       EXISTE = "N"
       fdolar = para.cotizacion
  Else
      q = "select [num_int] from a5 where [sucursal] = " & Val(sucursal) & " and [num_comprobante] = " & Val(t_numop) & " and [id_tipocomp] = 50"
      Set rs = New ADODB.Recordset
      rs.MaxRecords = 1
      rs.Open q, cn1
      If Not rs.BOF And Not rs.EOF Then
         MsgBox ("O.P. Existente")
         EXISTE = "S"
      Else
         EXISTE = "N"
      End If
      Set rs = Nothing
  End If

End Sub

Private Sub t_pago_KeyPress(KeyAscii As Integer)
  
  Call solonum(KeyAscii, 1)
End Sub



Private Sub t_pago_LostFocus()
Call totales2
End Sub

Private Sub t_retib_LostFocus()
Call totales2
End Sub

Private Sub t_total_LostFocus()
Call fm(t_total)

End Sub

Private Sub t_totald_LostFocus()
   Call fm(t_totald)
   
End Sub


Sub totales2()
t_pago = Format$(t_pago, "######0.00")
t_total = Format$(Val(t_pago) - Val(retencion) - Val(t_retib), "######0.00")
t_aingresar = t_total
If Val(fdolar) < 1 Then
  fdolar = "1.00"
End If
t_totald = Format$(Val(t_total) / Val(fdolar), "######0.00")
t_diferencia = Format$(Val(t_total) - Val(t_ingresado), "######0.00")
op2.t_cotiz = fdolar
op2.t_op = Format$(Val(t_totald), "######0.00")
op2.t_retgan = Format$(Val(retencion) / Val(fdolar), "######0.00")
op2.t_retib = Format$(Val(t_retib) / Val(fdolar), "######0.00")
op2.t_total = Format$(Val(op2.t_op) + Val(op2.t_retgan) + Val(op2.t_retib), "######0.00")
End Sub

