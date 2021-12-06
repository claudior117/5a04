VERSION 5.00
Object = "{0A6BE9FC-5039-11D5-98EC-0800460222F0}#1.0#0"; "IFEpson.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio_vta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO VENTAS"
   ClientHeight    =   8190
   ClientLeft      =   90
   ClientTop       =   90
   ClientWidth     =   12285
   FontTransparent =   0   'False
   Icon            =   "inicio_vta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8190
   ScaleWidth      =   12285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      Caption         =   "FACTURA ELECTRONICA"
      Height          =   1455
      Left            =   120
      TabIndex        =   32
      Top             =   1440
      Width           =   3015
      Begin VB.CommandButton b_probarconexion 
         Caption         =   "Probar Conexion"
         Height          =   1095
         Left            =   1560
         Picture         =   "inicio_vta.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton b_facte 
         Height          =   1095
         Left            =   120
         Picture         =   "inicio_vta.frx":0B8C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "VENTAS POR TERCEROS"
      Height          =   1215
      Left            =   8880
      TabIndex        =   29
      Top             =   1560
      Width           =   3255
      Begin MSComctlLib.Toolbar Toolbar8 
         Height          =   870
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1535
         ButtonWidth     =   2170
         ButtonHeight    =   1429
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "VTA. DIRECTA"
               Key             =   "B1"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "LIQ. CEREAL"
               Key             =   "B2"
               ImageIndex      =   5
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ACCESO A CAJA"
      Height          =   975
      Left            =   3360
      TabIndex        =   27
      Top             =   2760
      Width           =   5175
      Begin MSComctlLib.Toolbar Toolbar6 
         Height          =   645
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1138
         ButtonWidth     =   2037
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList6"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Caja Diaria"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Cartera de Cheques de Tercero"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Object.ToolTipText     =   "Informe de Resultados"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Description     =   "Informe Caja"
               Object.ToolTipText     =   "Informe Caja"
               ImageIndex      =   4
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresora Actual del Sistema"
      Height          =   735
      Left            =   4920
      TabIndex        =   23
      Top             =   7080
      Width           =   4815
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   4080
         Picture         =   "inicio_vta.frx":1AE6
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label7"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EMISION DE COMPROBANTES FISCALES GEN1"
      Height          =   1095
      Left            =   3360
      TabIndex        =   21
      Top             =   3960
      Width           =   8175
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   645
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1138
         ButtonWidth     =   1879
         ButtonHeight    =   1032
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         ImageList       =   "ImageList5"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Description     =   "Emision de Tique Fiscal "
               Object.ToolTipText     =   "Tique Fiscal a Consumidor Final maximo $ 1000.00 "
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Description     =   "Tique Factura Fiscal"
               Object.ToolTipText     =   "Tique Factura Fiscal maximo establecido por el controlador"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Description     =   "Remitos Fiscales"
               Object.ToolTipText     =   "Remitos fiscales solo disponibles en algunos controladores fiscales"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Description     =   "Recibos Fiscales"
               Object.ToolTipText     =   "Recibo fiscal solo disponible en algunos controladores"
               ImageIndex      =   4
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   645
         Left            =   4680
         TabIndex        =   26
         Top             =   240
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   1138
         ButtonWidth     =   1879
         ButtonHeight    =   1032
         Appearance      =   1
         ImageList       =   "ImageList5"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Description     =   "Cierre X"
               Object.ToolTipText     =   "Cierre X"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Description     =   "Cierre Z"
               Object.ToolTipText     =   "Cierre Z"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Description     =   "Verifica Estado de la Impresora Fiscal"
               Object.ToolTipText     =   "Verifica Estado de la Impresora Fiscal"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CONSULTAS RAPIDAS CLIENTES"
      Height          =   1095
      Left            =   3360
      TabIndex        =   19
      Top             =   1560
      Width           =   5415
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   750
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1323
         ButtonWidth     =   2117
         ButtonHeight    =   1217
         Appearance      =   1
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Estado Cuenta Clientes"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Saldos Clientes"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Object.ToolTipText     =   "Listado de Iva ventas"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Object.ToolTipText     =   "Informe de Fletes"
               ImageIndex      =   4
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modulo "
      Height          =   1455
      Left            =   6360
      TabIndex        =   17
      Top             =   5640
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "inicio_vta.frx":1DF0
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ARCHIVOS MAESTROS"
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3015
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   870
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   1535
         ButtonWidth     =   2275
         ButtonHeight    =   1429
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "CLIENTES"
               Key             =   "B1"
               Description     =   "Archivo de Clientes"
               Object.ToolTipText     =   "Archivo de Clientes"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "LISTA PRECIOS"
               Key             =   "B2"
               Description     =   "Listado de Productos y Precios"
               Object.ToolTipText     =   "Lista de Productos y Precios"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EMISION DE COMPROBANTES"
      Height          =   1215
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   8775
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   885
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   1561
         ButtonWidth     =   2328
         ButtonHeight    =   1455
         Appearance      =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "FACTURACION"
               Key             =   "B1"
               Description     =   "Emision de Facturas y ND Fiscales"
               Object.ToolTipText     =   "Emision de Facturas y ND Fiscales"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "RECIBOS"
               Key             =   "B2"
               Description     =   "Recibos por Cobranzas"
               Object.ToolTipText     =   "Recibos por Cobranzas"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "REMITOS"
               Key             =   "B3"
               Description     =   "REMITOS"
               Object.ToolTipText     =   "Emision de Remitos y Notas Dev."
               ImageIndex      =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "COMP. VARIOS"
               Key             =   "B4"
               Description     =   "Comprobantes Varios Manuales "
               Object.ToolTipText     =   "Comprobantes Varios Manuales"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "RETENC."
               Key             =   "B5"
               Description     =   "Carga Retenciones"
               Object.ToolTipText     =   "Carga de retenciones"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Tique NF"
               Key             =   "B6"
               Object.ToolTipText     =   "Facturacion de Fletes"
               ImageIndex      =   10
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   6120
      Width           =   4575
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "CUIT:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Telefono:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Direccion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Razon Social:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9960
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "inicio_vta.frx":20FA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "inicio_vta.frx":297C
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   7935
      Width           =   12285
      _ExtentX        =   21669
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
            TextSave        =   "03/12/2021"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:44 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":31FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":36C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":3BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":4089
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":4519
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":49B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":4E5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":52E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":5740
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":5BCD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":687F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":6B99
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":6EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":80EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":8409
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":88EB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   -600
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   60
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":8E09
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":9437
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":9A8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":A09D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":A6FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   840
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   73
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":AD27
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":AF91
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":B1DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":B3E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   1920
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":B643
            Key             =   "B2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":BD8D
            Key             =   "B1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":C42A
            Key             =   "B3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":CB5A
            Key             =   "B4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":D242
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":D797
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":DB3C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList6 
      Left            =   1920
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   70
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":E78E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":EEDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":F664
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":FDFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio_vta.frx":10596
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal epson4 
      Left            =   11040
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      TabIndex        =   31
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Menu M_tablas 
      Caption         =   "&Tablas"
      Begin VB.Menu M_clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu M_vend 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu M_transp 
         Caption         =   "Transportes"
         Begin VB.Menu M_transpa 
            Caption         =   "Empresas Transportes"
         End
         Begin VB.Menu M_camiones 
            Caption         =   "Camiones/Unidades"
         End
         Begin VB.Menu M_contyrolviajes 
            Caption         =   "Control de Viajes"
         End
      End
      Begin VB.Menu Mretvd 
         Caption         =   "Retenciones para ventas directas"
      End
   End
   Begin VB.Menu M_comprobantes 
      Caption         =   "Comprobantes"
      Begin VB.Menu M_fact 
         Caption         =   "Facturacion (F/NC/ND)"
      End
      Begin VB.Menu M_rto 
         Caption         =   "Remitos"
      End
      Begin VB.Menu M_rbo 
         Caption         =   "Recibos"
      End
      Begin VB.Menu M_pto 
         Caption         =   "Presupuestos"
      End
      Begin VB.Menu M_ret 
         Caption         =   "Retenciones"
      End
      Begin VB.Menu M_ff 
         Caption         =   "Factura de Fletes"
      End
      Begin VB.Menu M_vd 
         Caption         =   "Venta Directa(por terceros)"
      End
      Begin VB.Menu M_vtater 
         Caption         =   "Venta por Terceros"
      End
      Begin VB.Menu M_lq 
         Caption         =   "Liquidacion Cereales"
      End
      Begin VB.Menu m_otros 
         Caption         =   "Otros Comprobantes"
      End
   End
   Begin VB.Menu M_consultas 
      Caption         =   "Consultas"
      Begin VB.Menu M_saldos 
         Caption         =   "Saldos y Estados de Cuenta"
         Begin VB.Menu M_estadocuenta 
            Caption         =   "Estado Cuenta Cliente"
         End
         Begin VB.Menu M_saldoscli 
            Caption         =   "Saldos Clientes"
         End
         Begin VB.Menu M_dual 
            Caption         =   "Estado de Cuenta Dual: Cliente-Proveedor"
         End
         Begin VB.Menu M_actudeuda 
            Caption         =   "Actualizacion de Deuda(Calculo Interes por Mora)"
         End
      End
      Begin VB.Menu M_infcomp 
         Caption         =   "Informes de Comprobantes"
         Begin VB.Menu M_compemitidos 
            Caption         =   "Comprobantes Emitidos"
         End
         Begin VB.Menu M_remitos 
            Caption         =   "Remitos Emitidos"
         End
      End
      Begin VB.Menu M_imp 
         Caption         =   "Informes Impositivos"
         Begin VB.Menu M_ivaventas 
            Caption         =   "Informe de Iva Ventas"
         End
         Begin VB.Menu M_ib 
            Caption         =   "Informe de Ingresos Brutos"
         End
         Begin VB.Menu M_retperc 
            Caption         =   "Informe de Ret. y Perc. recibidas"
         End
         Begin VB.Menu M_retypercvtas 
            Caption         =   "Percepciones Realizadas"
         End
         Begin VB.Menu M_posicioniva 
            Caption         =   "Posicion frente al IVA"
         End
         Begin VB.Menu M_citi 
            Caption         =   "Citi Ventas"
         End
         Begin VB.Menu M_arbacorralones 
            Caption         =   "AICYC Informe Arba para empresas constructoras y corralones"
         End
      End
      Begin VB.Menu M_infventas 
         Caption         =   "Informes de Ventas"
         Begin VB.Menu M_vtaimp 
            Caption         =   "Importes"
            Begin VB.Menu M_cierreventa 
               Caption         =   "Cierre diario Ventas (Importes)"
            End
            Begin VB.Menu M_ventadet 
               Caption         =   "Informe de Ventas Detallado por comprobantes(Importes)"
            End
         End
         Begin VB.Menu M_vtaunid 
            Caption         =   "Unidades"
            Begin VB.Menu M_infvtaun 
               Caption         =   "Informe de Ventas acumulado por Producto(Unidades)"
            End
            Begin VB.Menu M_movproid 
               Caption         =   "Informe de Movimientos de un Producto(Unidades)"
            End
            Begin VB.Menu M_pendientes 
               Caption         =   "Informe de productos pendientes de facturacion(Unidades)"
            End
         End
         Begin VB.Menu M_histyroico 
            Caption         =   "Historico de Ventas por Producto"
         End
         Begin VB.Menu m_infvtacum 
            Caption         =   "Historico de Ventas por Acumulados"
         End
         Begin VB.Menu m_fletes 
            Caption         =   "Fletes"
         End
      End
      Begin VB.Menu M_gerencial1 
         Caption         =   "Informe Resultado"
      End
      Begin VB.Menu M_venc 
         Caption         =   "Vencimientos"
      End
      Begin VB.Menu M_stock 
         Caption         =   "Stock"
         Begin VB.Menu M_stockcli 
            Caption         =   "Movimientos de Stock por Clientes y productos"
         End
         Begin VB.Menu M_stockcli2 
            Caption         =   "Stock  Cliente acumulado por producto"
         End
      End
   End
   Begin VB.Menu m_utiles 
      Caption         =   "&Utiles"
      Begin VB.Menu M_reindexastock 
         Caption         =   "Reindexa Base datos Productos"
      End
      Begin VB.Menu M_confcomp 
         Caption         =   "Configura Comprobantes"
      End
      Begin VB.Menu M_sucventa 
         Caption         =   "Habilitar Sucursal de Venta"
      End
      Begin VB.Menu M_tools 
         Caption         =   "Utilitarios Varios(Tools)"
      End
      Begin VB.Menu M_corrigecuit 
         Caption         =   "Corrige Cuits"
      End
      Begin VB.Menu M_elimina2 
         Caption         =   "Elimina Movimientos por Lotes"
      End
      Begin VB.Menu M_actuprecioprov 
         Caption         =   "Actualiza precios desde Excel Proveedor"
      End
      Begin VB.Menu M_imprtoex 
         Caption         =   "Importa Remitos desde Excel"
      End
      Begin VB.Menu M_importarlista 
         Caption         =   "Importar Lista precios de otro sistema GestionE"
      End
      Begin VB.Menu M_renueva 
         Caption         =   "Renueva datos listado de Iva ventas"
      End
      Begin VB.Menu M_cambiomemfiscal 
         Caption         =   "Cambio Memoria Fiscal(Pasa comp. de un punto de Venta a otro)"
      End
   End
   Begin VB.Menu M_facte 
      Caption         =   "&Factura Electronica"
      Begin VB.Menu M_duplicado 
         Caption         =   "Importar de Duplicado  Electronico (R.G.1361)"
      End
   End
   Begin VB.Menu M_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "inicio_vta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

 'nuevo codigo driver IF Universal
     
       
  
       

Private Sub b_facte_Click()
 vta_facturacion.Show
 vta_facturacion.t_sucursal = Format$(glo.sucursale, "0000")
 vta_facturacion.c_sucursal.ListIndex = buscaindice(vta_facturacion.c_sucursal, glo.sucursale)
 vta_facturacion.c_sucursal.Enabled = False
End Sub

Private Sub b_probarconexion_Click()

If fe_valida_tique() Then
    MsgBox ("Tique Wsaa validado correctamente. Paso (1/2)")
Else
    If fe_genera_wsaa() Then
       MsgBox ("Tique Wsaa generado correctamente. Paso (1/2)")
    Else
       MsgBox ("Error al generar el tique wsaa, verifique el archivo exepciones en la carpeta log")
    End If
End If
 
If fe_valida_wsfe() Then
   MsgBox ("Servidor Wsfe validado correctamente. Paso (2/2). CONEXION EXITOSA")
Else
  MsgBox ("Error al conectar al servidor wsfe, verifique el archivo exepciones en la carpeta log")
End If
   
 
End Sub

Private Sub btnsale_Click()
inicio.Show
Unload Me
End Sub


Function conectarFiscalGen2()
    Dim retorno As Long
    Dim Port As Integer
    Dim myPort As String
        
     Set cl_fiscal = New fiscal
     cl_fiscal.carga (glo.sucursalf)

    'Call SetData(cmbProtocolo.ListIndex, cmbEquipo.ListIndex) 'Opcional solo para efectos demostrativos en el ejemplo


    Do While (True)

        'retorno = FP.ConfigurarVelocidad(115200)
        'If Not (retorno = ERROR_NONE) Then
        '    MsgBox ("Error al conectar impresora fiscal")
        '    Exit Do
        'End If

        'retorno = FP.ConfigurarPuerto(cl_fiscal.puerto)
        'If Not (retorno = ERROR_NONE) Then
        '    MsgBox ("Error al conectar impresora fiscal")
        '    Exit Do
        'End If

        'retorno = FP.ConfigurarProtocolo(cmbProtocolo.ListIndex)
        'If Not (retorno = ERROR_NONE) Then
        '    Exit Do
        'End If

        retorno = FP.NewConectar()
        If Not (retorno = ERROR_NONE) Then
            MsgBox ("Error al conectar impresora fiscal")
            Exit Do
        End If

        Exit Do
    Loop

    ConectarFiscal = retorno
End Function





Private Sub Command1_Click()
F = DateValue("07/04/2019") + 10
MsgBox ("hola" & F)
End Sub

Private Sub Command5_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Form_Activate()
Call barraesag(Me)
If para.fiscal = 0 Then
   Frame3.Visible = False
Else
   Set cl_fiscal = New fiscal
   cl_fiscal.carga (glo.sucursalf)
   If cl_fiscal.id > 0 Then
      cMODELO = cl_fiscal.idmodelo
      cPUERTO = cl_fiscal.puerto
      cBAUDIOS = cl_fiscal.baudios
      Set fiscal = New Driver
      Frame3.Visible = True
   Else
      MsgBox ("Impresora Fiscal No definida")
   End If
End If

If glo.sucursale = 0 Then
   Frame9.Visible = False
Else
   Frame9.Visible = True
   Call actu_fe
End If

Select Case para.HABILITACION

Case Is = 1912
               'fiscal sin ventas normales
               Frame6.Visible = True
               Frame4.Visible = False
               Frame8.Visible = True
               Frame11.Visible = False
               Frame10.Visible = False
               'Frame3.Visible = True
               
               
Case Is = 7422 'agropecuarias
    Frame11.Visible = True
    
Case Is = 1723 'veterinarias
    

Case Is = 1721 'électrinica sin boton facturacion en pc que factura
    'Frame3.Visible = True
    Frame11.Visible = False
    Toolbar4.Buttons(2).Visible = False
    Toolbar4.Buttons(3).Visible = False
    Toolbar4.Buttons(4).Visible = False
    
    Toolbar3.Buttons(4).Visible = False
    
    If glo.sucursale <> 0 Then
       Toolbar2.Buttons(1).Enabled = False
    End If
    
    Toolbar5.Visible = False

Case Is = 1722 'électrinica
    'Frame3.Visible = True
    Frame11.Visible = False
    Toolbar4.Buttons(2).Visible = False
    Toolbar4.Buttons(3).Visible = False
    Toolbar4.Buttons(4).Visible = False
    
    Toolbar3.Buttons(4).Visible = False
    
    
    Toolbar5.Visible = False
    
Case Is = 9999 'todos
    Frame11.Visible = True
    
Case Else
    Frame11.Visible = False
    
End Select
    
Label7 = para.impresora_actual

End Sub
Sub actu_fe()
  q = "select * from fe_01 where id= 1"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
    para.facte_token = rs("token")
    para.facte_sign = rs("sign")
    para.facte_expira = rs("fecha_expira")
    para.facte_certificado = rs("certificado")
    para.facte_claveprivada = rs("clave_privada")
    para.facte_servidor_wsaa = rs("servidor_wsaa")
    para.facte_servidor_wsfe = rs("servidor_wsfe")
    
                 
  End If
  Set rs = Nothing
  
  
  
  

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If

End Sub

Private Sub Form_Load()

Dim retorno As Long
Dim cmd As String

Call titulos(Me)


Select Case para.HABILITACION
Case Is = 1825 'facturacion fletes
     Toolbar2.Buttons(6).Visible = True
     m_fletes.Visible = True
     'Frame9.Visible = False
Case Is = 1723 'veterinaria
  Frame9.Visible = True

Case Is = 9999
     Toolbar2.Buttons(6).Visible = True
     m_fletes.Visible = True
   ' Frame9.Visible = True
Case Else
  
  m_fletes.Visible = False
  Toolbar2.Buttons(6).Visible = False

End Select


q = "SELECT * FROM G1 WHERE id_usuario = " & para.id_usuario
Set rs = New ADODB.Recordset
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  para.tipoprecioventa = rs("tipo_precio_venta")
End If
Set rs = Nothing


Exit Sub

e1:
  MsgBox ("Error al Inicializar Parametros INICIO.LOAD")
  End

End Sub




Private Sub M_actudeuda_Click()
vta_int_mora.Show
End Sub

Private Sub M_actuprecioprov_Click()
vta_actu_listaprov.Show
End Sub

Private Sub M_arbacorralones_Click()
vta_arba_corralones.Show
End Sub

Private Sub M_cambiomemfiscal_Click()
J = InputBox$("Ingrese calve de Administrador General", "Cambio Memoria Fiscal")
If J = "0969" Then
   MsgBox ("Recurde hacer backup antes de realizar esta operacion")
   gen_cambiamemoriafiscal.Show
End If

End Sub

Private Sub M_camiones_Click()
GEN_ABMCAMION.Show
End Sub

Private Sub M_cierreventa_Click()
 vta_informevta4.Show
End Sub

Private Sub M_citi_Click()
vta_citi.Show
End Sub

Private Sub M_clientes_Click()

  vta_ABM_cli.Show

End Sub

Private Sub M_compemitidos_Click()
If para.id_grupo_modulo_actual >= 2 Then
  vta_vercomp.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_confcomp_Click()
  

vta_config_comp.Show
End Sub

Private Sub M_contyrolviajes_Click()
vta_verviajes.Show
End Sub

Private Sub M_corrigecuit_Click()
Call verificacuits

End Sub

Sub verificacuits()
h = MsgBox("Formatea(numerico sin guines de longitud 11) los cuit de los Clientes   . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_01"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("id_tipoiva") <> 3 And rs("id_tipoiva") <> 8 Then
       'lleva cuit
       If Len(rs("cuit")) = 13 Then
          c = Mid$(rs("cuit"), 1, 2) & Mid$(rs("cuit"), 4, 8) & Mid$(rs("cuit"), 13, 1)
          rs("cuit") = c
          rs.Update
       End If
    Else
      If Val(rs("cuit")) <= 0 Then
        rs("cuit") = 0
      End If
    End If
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

h = MsgBox("verifca Cuit en Comprobates   . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_02, vta_01 where vta_02.[id_cliente] > 1 and vta_02.[id_cliente] = vta_01.[id_cliente]"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("cuit02") = rs("cuit")
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If




End Sub

Private Sub M_dual_Click()
vta_estadocuentadual.Show
End Sub

Private Sub M_duplicado_Click()
vta_facte1.Show
End Sub

Private Sub M_elimina2_Click()
J = InputBox$("Ingrese calve de Administrador General", "Eliminar Comprobantes por lotes")
If J = "1975" Then
   gen_borra.Show
End If

End Sub

Private Sub M_estadocuenta_Click()
If para.id_grupo_modulo_actual >= 2 Then
  vta_estadocuenta.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_fact_Click()
 If para.id_grupo_modulo_actual >= 4 Then
     vta_facturacion.Show
     vta_facturacion.t_sucursal = Format$(glo.sucursal, "0000")
     vta_facturacion.c_sucursal.ListIndex = buscaindice(vta_facturacion.c_sucursal, glo.sucursal)
    
    Else
     Call sinpermisos
    End If
End Sub

Private Sub M_ff_Click()
 If para.id_grupo_modulo_actual >= 5 Then
      vta_fact_viaje.Show
    Else
      Call sinpermisos
    End If
End Sub

Private Sub m_fletes_Click()
vta_verfletes.Show
End Sub

Private Sub M_gerencial1_Click()
 If para.id_grupo_modulo_actual >= 7 Then
   vta_gerencial1.Show
 Else
   Call sinpermisos
 End If
End Sub

Private Sub M_histyroico_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual >= 5 Then
  con_HISTORICOcompras.Show
  con_HISTORICOcompras.Check1 = 0
Else
  Call sinpermisos
End If

End Sub

Private Sub M_ib_Click()
vta_ib.Show
End Sub

Private Sub M_importarlista_Click()
vta_importalistaprecios.Show
End Sub

Private Sub M_imprtoex_Click()
vta_importa_rto.Show
End Sub

Private Sub m_infvtacum_Click()
vta_informevta5.Show
End Sub

Private Sub M_infvtaun_Click()
vta_informevta2.Show
End Sub

Private Sub M_ivaventas_Click()
vta_ivaventas.Show
End Sub

Private Sub M_lq_Click()
If para.id_grupo_modulo_actual >= 4 Then
      vta_liqcereal.Show
      vta_liqcereal.t_sucursal = Format$(glo.sucursal, "0000")
      vta_liqcereal.c_sucursal.ListIndex = buscaindice(vta_liqcereal.c_sucursal, glo.sucursal)
    
    Else
      Call sinpermisos
    End If
End Sub

Private Sub M_movproid_Click()
vta_movprodcli.Show
End Sub

Private Sub m_otros_Click()
If para.id_grupo_modulo_actual >= 4 Then
      vta_COMPVARIOS.Show
    Else
      Call sinpermisos
    End If
    
End Sub

Private Sub M_pendientes_Click()
vta_informevta3.Show
End Sub

Private Sub M_posicioniva_Click()
gen_posicioniva.Show
End Sub


Private Sub M_pto_Click()
vta_presup.Show
End Sub

Private Sub M_rbo_Click()
If para.id_grupo_modulo_actual >= 6 Then
       vta_recibo.Show
       vta_recibo.sucursal = Format$(glo.sucursal, "0000")
       vta_recibo.c_sucursal.ListIndex = buscaindice(vta_recibo.c_sucursal, glo.sucursal)
       
    Else
       Call sinpermisos
    End If
End Sub

Private Sub M_reindexastock_Click()
Set cl_stock = New STOCK
cl_stock.actualizastock
Set cl_stock = Nothing

End Sub

Private Sub M_remitos_Click()
vta_verremitos.Show
End Sub

Private Sub M_renueva_Click()
p = InputBox$("Ingrese periodo mmaaaa", "Renovacion datops clientes en listado de iva", "012013")
b = 1
If Len(p) <> 6 Then
  b = 0
  MsgBox ("Verificar el formato mmaaaa para el periodod")
Else
  m = Mid$(p, 1, 2)
  If Val(m) < 1 Or Val(m) > 12 Then
     b = 0
     MsgBox ("El valor para el mes debe estar entre 1-12")
  End If
  a = Mid$(p, 3, 4)
  If Val(a) < 2000 Or Val(m) > 2050 Then
     b = 0
     MsgBox ("Verifique el balor ingresado para el año")
  End If
End If
If b = 1 Then
   'actualiza
     f1 = "01/" & Mid$(p, 1, 2) & "/" & Mid$(p, 3, 4)
    Select Case Val(Mid$(p, 1, 2))
    Case Is = 1, Is = 3, Is = 5, Is = 7, Is = 8, Is = 10, Is = 12
       d = 31
    Case Is = 2, Is = 4, Is = 6, Is = 9, Is = 11
        d = 30
    Case Else
        d = 28
    End Select
    f2 = d & "/" & Mid$(p, 1, 2) & "/" & Mid$(p, 3, 4)
 

    Set rs = New ADODB.Recordset
    espere.Show
    espere.Label1 = "Espere...... Actualizando Listado de Iva"
    espere.Refresh
    q = "select * from VTA_02, vta_01 where  vta_02.[id_cliente] > 1 and  vta_02.[id_cliente] = vta_01.[id_cliente] "
    c = " and "
    q = q & c & " datevalue([fecha]) >= datevalue('" & f1 & "') "
    q = q & c & " datevalue([fecha]) <= datevalue('" & f2 & "') "
    'MsgBox (q)
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    i = 0
    While Not rs.EOF
        i = i + 1
        espere.Label1 = "Espere...... Actualizando Listado de Iva " & i
        espere.Label1.Refresh
        rs("cliente02") = rs("denominacion")
        rs("cuit02") = rs("cuit")
        rs("id_tipo_iva02") = rs("id_tipoiva")
        rs("direccion02") = rs("direccion")
        rs("localidad02") = rs("localidad")
        rs.Update
        rs.MoveNext
    Wend
    Unload espere
    Set rs = Nothing

End If


End Sub
Sub actuiva()



End Sub
Private Sub M_ret_Click()
 
    If para.id_grupo_modulo_actual >= 4 Then
      vta_retenciones.Show
    Else
      Call sinpermisos
    End If
End Sub

Private Sub M_retperc_Click()
vta_retyperc.Show
End Sub

Private Sub M_retypercvtas_Click()
vta_retyperc_realizadas.Show

End Sub

Private Sub M_rto_Click()
If para.id_grupo_modulo_actual >= 4 Then
      vta_remitos.Show
      vta_remitos.t_sucursal = Format$(glo.sucursal, "0000")
      vta_remitos.c_sucursal.ListIndex = buscaindice(vta_remitos.c_sucursal, glo.sucursal)
    
    Else
      Call sinpermisos
    End If
End Sub

Private Sub M_saldoscli_Click()
vta_saldoscli.Show
End Sub

Private Sub M_salir_Click()
inicio.Show
Unload Me
End Sub


Private Sub M_stockcli_Click()
vta_stockcli.Show

End Sub

Private Sub M_stockcli2_Click()
vta_stockcli2.Show

End Sub

Private Sub M_sucventa_Click()
J = InputBox$("Ingrese Password Administrativa")
prueba = "N"
If J = para.password_adm Then
  s = InputBox("Ingrese numero de sucursal a habilitar. CUIDADO si ingresa una sucursal existente se regneraran todos los parametros")
  If Val(s) > 0 Then
   If Val(s) <> glo.sucursal Then
     Set rs = New ADODB.Recordset
     q = "select * from vta_06 where [sucursal] = " & Val(s)
     rs.Open q, cn1, adOpenDynamic, adLockOptimistic
     While Not rs.EOF
        rs.Delete
        rs.MoveNext
     Wend
     Set rs = Nothing
     
    
    Set rs = New ADODB.Recordset
    q = "select * from vta_06 where [sucursal] = " & glo.sucursal
    rs.Open q, cn1
    While Not rs.EOF
      Set rs1 = New ADODB.Recordset
      q = "select * from vta_06 where [sucursal] = " & Val(s) & " and [id_tipocomp] = " & rs("id_tipocomp")
      rs1.Open q, cn1, adOpenStatic, adLockOptimistic
      If rs1.EOF And rs1.BOF Then
         rs1.AddNew
          rs1("sucursal") = Val(s)
          rs1("id_tipocomp") = rs("id_tipocomp")
          rs1("descripcion") = rs("descripcion")
          rs1("abreviatura") = rs("abreviatura")
          rs1("ult_num_a") = 0
          rs1("ult_num_b") = 0
          rs1("ult_num_c") = 0
          rs1("stock") = rs("stock")
          rs1("ctacte") = rs("ctacte")
          rs1("iva") = rs("iva")
          rs1("tipo_impresora") = rs("tipo_impresora")
          rs1("cant_lineas") = rs("cant_lineas")
          rs1("cant_copias_a") = rs("cant_copias_a")
          rs1("moneda") = rs("moneda")
          rs1("cant_copias_b") = rs("cant_copias_b")
          rs1("cant_copias_c") = rs("cant_copias_c")
          rs1("venta") = rs("venta")
          rs1("contabilidad") = rs("contabilidad")
          rs1("ib") = rs("ib")
          rs1("ult_num_e") = 0
          rs1("cant_copias_e") = rs("cant_copias_e")
          rs1("propio") = rs("propio")
          rs1("formato") = rs("formato")
          rs1("imprime_desc_extra") = rs("imprime_desc_extra")
          'rs1("ubicacion06") = rs("ubicacion06")
        rs1.Update
      
      End If
      Set rs1 = Nothing
      rs.MoveNext
    Wend
    Set rs = Nothing
  Else
   MsgBox ("Imposble volver a generar sucursal principal")
  End If
 End If
End If
End Sub

Private Sub M_tools_Click()
gen_tools.Show
End Sub

Private Sub M_transpa_Click()
ABM_PROv.Show
End Sub

Private Sub M_vd_Click()
If para.id_grupo_modulo_actual >= 4 Then
      vta_directa.Show
      vta_directa.t_sucursal = Format$(glo.sucursal, "0000")
      vta_directa.c_sucursal.ListIndex = buscaindice(vta_directa.c_sucursal, glo.sucursal)
    
    Else
      Call sinpermisos
    End If

End Sub

Private Sub M_venc_Click()
vta_vencimientos.Show
End Sub

Private Sub M_vend_Click()

vta_ABM_vend.Show
End Sub

Private Sub M_ventadet_Click()
vta_informevta.Show
End Sub

Private Sub M_vtater_Click()
vta_porterceros.Show
End Sub

Private Sub Mretvd_Click()
ABM_perc.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    vta_ABM_cli.Show

  Case Is = "B2"
    vta_listaprecios.Show

End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    If para.id_grupo_modulo_actual >= 4 Then
     vta_facturacion.Show
     vta_facturacion.t_sucursal = Format$(glo.sucursal, "0000")
     vta_facturacion.c_sucursal.ListIndex = buscaindice(vta_facturacion.c_sucursal, glo.sucursal)
     vta_facturacion.t_cae = "0"
     vta_facturacion.t_cae_vence = "01/01/2000"
    
    Else
     Call sinpermisos
    End If

  Case Is = "B2"
    If para.id_grupo_modulo_actual >= 6 Then
       vta_recibo.Show
       vta_recibo.sucursal = Format$(glo.sucursal, "0000")
       vta_recibo.c_sucursal.ListIndex = buscaindice(vta_recibo.c_sucursal, glo.sucursal)
       
    Else
       Call sinpermisos
    End If

  Case Is = "B3"
    If para.id_grupo_modulo_actual >= 4 Then
      vta_remitos.Show
      vta_remitos.t_sucursal = Format$(glo.sucursal, "0000")
      vta_remitos.c_sucursal.ListIndex = buscaindice(vta_remitos.c_sucursal, glo.sucursal)
    
    Else
      Call sinpermisos
    End If

  Case Is = "B4"
     If para.id_grupo_modulo_actual >= 4 Then
      vta_COMPVARIOS.Show
    Else
      Call sinpermisos
    End If
    
  Case Is = "B5"
    If para.id_grupo_modulo_actual >= 4 Then
      vta_retenciones.Show
    Else
      Call sinpermisos
    End If

Case Is = "B6"
    If para.id_grupo_modulo_actual >= 5 Then
      fsc_tiqueNF.Show
    Else
      Call sinpermisos
    End If


End Select


End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    If para.id_grupo_modulo_actual >= 3 Then
     vta_estadocuenta.Show
    Else
     Call sinpermisos
    End If

  Case Is = "B2"
    If para.id_grupo_modulo_actual >= 3 Then
       vta_saldoscli.Show
    Else
       Call sinpermisos
    End If
  Case Is = "B3"
    If para.id_grupo_modulo_actual >= 5 Then
       vta_ivaventas.Show
    Else
       Call sinpermisos
    End If
 Case Is = "B4"
    If para.id_grupo_modulo_actual >= 5 Then
       vta_verfletes.Show
    Else
       Call sinpermisos
    End If

End Select
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim impf As String
'If glo.sucursalf > 0 Then
 Set cl_fiscal = New fiscal
 cl_fiscal.carga (glo.sucursalf)
 If cl_fiscal.id > 0 Then
  Select Case Button.Key
  Case Is = "B1"
   If para.id_grupo_modulo_actual >= 4 Then
     If cl_fiscal.imprimetique = "S" Then
       impf = cl_fiscal.impresora
       fsc_tique.Show
       fsc_tique.t_impfiscal = impf
     Else
       MsgBox ("La Impresora Fiscal Definida no Imprime Tique")
         
         
     
     End If
    Else
     Call sinpermisos
    End If
    

  Case Is = "B2"
    If para.id_grupo_modulo_actual >= 4 Then
     If cl_fiscal.imprimetf = "S" Or cl_fiscal.imprimefact = "S" Then
          vta_facturacion.Show
          vta_facturacion.t_sucursal = Format$(glo.sucursalf, "0000")
          vta_facturacion.c_sucursal.ListIndex = buscaindice(vta_facturacion.c_sucursal, glo.sucursalf)
          vta_facturacion.t_cae = "0"
          vta_facturacion.t_cae_vence = "01/01/2000"
    Else
         MsgBox ("Impresora Fiscal no emite facturas")
     End If
    Else
     Call sinpermisos
    End If
  Case Is = "B3"
    If para.id_grupo_modulo_actual >= 6 Then
        vta_remitos.Show
        vta_remitos.t_sucursal = Format$(glo.sucursalf, "0000")
        vta_remitos.c_sucursal.ListIndex = buscaindice(vta_remitos.c_sucursal, glo.sucursalf)
    Else
     Call sinpermisos
    End If
    
  Case Is = "B4"
    If para.id_grupo_modulo_actual >= 6 Then
       vta_recibo.Show
       vta_recibo.sucursal = Format$(glo.sucursalf, "0000")
       vta_recibo.c_sucursal.ListIndex = buscaindice(vta_recibo.c_sucursal, glo.sucursalf)
    Else
     Call sinpermisos
    End If
    
End Select

Else
 MsgBox ("Error en la definicion de la Impresora Fiscal")
End If
Set cl_fiscal = Nothing

'Else
'  MsgBox ("Impresora Fiscal No inicializada")
'End If

End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
If glo.sucursalf > 0 Then
 Select Case Button.Key
  Case Is = "B1"
   
    
    J = MsgBox("Confirma la Emision del Cierre X", 4)
    If J = 6 Then
       espere.Show
       espere.Refresh
       espere.Label1 = "Espere.... Emitiendo Cierre X"
      
      'nuevo codigo driver IF Universal
       'Dim Fiscal As Driver
       Set fiscal = New Driver
  
       fiscal.Modelo = cMODELO
       fiscal.puerto = cPUERTO
       fiscal.baudios = cBAUDIOS
  
       
  
        If fiscal.Inicializar Then
  
            fiscal.CancelarComprobante
            If fiscal.CierreX Then
                MsgBox ("Cierre realizado exitosamente")
            Else
                MsgBox (fiscal.ErrorDesc)
            End If
    
            fiscal.Finalizar
        Else
            MsgBox (fiscal.ErrorDesc)
        End If
        Unload espere
       
    End If
    
   Case Is = "B2"
    J = MsgBox("Confirma la Emision del Cierre Z", 4)
    If J = 6 Then
      fsc_cierrez.Show
      
       
    End If
    
   Case Is = "B3"
     fsc_errorfiscal.Show

 End Select

Else
  MsgBox ("Impresora Fiscal No Inicializada")
End If
End Sub

Private Sub Toolbar6_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case Is = "B1"
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 4 Then
     cja_cajadiaria.Show
    Else
     Call sinpermisos
    End If
 Case Is = "B2"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual >= 4 Then
     cyb_carterach.Show
   Else
    Call sinpermisos
   End If
 Case Is = "B3"
   Call nivel_acceso(1)
   If para.id_grupo_modulo_actual >= 6 Then
     vta_gerencial1.Show
   Else
    Call sinpermisos
   End If
Case Is = "B4"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual >= 4 Then
     cja_detallemov.Show
   Else
    Call sinpermisos
   End If


End Select

End Sub


Private Sub Toolbar7_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim impf As String
'If glo.sucursalf > 0 Then
 Set cl_fiscal = New fiscal
 cl_fiscal.carga (glo.sucursalf)
 If cl_fiscal.id > 0 Then
  Select Case Button.Key
  Case Is = "B1"
   If para.id_grupo_modulo_actual >= 4 Then
     If cl_fiscal.imprimetique = "S" Then
       impf = cl_fiscal.impresora
       fsc_tique.Show
       fsc_tique.t_impfiscal = impf
     Else
       MsgBox ("La Impresora Fiscal Definida no Imprime Tique")
     End If
    Else
     Call sinpermisos
    End If
    

  Case Is = "B2"
    If para.id_grupo_modulo_actual >= 4 Then
     If cl_fiscal.imprimetf = "S" Or cl_fiscal.imprimefact = "S" Then
          vta_facturacion.Show
          vta_facturacion.t_sucursal = Format$(glo.sucursalf, "0000")
          vta_facturacion.c_sucursal.ListIndex = buscaindice(vta_facturacion.c_sucursal, glo.sucursalf)
          vta_facturacion.t_cae = "0"
          vta_facturacion.t_cae_vence = "01/01/2000"
    Else
       MsgBox ("La Impresora Fiscal Definida no Imprime Tique Factura / Factura ")
     End If
    Else
     Call sinpermisos
    End If
  Case Is = "B3"
    If para.id_grupo_modulo_actual >= 6 Then
        vta_remitos.Show
        vta_remitos.t_sucursal = Format$(glo.sucursalf, "0000")
        vta_remitos.c_sucursal.ListIndex = buscaindice(vta_remitos.c_sucursal, glo.sucursalf)
    Else
     Call sinpermisos
    End If
    
  Case Is = "B4"
    If para.id_grupo_modulo_actual >= 6 Then
       vta_recibo.Show
       vta_recibo.sucursal = Format$(glo.sucursalf, "0000")
       vta_recibo.c_sucursal.ListIndex = buscaindice(vta_recibo.c_sucursal, glo.sucursalf)
    Else
     Call sinpermisos
    End If
    
End Select

Else
 MsgBox ("Error en la definicion de la Impresora Fiscal")
End If
Set cl_fiscal = Nothing

'Else
'  MsgBox ("Impresora Fiscal No inicializada")
'End If

End Sub

Private Sub Toolbar8_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case Is = "B1"
    If para.id_grupo_modulo_actual >= 4 Then
      vta_directa.Show
      vta_directa.t_sucursal = Format$(glo.sucursal, "0000")
      vta_directa.c_sucursal.ListIndex = buscaindice(vta_directa.c_sucursal, glo.sucursal)
    
    Else
      Call sinpermisos
    End If

Case Is = "B2"
    If para.id_grupo_modulo_actual >= 4 Then
      vta_liqcereal.Show
      vta_liqcereal.t_sucursal = Format$(glo.sucursal, "0000")
      vta_liqcereal.c_sucursal.ListIndex = buscaindice(vta_liqcereal.c_sucursal, glo.sucursal)
    
    Else
      Call sinpermisos
    End If


End Select

End Sub

Private Sub Toolbar9_ButtonClick(ByVal Button As MSComctlLib.Button)

If glo.sucursalf > 0 Then
 Select Case Button.Key
  Case Is = "B1"
    J = MsgBox("Confirma la Emision del Cierre X", 4)
    If J = 6 Then
       espere.Show
       espere.Refresh
       espere.Label1 = "Espere.... Emitiendo Cierre X"
       
       
       cmd = X_REPORT
       retorno = FP.EnviarComando(cmd)
       Unload espere
       MsgBox ("Cierre X Emitido --> Estado: " & retono)
       
       
    End If
    
   Case Is = "B2"
     J = MsgBox("Confirma la Emision del Cierre Z", 4)
     If J = 6 Then
       espere.Show
       espere.Refresh
       espere.Label1 = "Espere.... Emitiendo Cierre Z"
       
       
       cmd = Z_REPORT
       retorno = FP.EnviarComando(cmd)
       Unload espere
       MsgBox ("Cierre Z Emitido --> Estado: " & retono)
       
       
    End If
   Case Is = "B3"
     fsc_errorfiscal.Show

 End Select

Else
  MsgBox ("Impresora Fiscal No Inicializada")
End If
End Sub
