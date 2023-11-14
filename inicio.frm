VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form inicio 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODULO PRINCIPAL SISTEMA DE GESTION INTEGRADO PARA EMPRESAS"
   ClientHeight    =   8190
   ClientLeft      =   90
   ClientTop       =   -570
   ClientWidth     =   11880
   FontTransparent =   0   'False
   Icon            =   "inicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresora Actual del Sistema"
      Height          =   735
      Left            =   4920
      TabIndex        =   25
      Top             =   7080
      Width           =   4815
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   4080
         Picture         =   "inicio.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label7"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3735
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   650
      ImageHeight     =   184
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":0614
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   10080
      TabIndex        =   21
      Top             =   120
      Width           =   1575
      Begin VB.Image Image6 
         Height          =   495
         Left            =   480
         Picture         =   "inicio.frx":5AC4
         Top             =   2760
         Width           =   765
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   360
         Picture         =   "inicio.frx":6075
         Top             =   240
         Width           =   810
      End
      Begin VB.Image Image2 
         Height          =   315
         Left            =   120
         Picture         =   "inicio.frx":6653
         Top             =   840
         Width           =   1380
      End
      Begin VB.Image Image3 
         Height          =   330
         Left            =   240
         Picture         =   "inicio.frx":6C2A
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Image Image4 
         Height          =   720
         Left            =   480
         Picture         =   "inicio.frx":712E
         Top             =   1800
         Width           =   720
      End
   End
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
      Height          =   2055
      Left            =   8040
      TabIndex        =   18
      Top             =   4080
      Width           =   3735
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "¡Verifique la hora de su computador para un correcto funcionamiento del sistema y del Red!"
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
         Height          =   855
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   2775
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   120
         Picture         =   "inicio.frx":7FF8
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "¡No olvide realizar BACKUP periodicos!"
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
         Height          =   495
         Left            =   720
         TabIndex        =   19
         Top             =   1320
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tools(F12)"
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   5655
      Begin VB.CommandButton Command5 
         Caption         =   "Definir Imp."
         Height          =   735
         Left            =   4560
         Picture         =   "inicio.frx":8302
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Calculadora"
         Height          =   735
         Left            =   2400
         Picture         =   "inicio.frx":860C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Links"
         Height          =   735
         Left            =   3480
         Picture         =   "inicio.frx":8916
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Calendario"
         Height          =   735
         Left            =   1320
         Picture         =   "inicio.frx":8D31
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton command1 
         Caption         =   "Agenda"
         Height          =   735
         Left            =   240
         Picture         =   "inicio.frx":9177
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MODULOS DEL SISTEMA"
      Height          =   2295
      Left            =   480
      TabIndex        =   13
      Top             =   240
      Width           =   8415
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   1680
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   2963
         ButtonWidth     =   2672
         ButtonHeight    =   1429
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "VENTAS"
               Key             =   "B1"
               Description     =   "Ventas y Cuentas Corrientes Clientes"
               Object.ToolTipText     =   "Ventas y Cuentas Corrientes Clientes"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "COMPRAS"
               Key             =   "B2"
               Description     =   "Compras y Cuentas Corrientes Proveedores"
               Object.ToolTipText     =   "Compras y Cuentas Corrientes Proveedores"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "STOCK y C.M."
               Key             =   "B3"
               Description     =   "Productos y Stock"
               Object.ToolTipText     =   "Productos y Stock"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "CAJA"
               Key             =   "B4"
               Description     =   "Modulo Caja"
               Object.ToolTipText     =   "Modulo Caja"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "BANCOS"
               Key             =   "B5"
               Description     =   "Modulo Bancos"
               Object.ToolTipText     =   "Modulo Bancos"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "PRODUCCION"
               Key             =   "B6"
               Description     =   "Produccion"
               Object.ToolTipText     =   "Produccion"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "CONTABILIDAD"
               Key             =   "B7"
               Description     =   "Contabilidad"
               Object.ToolTipText     =   "Contabilidad"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "EMPLEADOS"
               Key             =   "B8"
               Description     =   "Registro de Empleados"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "EXPORTACIONES"
               Key             =   "B9"
               Object.ToolTipText     =   "Modulo para Exportaciones"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "CIERRE MENSUAL"
               Key             =   "B10"
            EndProperty
         EndProperty
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
         Height          =   615
         Left            =   840
         Picture         =   "inicio.frx":95EE
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
         Picture         =   "inicio.frx":9E70
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
      Width           =   11880
      _ExtentX        =   20955
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
            TextSave        =   "10/11/2023"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:39 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":A6F2
            Key             =   "I1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":AA0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":AFA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":B2BB
            Key             =   "I2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":B5D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":B8EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":C1C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":C4E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":CD75
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "inicio.frx":D363
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   360
      TabIndex        =   28
      Top             =   5400
      Width           =   4575
   End
   Begin VB.Menu M_info 
      Caption         =   "&Utiles..."
      Begin VB.Menu m_v1 
         Caption         =   "Acerca del Sistema"
      End
      Begin VB.Menu M_parametros 
         Caption         =   "Parametros Generales"
      End
      Begin VB.Menu M_parametrosusuarios 
         Caption         =   "Parametros por Usuario"
      End
      Begin VB.Menu M_impuestos 
         Caption         =   "Impuestos"
      End
      Begin VB.Menu M_actividades 
         Caption         =   "Actividades Comerciales"
      End
      Begin VB.Menu M_actualizadatos 
         Caption         =   "Actualizacion de  Datos"
      End
      Begin VB.Menu M_definir 
         Caption         =   "Definir Archivo de Exportacion"
      End
      Begin VB.Menu M_membrete 
         Caption         =   "Membretar Hojas"
      End
      Begin VB.Menu M_tools 
         Caption         =   "Utilitarios varios(Tools)"
      End
   End
   Begin VB.Menu M_procesos 
      Caption         =   "&Procesos"
      Begin VB.Menu M_arba 
         Caption         =   "Procesos ARBA/siap"
         Begin VB.Menu m_actualizapadron 
            Caption         =   "Actualizar Padron Retenciones de  Ingresos Brutos"
         End
         Begin VB.Menu m_actualizapadronp 
            Caption         =   "Actualizar Padron de Percepciones de Ingresos Brutos"
         End
         Begin VB.Menu M_embargados 
            Caption         =   "Actualizar Padron de Embargados por  ARBA"
         End
         Begin VB.Menu M_retibweb 
            Caption         =   "Genera Retenciones IB realizadas para presentacion WEB"
         End
         Begin VB.Menu M_generaib2 
            Caption         =   "Generar ret. y  perc. IB para aplic. IB MEN/BIM"
         End
         Begin VB.Menu M_aicyc 
            Caption         =   "AICYC Arba para empresas Constructoras y Corralones"
         End
      End
      Begin VB.Menu M_afip 
         Caption         =   "Procesos AFIP/siap"
         Begin VB.Menu M_actu03 
            Caption         =   "Actualiza Padron facturas Apócrifas(AFIP)"
         End
         Begin VB.Menu M_retsicore 
            Caption         =   "Generar Retenciones de Ganancias  para SICORE"
         End
         Begin VB.Menu M_genretib 
            Caption         =   "Generar Retenciones IB RALIZADAS  para SICORE"
         End
         Begin VB.Menu M_genretiva 
            Caption         =   "Generar retenciones Iva para aplic. Siap/Iva"
         End
         Begin VB.Menu M_citiventas 
            Caption         =   "Generar Exportacion al CITI Ventas(Unificado)"
         End
         Begin VB.Menu Genciticom 
            Caption         =   "Generar Exportacion al CITI Compras(Unificado)"
         End
         Begin VB.Menu m_programa 
            Caption         =   "Programa Asistencia Trabajo y Produccion"
         End
         Begin VB.Menu M_lidv 
            Caption         =   "Libro Iva Digital Ventas"
         End
         Begin VB.Menu M_lidc 
            Caption         =   "Libro Iva Digital Compras(No Bienes de Uso)"
         End
         Begin VB.Menu M_cf1 
            Caption         =   "Informe entre fechas Controladores Fiscales"
         End
      End
      Begin VB.Menu M_borrar 
         Caption         =   "Borrar Datos del Sistema"
         Begin VB.Menu M_depper 
            Caption         =   "Depura Periodo generando saldos"
         End
         Begin VB.Menu M_elimina 
            Caption         =   "Elimina Datos sin generar Saldos"
         End
      End
      Begin VB.Menu M_importarexma 
         Caption         =   "Importar Datos de EXMA"
      End
      Begin VB.Menu M_sincronizar 
         Caption         =   "Sincronizar Datos en la Nube"
      End
      Begin VB.Menu M_agregaprodp 
         Caption         =   "Agrega Productos Lista Proveedor"
      End
   End
   Begin VB.Menu M_seguridad 
      Caption         =   "&Seguridad"
      Begin VB.Menu M_opassword 
         Caption         =   "Definir Nueva Password"
      End
      Begin VB.Menu M_defineseg 
         Caption         =   "Definir Nivel Seguridad"
      End
   End
   Begin VB.Menu M_gerencia 
      Caption         =   "&Informacion Gerencial"
      Begin VB.Menu M_estadoresul 
         Caption         =   "Estado de Resultados"
      End
      Begin VB.Menu M_flujofondos 
         Caption         =   "Flujo de Fondos"
      End
      Begin VB.Menu M_balance 
         Caption         =   "Balance General"
      End
      Begin VB.Menu M_sumasysaldos 
         Caption         =   "Balance comprobacion Sumas y Saldos"
      End
   End
   Begin VB.Menu M_salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnsale_Click()
 End
End Sub


Sub HABILITACION()
'numeros de modulos 5732 ----> Todo sin fiscal
'                   2514 ----> Sin Producion y Contabilidad
'                   1595 ----> Fiscal
'                   1320 ----> Solo Caja
'                   1722 ----> Sin Produccion, Empleados y Contabilidad
'                   9999 ----> todo
'                   1723 ----> VETERINARIAS Sin Produccion, Empleados y Contabilidad
'                   1922 ----> Ventas, Compras, caja, bancos
'                   8574 ----> Solo Compras

sc = 0
For i = 1 To 2
  sc = sc + Val(Mid$(glo.CUIT, i, 1))
Next i
For i = 4 To 11
  sc = sc + Val(Mid$(glo.CUIT, i, 1))
Next i
sc = sc + Val(Mid$(glo.CUIT, 13, 1))

Set rs = New ADODB.Recordset
q = "select [habilitacion] from g0 where [sucursal] = 0"
rs.Open q, cn1
If Val(Mid$(rs("habilitacion"), 5, 2)) = sc Then
  para.HABILITACION = Val(Mid$(rs("habilitacion"), 1, 4))
  Select Case para.HABILITACION
    Case Is = 5732 ' todo sin fiscal
       Toolbar1.Buttons.item(1).Visible = True
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = True
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = True
       Toolbar1.Buttons.item(6).Visible = True
       Toolbar1.Buttons.item(7).Visible = True
       Toolbar1.Buttons.item(8).Visible = True
       Toolbar1.Buttons.item(9).Visible = True
       Toolbar1.Buttons.item(10).Visible = True
       
       para.fiscal = 0
   Case Is = 1795 'sin produccion y sin fiscal y exportacion
       Toolbar1.Buttons.item(1).Visible = True
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = True
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = True
       Toolbar1.Buttons.item(6).Visible = False
       Toolbar1.Buttons.item(7).Visible = True
       Toolbar1.Buttons.item(8).Visible = True
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = True
       
       para.fiscal = 0
   Case Is = 1595 ' fiscal
     
       Toolbar1.Buttons.item(1).Visible = True
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = True
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = True
       Toolbar1.Buttons.item(6).Visible = False
       Toolbar1.Buttons.item(7).Visible = False
       Toolbar1.Buttons.item(8).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = True
       
     If glo.sucursalf > 0 Then
       para.fiscal = 1
     Else
       para.fiscal = 0
     End If
     
      Case Is = 1912 ' fiscal sin ventas normales
     
       Toolbar1.Buttons.item(1).Visible = True
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = True
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = True
       Toolbar1.Buttons.item(6).Visible = False
       Toolbar1.Buttons.item(7).Visible = False
       Toolbar1.Buttons.item(8).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = True
       
     If glo.sucursalf > 0 Then
       para.fiscal = 1
     Else
       para.fiscal = 0
     End If

   Case Is = 1320 ' solo caja
       Toolbar1.Buttons.item(1).Visible = False
       Toolbar1.Buttons.item(2).Visible = False
       Toolbar1.Buttons.item(3).Visible = False
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = False
       Toolbar1.Buttons.item(6).Visible = False
       Toolbar1.Buttons.item(7).Visible = False
       Toolbar1.Buttons.item(8).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = False
       
   Case Is = 1722, Is = 1723, Is = 1825, Is = 1721 'sin produccion, sin fiscal, sin empleados y sin contabilidad
       Toolbar1.Buttons.item(1).Visible = True
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = True
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = True
       Toolbar1.Buttons.item(6).Visible = False
       Toolbar1.Buttons.item(7).Visible = False
       Toolbar1.Buttons.item(8).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = False
       
   
       para.fiscal = 0
   
   Case Is = 8574 'solo cmpras
       Toolbar1.Buttons.item(1).Visible = False
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = False
       Toolbar1.Buttons.item(4).Visible = False
       Toolbar1.Buttons.item(5).Visible = False
       Toolbar1.Buttons.item(6).Visible = False
       Toolbar1.Buttons.item(7).Visible = False
       Toolbar1.Buttons.item(8).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = False
   
   
   Case Is = 9999 ' todo
       Toolbar1.Buttons.item(1).Visible = True
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = True
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = True
       Toolbar1.Buttons.item(6).Visible = True
       Toolbar1.Buttons.item(7).Visible = True
       Toolbar1.Buttons.item(8).Visible = True
       Toolbar1.Buttons.item(9).Visible = True
       Toolbar1.Buttons.item(10).Visible = True
       
       If glo.sucursalf > 0 Then
       para.fiscal = 1
     Else
       para.fiscal = 0
     End If
   
   Case Is = 1922 ' ventas, compras, caja, bancos
       Toolbar1.Buttons.item(1).Visible = True
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = False
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = True
       Toolbar1.Buttons.item(6).Visible = False
       Toolbar1.Buttons.item(7).Visible = False
       Toolbar1.Buttons.item(8).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = False
              
       para.fiscal = 0
   
   Case Is = 7422 ' todo sin fiscal
       Toolbar1.Buttons.item(1).Visible = True
       Toolbar1.Buttons.item(2).Visible = True
       Toolbar1.Buttons.item(3).Visible = True
       Toolbar1.Buttons.item(4).Visible = True
       Toolbar1.Buttons.item(5).Visible = True
       Toolbar1.Buttons.item(6).Visible = False
       Toolbar1.Buttons.item(7).Visible = False
       Toolbar1.Buttons.item(8).Visible = False
       Toolbar1.Buttons.item(9).Visible = False
       Toolbar1.Buttons.item(10).Visible = False
       
       para.fiscal = 0
   
   Case Else
       MsgBox ("Error Inesperado de Habilitacion: E1001 ")
       End
       
  End Select
Else
 MsgBox ("Error Inesperado de Habilitacion: E1002 ")
 End
End If

End Sub





Private Sub Command1_Click()
gen_agenda.Show
End Sub


Private Sub Command2_Click()
gen_calendario.Show
End Sub

Private Sub Command3_Click()
gen_links.Show
End Sub

Private Sub Command4_Click()
s = Shell(App.Path & "\tools\calc.exe", vbNormalFocus)
End Sub



Private Sub Command5_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Command6_Click()
gen_seleccionarimp.Show
End Sub





Private Sub Command7_Click()
F = "04/08/2022"
f2 = DateValue(F) + (2 * 30)
MsgBox (f2)


End Sub

Private Sub Form_Activate()
Call barraesag(Me)
Call HABILITACION
Label7 = para.impresora_actual
Text5 = para.punto_venta_usuario
Label8.Caption = para.usuario
End Sub

Private Sub Form_Load()
 
Call titulos(Me)
Call carga_iva
Load vta_listaprecios
If para.muestraagenda = "S" Then
 gen_agenda.Show
End If
Exit Sub

e1:
  MsgBox ("Error al Inicializar Parametros INICIO.LOAD")
  End

End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload vta_listaprecios
End Sub

Private Sub Genciticom_Click()
gen_citicom.Show
End Sub

Private Sub Image6_Click()
'FIXIT: Declare 'intobj' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim intobj As Object
Set intobj = CreateObject("InternetExplorer.Application")
intobj.Visible = -1
intobj.Navigate "http://www.arba.gov.ar"

End Sub

Private Sub M_actividades_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
 gen_ABM_actividad.Show
Else
  Call sinpermisos
End If

End Sub

Private Sub M_actu03_Click()
gen_factapocrifas.Show
End Sub






Private Sub M_actualizapadron_Click()
'el mismo nivel de acceso que ventas
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
  gen_padronib.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub m_actualizapadronp_Click()
'el mismo nivel de acceso que ventas
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
  gen_padronibp.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_agregaprodp_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 7 Then
  vta_cargaprod_listaprov.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_aicyc_Click()
'el mismo nivel de acceso que compras
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 7 Then
  vta_arba_corralones.Show
Else
  Call sinpermisos
End If



End Sub

Private Sub M_balance_Click()
cgr_balanceprov.Show
End Sub

Private Sub M_cf1_Click()
gen_cf.Show
End Sub

Private Sub M_citiventas_Click()
gen_citi.Show
End Sub

Private Sub M_defineseg_Click()
'el mismo nivel de acceso que ventas
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
  gen_cambioseguridad.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_definir_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
   exportar.Show
End If
End Sub


Private Sub M_depper_Click()
'el mismo nivel de acceso que compras
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
  gen_depuraperiodo.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_elimina_Click()
'el mismo nivel de acceso que compras
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
  gen_borradatos.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_embargados_Click()
'el mismo nivel de acceso que ventas
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
  gen_embargoib.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_estadoresul_Click()
gen_estadoresultado.Show
End Sub

Private Sub M_flujofondos_Click()
cja_detallemov2.Show
End Sub

Private Sub M_generaib2_Click()
gen_exporta_retib2.Show

End Sub

Private Sub M_genretib_Click()
  Call nivel_acceso(2)
  If para.id_grupo_modulo_actual > 6 Then

   gen_exportaretib.Show
  Else
    Call sinpermisos
  End If
End Sub

Private Sub M_genretiva_Click()
gen_exporta_retiva.Show
End Sub

Private Sub M_importarexma_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
   gen_migrardatos.Show
End If
End Sub

Private Sub M_impuestos_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
 gen_impuestos.Show
Else
 Call sinpermisos
End If


End Sub

Private Sub M_lidc_Click()
 Call nivel_acceso(1)
  If para.id_grupo_modulo_actual > 6 Then
    gen_libroivadigitalC.Show
  Else
    Call sinpermisos
  End If
End Sub

Private Sub M_lidv_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 6 Then
  gen_libroivadigitalV.Show
Else
  Call sinpermisos
End If
End Sub

Private Sub M_membrete_Click()
J = InputBox$("Cantidad de Copias", "Membretar Hojas", 1)
If Val(J) > 0 Then
 Load gen_logo
 For i = 1 To Val(J)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.PaintPicture gen_logo.Picture1.Picture, 0, 0
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.NewPage
 Next i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc
 Unload gen_logo
End If
End Sub

Private Sub M_opassword_Click()
gen_cambiopassword.Show
End Sub

Private Sub M_parametros_Click()
'el mismo nivel de acceso que ventas
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
  gen_parametros.Show
Else
  Call sinpermisos
End If


End Sub

Private Sub M_parametrosusuarios_Click()
Call nivel_acceso(1)
If para.id_grupo_modulo_actual > 8 Then
 gen_parametrosusuarios.Show
End If
End Sub

Private Sub m_programa_Click()
gen_asistencia.Show
End Sub

Private Sub M_retibweb_Click()
gen_exportaretibweb.Show
End Sub

Private Sub M_retsicore_Click()
  Call nivel_acceso(2)
  If para.id_grupo_modulo_actual > 6 Then
    gen_exportasicore.Show
  Else
    Call sinpermisos
  End If
End Sub

Private Sub M_salir_Click()
End
End Sub


Private Sub M_sincronizar_Click()
Call nivel_acceso(1)
If para.idsistema > 0 And para.id_grupo_modulo_actual > 8 Then
   gen_sincronizar.Show
Else
  MsgBox ("El sistema WEB(OnLine) no esta habilitado o usted no tiene permisos para esta operación")
End If
End Sub

Private Sub M_sumasysaldos_Click()
cgr_sumasysaldosp.Show
End Sub

Private Sub M_tools_Click()
gen_tools.Show
End Sub

Private Sub m_v1_Click()
frmAbout.Show
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
 Case Is = "B1"
  Call nivel_acceso(1)
  If para.id_grupo_modulo_actual > 0 Then
   inicio_vta.Show
   
  Else
   Call sinpermisos
  End If
 
 Case Is = "B2"
   Call nivel_acceso(2)
   If para.id_grupo_modulo_actual > 0 Then
     inicio_compras.Show
     
   Else
     Call sinpermisos
   End If

 Case Is = "B3"
   Call nivel_acceso(5)
   If para.id_grupo_modulo_actual > 0 Then
     inicio_stk.Show
     
   Else
     Call sinpermisos
   End If


 Case Is = "B4"
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual > 0 Then
    inicio_caja.Show
    
   Else
     Call sinpermisos

   End If
   
 Case Is = "B5"
   Call nivel_acceso(4)
   If para.id_grupo_modulo_actual > 0 Then
    inicio_bancos.Show
    
   Else
     Call sinpermisos

   End If
 
 Case Is = "B6"
   Call nivel_acceso(6)
   If para.id_grupo_modulo_actual > 0 Then
     inicio_produccion.Show
    
   Else
     Call sinpermisos
   End If
 
  Case Is = "B7"
   Call nivel_acceso(7)
   If para.id_grupo_modulo_actual > 0 Then
     inicio_CGR.Show
    
   Else
     Call sinpermisos
   End If

 Case Is = "B8"
   'compras
   Call nivel_acceso(1)
   If para.id_grupo_modulo_actual > 0 Then
     inicio_empleados.Show
    
   Else
     Call sinpermisos
   End If

 Case Is = "B9" 'exportaciones
   'compras
   Call nivel_acceso(2)
   If para.id_grupo_modulo_actual > 5 Then
     inicio_exporta.Show
    
   Else
     Call sinpermisos
   End If
 
 Case Is = "B10" 'cierre
   'compras
   Call nivel_acceso(2)
   If para.id_grupo_modulo_actual = 9 Then
     gen_cierremes.Show
    
   Else
     Call sinpermisos
   End If


End Select
End Sub
