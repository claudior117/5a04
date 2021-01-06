VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_listaprecios3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8925
   ScaleWidth      =   12270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame10 
      Caption         =   "¡IMPORTANTE! Definir tipo de Caluclo de PU"
      Height          =   1455
      Left            =   5760
      TabIndex        =   69
      Top             =   4800
      Width           =   5055
      Begin VB.CheckBox Check2 
         Caption         =   "Calcula PF desde PU"
         Height          =   375
         Left            =   2760
         TabIndex        =   73
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Calcula PU a partir del costo"
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Calcula PU a partir PF"
         Height          =   195
         Left            =   240
         TabIndex        =   71
         Top             =   720
         Width           =   4335
      End
      Begin VB.OptionButton Option7 
         Caption         =   "No Modifica PU"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   1080
         Width           =   4335
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "¡IMPORTANTE! Definir tipo de Calculo de Costos"
      Height          =   1455
      Left            =   960
      TabIndex        =   65
      Top             =   4800
      Width           =   4575
      Begin VB.OptionButton Option6 
         Caption         =   "No Modifica Costos"
         Height          =   195
         Left            =   240
         TabIndex        =   68
         Top             =   1080
         Width           =   4335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Calcula Costos a partir del Nuevo  Precio Final Venta"
         Height          =   195
         Left            =   240
         TabIndex        =   67
         Top             =   720
         Width           =   4335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Calcula Costos a parti del Nuevo  Precio Compra"
         Height          =   195
         Left            =   240
         TabIndex        =   66
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Redondeo para los calculos"
      Height          =   975
      Left            =   240
      TabIndex        =   56
      Top             =   6360
      Width           =   11055
      Begin VB.OptionButton Option11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "10/100  Ej. 8.10"
         Height          =   255
         Left            =   4560
         TabIndex        =   77
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "25/100  Ej. 8.25"
         Height          =   255
         Left            =   6840
         TabIndex        =   76
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Aplicar redondeos sin actualizar precios"
         Height          =   255
         Left            =   2640
         TabIndex        =   62
         Top             =   600
         Width           =   5175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin redondeo Ej. 8.63"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entero Ej. 8.00"
         Height          =   255
         Left            =   2400
         TabIndex        =   58
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "50/100  Ej.  8.50"
         Height          =   255
         Left            =   9000
         TabIndex        =   57
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Etiquetas"
      Height          =   615
      Left            =   8280
      TabIndex        =   54
      Top             =   3960
      Width           =   3375
      Begin VB.ComboBox c_etiquetas 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "vta006B.frx":0000
         Left            =   120
         List            =   "vta006B.frx":000D
         TabIndex        =   55
         Text            =   "Combo1"
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Variaciones en $ fija"
      Height          =   1215
      Left            =   120
      TabIndex        =   47
      Top             =   3360
      Width           =   3615
      Begin VB.TextBox t_pcompra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox t_pventa 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00800080&
         Caption         =   "Precio Fijo  Compra S/iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label14 
         BackColor       =   &H00800080&
         Caption         =   "Precio Fijo Venta Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estructura de Costos"
      Height          =   2295
      Left            =   3840
      TabIndex        =   29
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox t_dtocompra2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   60
         ToolTipText     =   "Se aplica el dto1 y despues el dto2"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox t_flete_compra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   36
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox t_dtocompra 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   34
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox c_iva 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   32
         Text            =   "Combo1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox t_utilidad 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label17 
         BackColor       =   &H00800080&
         Caption         =   "Dto. Compra2 (d1 + d2)"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00800080&
         Caption         =   "Flete Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00800080&
         Caption         =   "Dto. Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800080&
         Caption         =   "Tasa de Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "% Utilidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Variaciones Precios  en %"
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   3615
      Begin VB.TextBox t_porc_pv 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox t_porc_pc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "% Variacion Precio Venta Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "% Variacion Precio Compra"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modificaciones Varios"
      Height          =   3735
      Left            =   8280
      TabIndex        =   12
      Top             =   120
      Width           =   3255
      Begin VB.TextBox t_stock 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   78
         Top             =   3240
         Width           =   975
      End
      Begin VB.ComboBox c_vigente 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "vta006B.frx":004E
         Left            =   1080
         List            =   "vta006B.frx":005B
         TabIndex        =   64
         Text            =   "Combo1"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox c_tasaib 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "vta006B.frx":008B
         Left            =   1080
         List            =   "vta006B.frx":008D
         TabIndex        =   52
         Text            =   "Combo1"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox t_tipoc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   44
         ToolTipText     =   "[M] Manual - [A] Automatica"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox t_moneda 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   21
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox t_tipo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox t_envase 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox t_stockminimo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox c_unidad 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label25 
         BackColor       =   &H00800080&
         Caption         =   "Stock"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label23 
         BackColor       =   &H00800080&
         Caption         =   "Vigente"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00800080&
         Caption         =   "Tasa IB"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00800080&
         Caption         =   "Tipo Carga"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[Manual] - [Autom.]"
         Height          =   255
         Left            =   1560
         TabIndex        =   45
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[Pesos] - [Dolares]"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[Prod] - [Mat. Prima]"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00800080&
         Caption         =   "Moneda"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label21 
         BackColor       =   &H00800080&
         Caption         =   "Tipo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label20 
         BackColor       =   &H00800080&
         Caption         =   "Envase"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label19 
         BackColor       =   &H00800080&
         Caption         =   "Stock Min."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label18 
         BackColor       =   &H00800080&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   9120
      TabIndex        =   10
      Top             =   7440
      Width           =   2655
      Begin VB.CommandButton Command4 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Categoriazacion de los productos"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   6360
         Picture         =   "vta006B.frx":008F
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   6360
         Picture         =   "vta006B.frx":0194
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   6360
         Picture         =   "vta006B.frx":0299
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   6360
         Picture         =   "vta006B.frx":039E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   735
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
      Begin VB.Label Label7 
         BackColor       =   &H00800080&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
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
      Top             =   8565
      Width           =   12270
      _ExtentX        =   21643
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
            TextSave        =   "03/01/2006"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "01:35 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Este modulo solo modifica PF, se recomienda tambien modificar el PU desde las  opciones de Caluclo PU"
      Height          =   255
      Left            =   360
      TabIndex        =   75
      Top             =   7920
      Width           =   8415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Las actualizaciones por % tienen prioridad sobre las fijas. "
      Height          =   255
      Left            =   360
      TabIndex        =   74
      Top             =   7560
      Width           =   6735
   End
End
Attribute VB_Name = "vta_listaprecios3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984



Sub calcula()
d = Format$(Val(t_preciocompra) * Val(t_dtocompra) / 100, "#####0.000")
n = Val(t_preciocompra) - Val(d)
F = n * Val(t_fletecompra) / 100
n2 = F + n
t_costo = Format$(n2, "#####0.000")
t_pu = Format$(Val(t_costo) + (Val(t_costo) * Val(t_utilidad) / 100), "#####0.000")
t_impuesto = Format$(Val(t_preciocompra) * Val(t_tasaimpint) / 100, "####0.0000")
t_final = Format$(Val(t_pu) + (Val(t_pu) * Val(c_iva) / 100) + Val(t_impuesto), "######0.00")
t_preciocompra = Format$(Val(t_preciocompra), "#####0.000")
t_dtocompra = Format$(Val(t_dtocompra), "#####0.000")
t_fletecompra = Format$(Val(t_fletecompra), "#####0.000")
t_tasaimpint = Format$(Val(t_tasaimpint), "#####0.000")



End Sub

Private Sub c_iva_LostFocus()
If c_iva.ListIndex < 0 Then
  c_iva.ListIndex = 0
End If

End Sub

Private Sub c_tasaib_LostFocus()
If c_tasaib.ListIndex < 0 Then
  c_tasaib.ListIndex = 0
End If

End Sub

Private Sub c_unidad_LostFocus()
If c_unidad.ListIndex < 0 Then
  c_unidad.ListIndex = 0
End If
  
End Sub

Private Sub c_vigente_LostFocus()
If c_vigente.ListIndex < 0 Then
  c_vigente.ListIndex = 0
End If

End Sub

Private Sub Command1_Click()
ABM_grupos.Show
End Sub

Private Sub Command1_LostFocus()
Call carga_grupos(c_grupo)
End Sub

Private Sub Command2_Click()
Call graba
End Sub
Sub graba()
h = MsgBox("Confirma Valores para Grabar", 4)
If h = 6 Then
   'On Error GoTo ERRORGRABA
    espere.Show
    espere.Label1 = "Espere....  Actualizando Datos de Productos"
    espere.Refresh
     
  QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
  QUERY = QUERY & " VALUES ('Modificacion Grupal de Precios " & "', " & para.id_usuario & ", 'V', 0, '" & Now & "', ' ', 21, " & 0 & ")"
  cn1.BeginTrans
  cn1.Execute QUERY
  cn1.CommitTrans
     
    Select Case para.tipolistaprecios
    Case Is = 2
        Filas = vta_listaprecios_2.msf1.Rows
    Case Is = 3
        Filas = vta_listaprecios_3.msf1.Rows
    
    Case Else
        Filas = vta_listaprecios.msf1.Rows
    End Select
     
    J = 1
    While J < Filas
       Select Case para.tipolistaprecios
       Case Is = 2
        idprod = Val(vta_listaprecios_2.msf1.TextMatrix(J, 0))
       Case Is = 3
        idprod = Val(vta_listaprecios_3.msf1.TextMatrix(J, 0))
       
       Case Else
        idprod = Val(vta_listaprecios.msf1.TextMatrix(J, 0))
       End Select
    
    
      If idprod > 1 Then
         Set rs = New adodb.Recordset
         q = "select * from a2, g4 where [id_producto] = " & idprod & " and [cod_tasaiva] = [id_tasaiva]"
         rs.MaxRecords = 1
         rs.Open q, cn1, adOpenStatic, adLockOptimistic
         If Not rs.BOF And Not rs.EOF Then
           'comienzo actualizacion
           If c_grupo.ListIndex > 0 Then
             ig = c_grupo.ItemData(c_grupo.ListIndex)
           Else
             ig = rs("id_grupo")
           End If
   
            If c_depto.ListIndex > 0 Then
             id = c_depto.ItemData(c_depto.ListIndex)
           Else
             id = rs("id_departamento")
           End If
           
            If c_marca.ListIndex > 0 Then
             im = c_marca.ItemData(c_marca.ListIndex)
           Else
             im = rs("id_marca")
           End If
             
            If c_prov.ListIndex > 0 Then
             ip = c_prov.ItemData(c_prov.ListIndex)
           Else
             ip = rs("id_proveedor")
           End If
           
            If c_unidad.ListIndex > 0 Then
             iu = c_unidad.ItemData(c_unidad.ListIndex)
           Else
             iu = rs("id_unidad")
           End If
           
           If c_iva.ListIndex > 0 Then
             ii = c_iva.ItemData(c_iva.ListIndex)
             ti4 = c_iva
           Else
             ii = rs("cod_tasaiva")
             ti4 = rs("tasa")
           End If
           
            If c_tasaib.ListIndex > 0 Then
             iib = c_tasaib.ItemData(c_tasaib.ListIndex)
           Else
             iib = rs("id_tasaib")
           End If
           
            Select Case c_vigente.ListIndex
            Case Is = 1
             vig = True
            Case Is = 2
             vig = False
            Case Else
              vig = False
            End Select
                     
           If Val(t_porc_pc) <> 0 Then
              pc = Format(rs("precio_ult_compra") * (1 + (Val(t_porc_pc) / 100)), "######0.00")
           Else
             If Val(t_pcompra) <> 0 Then
               pc = Format(Val(t_pcompra), "#####0.00")
              Else
               pc = rs("precio_ult_compra")
             End If
           End If
           
           cambiapv = 0
           If Val(t_porc_pv) <> 0 Then
              'pu = rs("pu") + (rs("pu") * (Val(t_porc_pv) / 100))
              pf = rs("precio_final") + (rs("precio_final") * (Val(t_porc_pv) / 100))
              cambiapv = 1
           Else
             If Val(t_pventa) <> 0 Then
               pf = Format(Val(t_pventa), "#####0.00")
               'pu = Format(pf / (1 + (rs("tasa") / 100)), "#####0.00")
               cambiapv = 1
             Else
               'pu = rs("pu")
               pf = rs("precio_final")
             End If
        
           End If
           
           
          If Val(t_utilidad) > 0 Then
             u = Val(t_utilidad)
          Else
             u = rs("porc_utilidad")
          End If
          
         ' If Val(t_pu) > 0 Then
          '    pu = Format(Val(t_pu), "######0.00")
           '   pf = pu + (pu * para.tasaiva(ii) / 100)
         '     cambiapv = 1
          'End If
           
          If Val(t_dtocompra) > 0 Then
             dc = Val(t_dtocompra)
          Else
             dc = rs("dto_compra")
          End If
           
          If Val(t_dtocompra2) > 0 Then
             dc2 = Val(t_dtocompra2)
          Else
             dc2 = rs("dto_compra2")
          End If
           
           
          If Val(t_flete_compra) > 0 Then
             fc = Val(t_flete_compra)
          Else
             fc = rs("flete_compra")
          End If
           
          If Option3 = True Then
            'calcula costo a partir de precio compra
             d = pc * (dc / 100)
             n = pc - d
             d2 = n * (dc2 / 100)
             n = n - d2
             F = n * (fc / 100)
            cr = Format(F + n, "########0.00")
          Else
            If Option4 = True Then
            'calcula costo a partir de precio venta final
             n = pf / (1 + (ti4 / 100))
             n2 = n / (1 + (u / 100))
             cr = Format(n2, "########0.00")
            Else
             cr = Format(rs("costoreal"), "########0.00")
            End If
          End If
          
        
          If Option9 = True Then
            'calcula pu a partir del costo
            pu = cr * (1 + (u / 100))
          Else
           If Option8 = True Then
            'calcula pu a partir de pf
            pu = pf / (1 + (ti4 / 100))
           Else
            pu = rs("pu")
           End If
          End If
          
          
          If Check2 = 1 Then
            'actualizo pf a partir del pu
             If cambiapv = 1 Then
                MsgBox ("Ha seleccionado cambiar el precio final por porcentaje o a valores fijos, No es posible calcularlo a traves del PU")
             Else
               cambiapv = 1
               pf = pu * (1 + (ti4 / 100))
             End If
          End If
          
          If cambiapv = 1 Then
            fa = Format$(Now, "dd/mm/yyyy")
          
            'decimales
            If Option1 = True Then
              'dos decimales
              pf = Format(pf, "######0.00")
            Else
             If Option2 = True Then
               'entero
               pf = Format(pf, "######0")
             Else
              If Option10 = True Then
               '0.25
               pf = redondeanum(Format(pf, "#######0.00"), 1)
              Else
               If Option5 = True Then
                 '0.50
                  pf = redondeanum(Format(pf, "#######0.00"), 2)
               Else
                  '0.10
                  pf = Format(pf, "######0.0")
               End If
              End If
             End If
           End If
          
          Else
            fa = rs("fecha_actu_precio_venta")
          End If
          
          rs("id_grupo") = ig
          rs("id_departamento") = id
          rs("id_marca") = im
          rs("id_proveedor") = ip
          rs("id_unidad") = iu
          rs("cod_tasaiva") = ii
          rs("id_tasaib") = iib
          rs("precio_ult_compra") = pc
          rs("pu") = Format(pu, "#######0.00")
          rs("precio_final") = pf
          rs("porc_utilidad") = u
          rs("dto_compra") = dc
          rs("dto_compra2") = dc2
          rs("flete_compra") = fc
          rs("costoreal") = cr
          rs("fecha_actu_precio_venta") = fa
          If t_moneda <> "" Then
            rs("moneda") = t_moneda
          End If
          If t_tipo <> "" Then
            rs("tipo_producto") = t_tipo
          End If
          If t_tipoc <> "" Then
            rs("tipo_carga_tique") = t_tipoc
          End If
                           
         'etiquetas
          If c_etiquetas.ListIndex > 0 Then
            If c_etiquetas.ListIndex = 1 Then
               rs("emite_etiqueta") = "S"
            Else
              rs("emite_etiqueta") = "N"
            End If
          End If
          
          If c_vigente.ListIndex > 0 Then
            rs("vigente") = vig
          End If
          
          If t_stock <> "" Then
            rs("stock") = Val(t_stock)
          End If
          
          rs.Update
          Set rs = Nothing
        End If
       
       
       End If
      J = J + 1
    Wend
   'Call vta_listaprecios.carga
   Unload espere
   Me.Hide
    
End If

Exit Sub

ERRORGRABA:
  'cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub

Sub redondea()
h = MsgBox("Confirma redondear precios finales", 4)
If h = 6 Then
   'On Error GoTo ERRORGRABA
    espere.Show
    espere.Label1 = "Espere....  Actualizando Datos de Productos"
    espere.Refresh
     
     
    Select Case para.tipolistaprecios
    Case Is = 2
        Filas = vta_listaprecios_2.msf1.Rows
    Case Is = 3
        Filas = vta_listaprecios_3.msf1.Rows
    
    Case Else
        Filas = vta_listaprecios.msf1.Rows
    End Select
     
    J = 1
    While J < Filas
       Select Case para.tipolistaprecios
       Case Is = 2
        idprod = Val(vta_listaprecios_2.msf1.TextMatrix(J, 0))
       Case Is = 3
        idprod = Val(vta_listaprecios_3.msf1.TextMatrix(J, 0))
       
       Case Else
        idprod = Val(vta_listaprecios.msf1.TextMatrix(J, 0))
       End Select
    
    
      If idprod > 1 Then
         Set rs = New adodb.Recordset
         q = "select * from a2, g4 where [id_producto] = " & idprod & " and [cod_tasaiva] = [id_tasaiva]"
         rs.MaxRecords = 1
         rs.Open q, cn1, adOpenStatic, adLockOptimistic
         If Not rs.BOF And Not rs.EOF Then
           'comienzo actualizacion
            pf = rs("precio_final")
            'decimales
            If Option1 = True Then
              'dos decimales
              pf = Format(pf, "######0.00")
            Else
             If Option2 = True Then
               'entero
               pf = Format(pf, "######0")
             Else
              If Option10 = True Then
               '0.25
               pf = redondeanum(Format(pf, "#######0.00"), 1)
              Else
               If Option5 = True Then
                  '0.5
                  pf = redondeanum(Format(pf, "#######0.00"), 2)
               Else
                    '0.1
                    pf = Format(pf, "######0.0")
               End If
             End If
           End If
          End If
          rs("precio_final") = Format(pf, "#######0.00")
          
          rs.Update
          Set rs = Nothing
        End If
       
       
       End If
      J = J + 1
    Wend
   'Call vta_listaprecios.carga
   Unload espere
   Me.Hide
    
End If

Exit Sub

ERRORGRABA:
  'cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub

Private Sub Command3_Click()
 ABM_deptoS.Show
End Sub

Private Sub Command3_LostFocus()
Call carga_deptos_venta(c_depto)
End Sub

Private Sub Command4_Click()
Me.Hide
End Sub

Private Sub Command5_Click()
ABM_marcas.Show
End Sub

Private Sub Command5_LostFocus()
Call carga_marcas(c_marca)
Call carga_marcas(vta_listaprecios.c_marca)
vta_listaprecios.c_marca = "<Todos>,0"
End Sub

Private Sub Command6_Click()
ABM_PROv.Show
End Sub

Private Sub Command6_LostFocus()
Call carga_proveedores(c_prov)
End Sub

Private Sub Command7_Click()
Call redondea
End Sub

Private Sub Form_Activate()
c_etiquetas.ListIndex = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Me.Hide
End If

End Sub

Private Sub Form_Load()
  Call carga_grupos(c_grupo)
  c_grupo.AddItem "Sin Cambios", 0
  c_grupo.ListIndex = 0
  Call carga_deptos_venta(c_depto)
  c_depto.AddItem "Sin Cambios", 0
  c_depto.ListIndex = 0
  Call carga_marcas(c_marca)
  c_marca.AddItem "Sin Cambios", 0
  c_marca.ListIndex = 0
  Call carga_proveedores(c_prov)
  c_prov.AddItem "Sin Cambios", 0
  c_prov.ListIndex = 0
  Call carga_unidad(c_unidad)
  c_unidad.AddItem "Sin Cambios", 0
  c_unidad.ListIndex = 0
  Call carga_tasaiva(c_iva)
  c_iva.AddItem "Sin Cambios", 0
  c_iva.ListIndex = 0
  Call carga_tasaib(c_tasaib)
  c_tasaib.AddItem "Sin Cambios", 0
  c_tasaib.ListIndex = 0
  c_etiquetas.ListIndex = 0
  Option1 = True
  Check1 = 0
  c_vigente.ListIndex = 0

  Option6 = True
  Option7 = True
  Check2 = False
  Check2.Enabled = False
  
  Call carga_redondeo
End Sub

 Sub carga_redondeo()
Select Case para.tiporedondeo
Case Is = 0
  Option1 = True
Case Is = 1
  Option11 = True
Case Is = 2
 Option10 = True
Case Is = 3
 Option5 = True
Case Is = 4
 Option2 = True
Case Else
 Option1 = True
End Select

End Sub







Private Sub Option7_Click()
chekc2 = False
Check2.Enabled = False

End Sub

Private Sub Option8_Click()
chekc2 = False
Check2.Enabled = False
End Sub

Private Sub Option9_Click()

Check2.Enabled = True
Check2 = 1
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

Private Sub t_tipoc_LostFocus()
t_tipoc = Format$(t_tipoc, ">@")
If t_tipoc <> "M" And t_tipoc <> "A" Then
  t_tipoc = ""
End If

End Sub
