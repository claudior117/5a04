VERSION 5.00
Begin VB.Form gen_parametrosusuarios 
   Caption         =   "Parametros por usuarios "
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "IMPORTANTE"
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   4695
      Begin VB.Label Label23 
         Caption         =   "Si realiza cambios en la configuración, deberá salir del Sistema y volver a Ingresar para un correcto funcionamiento. "
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4800
      TabIndex        =   6
      Top             =   5160
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen011.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen011.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parametros"
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6375
      Begin VB.TextBox t_sucvta 
         Height          =   285
         Left            =   3000
         TabIndex        =   22
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Imprime Cabecera en los Reportes"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   4335
      End
      Begin VB.ComboBox c_rc 
         Height          =   315
         ItemData        =   "gen011.frx":1104
         Left            =   1920
         List            =   "gen011.frx":1111
         TabIndex        =   19
         Top             =   3360
         Width           =   3855
      End
      Begin VB.ComboBox c_estilo 
         Height          =   315
         ItemData        =   "gen011.frx":1141
         Left            =   1920
         List            =   "gen011.frx":1151
         TabIndex        =   15
         Top             =   2880
         Width           =   3855
      End
      Begin VB.ComboBox c_tlp 
         Height          =   315
         ItemData        =   "gen011.frx":1185
         Left            =   1920
         List            =   "gen011.frx":118F
         TabIndex        =   13
         Top             =   2400
         Width           =   3855
      End
      Begin VB.ComboBox c_tipoprecio 
         Height          =   315
         ItemData        =   "gen011.frx":11B0
         Left            =   1920
         List            =   "gen011.frx":11BA
         TabIndex        =   11
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox t_tl 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Imprime Pie Informativo en los Reportes"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   4335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Muestra Agenda al Inicio"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Punto Venta Incio"
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Estilo Resumen Cuenta"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Estilo Lista de Precios"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Precio Utilizado en Lista de Precios"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C00000&
         Caption         =   "Precio Utilizado para Facturar"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Lista de Precios"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Usuario"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.TextBox t_nombre 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox t_id 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "gen_parametrosusuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnacepta_Click()
J = MsgBox("Actualiza Parametros", 4)
If J = 6 Then
q = "SELECT * FROM G1 WHERE id_usuario = " & Val(t_id)
Set rs = New adodb.Recordset
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
   If Check1 = 1 Then
      rs("muestra_agenda") = "S"
      para.muestraagenda = "S"
   Else
      rs("muestra_agenda") = "N"
      para.muestraagenda = "N"
   End If
   
   If Check2 = 1 Then
     c2 = True
   Else
     c2 = False
   End If
   
   If Check3 = 1 Then
     c3 = "S"
   Else
     c3 = "N"
   End If

   rs("imprime_pie_reportes") = c2
   para.imprime_pie_reportes = c2
   rs("imprime_cabecera_reportes") = c3
   rs("tipo_lista_precios") = Val(t_tl)
   para.tipolistaprecios = Val(t_tl)
   rs("tipo_precio_venta") = c_tipoprecio.ListIndex
   rs("tipo_precio_lista") = c_tlp.ListIndex
   rs("estilo_lista_precios") = c_estilo.ListIndex
   rs("estilo_rc") = c_rc.ListIndex
   para.tipoprecioventa = c_tipoprecio.ListIndex
   rs("punto_venta_inicio") = Val(t_sucvta)
   rs.Update
  
  
End If


Set rs = Nothing
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub c_estilo_LostFocus()
If c_estilo.ListIndex < 0 Then
  c_estilo.ListIndex = 0
End If

End Sub

Private Sub c_rc_LostFocus()
If c_rc.ListIndex < 0 Then
  c_rc.ListIndex = 0
End If
End Sub

Private Sub c_tipoprecio_LostFocus()
If c_tipoprecio.ListIndex < 0 Then
  c_tipoprecio.ListIndex = 0
End If
End Sub

Private Sub c_tlp_LostFocus()
If c_tlp.ListIndex < 0 Then
  c_tlp.ListIndex = 0
End If

End Sub

Private Sub Form_Load()
t_id = para.id_usuario
t_nombre = para.usuario
If para.muestraagenda = "S" Then
   Check1 = 1
Else
    Check1 = 0
End If

If para.imprime_pie_reportes Then
  Check2 = 1
Else
  Check2 = 0
End If

If para.imprime_cabecera_reportes = "S" Then
  Check3 = 1
Else
  Check3 = 0
End If

t_tl = para.tipolistaprecios
c_tipoprecio.ListIndex = para.tipoprecioventa
t_sucvta = para.punto_venta_usuario

Set rs = New adodb.Recordset
q = "select * from g1 where [id_usuario] = " & para.id_usuario
rs.MaxRecords = 1
rs.Open q, cn1
c_tlp.ListIndex = rs("tipo_precio_lista")
c_estilo.ListIndex = rs("estilo_lista_precios")
c_rc.ListIndex = rs("estilo_rc")
Set rs = Nothing

End Sub
