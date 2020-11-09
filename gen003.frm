VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form gen_cambioseguridad 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAMBIO DE NIVEL DE SEGURIDAD POR MODULO"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4800
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen003.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen003.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   4095
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
      Begin VB.TextBox t_stock 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   29
         Top             =   3600
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox t_contabilidad 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   19
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox t_produccion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   17
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox t_productos 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   15
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox t_bancos 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox t_caja 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox t_Ventas 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox t_compras 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   1
         Top             =   720
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown5 
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown6 
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   2640
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown7 
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   3120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown8 
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   3600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C000C0&
         Caption         =   "Stock:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C000C0&
         Caption         =   "Contabilidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C000C0&
         Caption         =   "Produccion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C000C0&
         Caption         =   "Productos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C000C0&
         Caption         =   "Bancos:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C000C0&
         Caption         =   "Caja:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C000C0&
         Caption         =   "Ventas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C000C0&
         Caption         =   "Compras:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton Command1 
         Caption         =   "Permisos"
         Height          =   255
         Left            =   5040
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox c_usuarios 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "gen_cambioseguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private Sub btnacepta_Click()
On Error GoTo errp
J = MsgBox("Confirma Cambio de Permisos para el usuario", 4)
If J = 6 Then
  If Frame2.Visible = True Then
    Set rs = New ADODB.Recordset
    q = "select * from g1 where [id_usuario] = " & c_usuarios.ItemData(c_usuarios.ListIndex)
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    If Not rs.BOF And Not rs.EOF Then
      p = t_Ventas & t_compras & t_caja & t_bancos & t_productos & t_produccion & t_contabilidad & t_stock
      rs("grupo") = p
      rs.Update
      Frame2.Visible = False
      MsgBox ("Los permisos han sido actualizados. Salga del sistema y vuelva a ingresar para que los cambios se hagan efectivos!!!!")
    Else
      MsgBox ("El usuario no tiene cuenta asignada")
    End If
    Set rs = Nothing
  End If
End If
Exit Sub

errp:
MsgBox ("¡¡¡¡Error.No se pudo modificar los permisos!!!!!")
Unload Me

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub c_usuarios_LostFocus()
If c_usuarios.ListIndex < 0 Then
  c_usuarios.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
Set rs = New ADODB.Recordset
q = "select * from g1 where [id_usuario] = " & c_usuarios.ItemData(c_usuarios.ListIndex)
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  p = rs("grupo")
  t_Ventas = Val(Mid$(p, 1, 1))
  t_compras = Val(Mid$(p, 2, 1))
  t_caja = Val(Mid$(p, 3, 1))
  t_bancos = Val(Mid$(p, 4, 1))
  t_productos = Val(Mid$(p, 5, 1))
  t_produccion = Val(Mid$(p, 6, 1))
  t_contabilidad = Val(Mid$(p, 7, 1))
  t_stock = Val(Mid$(p, 8, 1))
  
  Frame2.Visible = True
Else
  MsgBox ("El usuario no tiene cuenta asignada")
End If
Set rs = Nothing
End Sub

Private Sub Form_Activate()
c_usuarios.ListIndex = buscaindice(c_usuarios, para.id_usuario)
Call nivel_acceso(1)
'utilizo el nivel 9 de ventas como administrador general del sistema
If para.id_grupo_modulo_actual <> 9 Then
  c_usuarios.Locked = True
  Frame1.Enabled = False
  Frame2.Enabled = False
Else
  Frame1.Enabled = True
  Frame2.Enabled = True
'FIXIT: c_usuarios.Locked property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c_usuarios.Locked = False
End If
Frame2.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 2)
End If

End Sub

Private Sub Form_Load()
Call carga_usuarios(c_usuarios)
End Sub

Private Sub t_pa_GotFocus()
t_pa = ""
End Sub

Private Sub t_pn_GotFocus()
t_pn = ""
End Sub

Private Sub t_pn2_GotFocus()
t_pn2 = ""
End Sub

Private Sub t_pn2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub UpDown1_DownClick()
If Val(t_Ventas) > 0 Then
  t_Ventas = Val(t_Ventas) - 1
End If

End Sub

Private Sub UpDown1_UpClick()
If Val(t_Ventas) < 9 Then
  t_Ventas = Val(t_Ventas) + 1
End If
End Sub

Private Sub UpDown2_DownClick()
If Val(t_compras) > 0 Then
  t_compras = Val(t_compras) - 1
End If

End Sub

Private Sub UpDown2_UpClick()
If Val(t_compras) < 9 Then
  t_compras = Val(t_compras) + 1
End If

End Sub

Private Sub UpDown3_DownClick()
If Val(t_caja) > 0 Then
  t_caja = Val(t_caja) - 1
End If

End Sub

Private Sub UpDown3_UpClick()
If Val(t_caja) < 9 Then
  t_caja = Val(t_caja) + 1
End If

End Sub

Private Sub UpDown4_DownClick()
If Val(t_bancos) > 0 Then
  t_bancos = Val(t_bancos) - 1
End If

End Sub

Private Sub UpDown4_UpClick()
If Val(t_bancos) < 9 Then
  t_bancos = Val(t_bancos) + 1
End If

End Sub

Private Sub UpDown5_DownClick()
If Val(t_productos) > 0 Then
  t_productos = Val(t_productos) - 1
End If

End Sub

Private Sub UpDown5_UpClick()
If Val(t_productos) < 9 Then
  t_productos = Val(t_productos) + 1
End If

End Sub

Private Sub UpDown6_DownClick()
If Val(t_produccion) > 0 Then
  t_produccion = Val(t_produccion) - 1
End If

End Sub

Private Sub UpDown6_UpClick()
If Val(t_produccion) < 9 Then
  t_produccion = Val(t_produccion) + 1
End If

End Sub

Private Sub UpDown7_DownClick()
If Val(t_contabilidad) > 0 Then
  t_contabilidad = Val(t_contabilidad) - 1
End If

End Sub

Private Sub UpDown7_UpClick()
If Val(t_contabilidad) < 9 Then
  t_contabilidad = Val(t_contabilidad) + 1
End If

End Sub

Private Sub UpDown8_DownClick()
If Val(t_stock) > 0 Then
  t_stock = Val(t_stock) - 1
End If

End Sub

Private Sub UpDown8_UpClick()
If Val(t_stock) < 9 Then
  t_stock = Val(t_stock) + 1
End If
End Sub
