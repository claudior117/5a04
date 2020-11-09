VERSION 5.00
Begin VB.Form gen_cambiopassword 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAMBIO DE PASSWORD"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   3600
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen002.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen002.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
      Begin VB.TextBox t_pn2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   14
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox t_pa 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_pn 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   14
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C000C0&
         Caption         =   "Confirma Password Nueva:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C000C0&
         Caption         =   "Password Actual:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C000C0&
         Caption         =   "Password Nueva:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   5055
      Begin VB.ComboBox c_usuarios 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "gen_cambiopassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
  Dim cat As ADOX.Catalog
  Dim usr As ADOX.User
  Dim t As String
On Error GoTo errp
J = MsgBox("Confirma Cambio de Password", 4)
If J = 6 Then
   If t_pn = t_pn2 Then
      t = c_usuarios
      Set cat = New ADOX.Catalog
      cat.ActiveConnection = cn1
      Set usr = cat.Users(t)
          usr.ChangePassword t_pa, t_pn
      MsgBox "Su Password ha sido Cambiada!!!!"
      Unload Me
   Else
     MsgBox ("No coinciden los datos ingresados")
   End If
 End If
Exit Sub

errp:
MsgBox ("¡¡¡¡Error.No se pudo realizar el cambio solicitado!!!!!")
Unload Me

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub Form_Activate()
'FIXIT: c_usuarios.Locked property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
c_usuarios.Locked = False
c_usuarios.ListIndex = buscaindice(c_usuarios, para.id_usuario)
If para.id_usuario <> 9 Then
'FIXIT: c_usuarios.Locked property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
  c_usuarios.Locked = True
End If
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
