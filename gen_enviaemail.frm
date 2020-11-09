VERSION 5.00
Begin VB.Form gen_enviaemail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio de email"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_path 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox t_email 
      Height          =   405
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Direccion de Correo"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "gen_enviaemail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo err1
Set rs = New ADODB.Recordset
q = "Select * from fe_01 where id = 1"
rs.Open q, cn1
  servidor = rs("servidor_email")
  usuario = rs("usuario_email")
  clave = rs("pass_email")

  Dim PyEmail As Object
    
  Set PyEmail = CreateObject("PyEmail")
  ok = PyEmail.Conectar(servidor, usuario, clave)

  ' Envio el o los correos (repetir por cada FE)
  remitente = rs("email_remite")
  destinatario = t_email
  mensaje = "Buenos días, adjuntamos factura electrónica"
  archivo = t_path

 ok = PyEmail.Enviar(remitente, motivo, destinatario, mensaje, archivo)
 
 Unload Me

 Exit Sub
err1:
  MsgBox ("Error!! Correo sin enviar")
  Exit Sub
  
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

