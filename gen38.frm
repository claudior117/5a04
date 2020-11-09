VERSION 5.00
Begin VB.Form gen_seleccionacarpeta 
   Caption         =   "Selecciona Carpeta:"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8460
   LinkTopic       =   "Form2"
   ScaleHeight     =   4455
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox t_llamada 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox t_carpeta 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3120
      Width           =   5775
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   3720
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "Seleccionar"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "Carpeta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "gen_seleccionacarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case t_llamada
 Case Is = "1"
   gen_citi.t_carpeta = t_carpeta
 Case Is = "2"
   gen_citicom.t_carpeta = t_carpeta
 Case Is = "3"
   gen_asistencia.t_carpeta = t_carpeta
End Select
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Dir1_Change()
Call camino
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
t_llamada = "0"
Call camino
End Sub

Sub camino()
  If Dir1 <> "C:\" Then
    t_carpeta = Dir1 & "\"
  Else
    t_carpeta = Dir1
  End If

End Sub
