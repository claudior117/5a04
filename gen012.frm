VERSION 5.00
Begin VB.Form gen_resguardo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resguardo de Archivos"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   3015
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Resguradar"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unidad o Carpeta destino"
      Height          =   2415
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Para realizar el backup seleccione la Unidad y carpeta  destino y presione Resguardar. RESGUARDE SUS ARCHIVOS PERIODICAMENTE!!!!"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "gen_resguardo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
J = MsgBox("Cierre el sistema en todas las terminales de la red y confirme", 4)
If J = 6 Then
  On Error GoTo errbak
  espere.Show
  espere.Label1 = "Espere... Resguardando el sistema"
  espere.Refresh
  Dim resguardo As New Scripting.FileSystemObject
  resguardo.CopyFolder App.Path, Dir1, True
  Unload espere
  MsgBox ("Operacion terminada con Exito!!!")
End If
Exit Sub
errbak:
 MsgBox ("¡¡¡Error!!! El Resguardo no pudo realizarse")
 End
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
 On Error GoTo ERRDSK
 Dir1.Path = Drive1.Drive

Exit Sub
ERRDSK:
 MsgBox ("Error en la Unidad Seleccionada")
 Drive1.Drive = "C:"
 Dir1.Path = Drive1

End Sub

Private Sub Form_Load()
Drive1.Drive = "C:"
Dir1.Path = Drive1 & "\"

End Sub
