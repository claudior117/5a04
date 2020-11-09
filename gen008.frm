VERSION 5.00
Begin VB.Form gen_tools 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TOOLS"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tools(F12)"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton Command5 
         Caption         =   "Definir Imp."
         Height          =   855
         Left            =   4560
         Picture         =   "gen008.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Calculadora"
         Height          =   855
         Left            =   2400
         Picture         =   "gen008.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton command1 
         Caption         =   "Agenda"
         Height          =   855
         Left            =   240
         Picture         =   "gen008.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Calendario"
         Height          =   855
         Left            =   1320
         Picture         =   "gen008.frx":0A8B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Links"
         Height          =   855
         Left            =   3480
         Picture         =   "gen008.frx":0ED1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "gen_tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


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
  x = Shell(App.Path & "\tools\calc.exe", vbNormalFocus)
End Sub

Private Sub Command5_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Unload Me
End If
End Sub

