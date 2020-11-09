VERSION 5.00
Begin VB.Form gen_links 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LINKS"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   1740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.Image Image5 
         Height          =   480
         Left            =   240
         Picture         =   "gen007.frx":0000
         Top             =   2640
         Width           =   1110
      End
      Begin VB.Image Image4 
         Height          =   450
         Left            =   240
         Picture         =   "gen007.frx":0729
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   360
         Picture         =   "gen007.frx":0D78
         Top             =   1440
         Width           =   885
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   360
         Picture         =   "gen007.frx":1359
         Top             =   720
         Width           =   765
      End
      Begin VB.Image Image1 
         Height          =   330
         Left            =   240
         Picture         =   "gen007.frx":190A
         Top             =   240
         Width           =   1155
      End
   End
End
Attribute VB_Name = "gen_links"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Unload Me
End If
End Sub

Private Sub Image1_Click()
'FIXIT: Declare 'intobj' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim intobj As Object
Set intobj = CreateObject("InternetExplorer.Application")
intobj.Visible = -1
intobj.Navigate "http://www.afip.gov.ar"

End Sub

Private Sub Image2_Click()
'FIXIT: Declare 'intobj' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim intobj As Object
Set intobj = CreateObject("InternetExplorer.Application")
intobj.Visible = -1
intobj.Navigate "http://www.arba.gov.ar"
'Do Until intobj.busy = False
'   Loop

End Sub

Private Sub Image3_Click()
'FIXIT: Declare 'intobj' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim intobj As Object
Set intobj = CreateObject("InternetExplorer.Application")
intobj.Visible = -1
intobj.Navigate "http://www.oncca.gov.ar"
'Do Until intobj.busy = False
'   Loop

End Sub

Private Sub Image4_Click()
'FIXIT: Declare 'intobj' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim intobj As Object
Set intobj = CreateObject("InternetExplorer.Application")
intobj.Visible = -1
intobj.Navigate "http://www.bolsadecereales.com"
'Do Until intobj.busy = False
'   Loop

End Sub

Private Sub Image5_Click()
'FIXIT: Declare 'intobj' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim intobj As Object
Set intobj = CreateObject("InternetExplorer.Application")
intobj.Visible = -1
intobj.Navigate "http://www.dolarsi.com"
'Do Until intobj.busy = False
'   Loop

End Sub
