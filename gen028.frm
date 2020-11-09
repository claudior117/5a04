VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form gen_factapocrifas 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PROCESO DE ACTUALIZACION DEL PADRON DE FACTURAS APOCRIFAS"
   ClientHeight    =   4890
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   7995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Link"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar"
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Archivo de facturas apócrifas descargado del AFIP"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6735
      Begin VB.TextBox t_path 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6120
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen028.frx":0000
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
         Picture         =   "gen028.frx":0882
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
      Top             =   4635
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "3) Ejecute el proceso de actualiacion"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "2) Utilice el boton SELECCIONAR para buscarlo."
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label Label4 
      Caption         =   "1) Descargue el padron haciendo click en LINK  y guardelo en su maquina."
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "PADRON DE FACTURAS APOCRIFAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   5535
   End
End
Attribute VB_Name = "gen_factapocrifas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean
Dim t1 As String



Private Sub btnacepta_Click()
Dim l As String
If verifica Then
 J = MsgBox("Confirma Actualizacion del Padron de Factutas Apócrifas ", 4)
 If J = 6 Then
  'borrar bd.actual
  'If abrirconexionib Then
  
   Label5 = "Borrando padron actual.."
   t = 0
   Set rs = New ADODB.Recordset
   q = "select * from fa"
   rs.Open q, cnib, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
    t = t + 1
    Label6 = t
    Label6.Refresh
    rs.Delete
    rs.MoveNext
   Wend
   Set rs = Nothing
 
 
 
   Label5 = "Importando facturs Apócrifas.."
   t = 0
   Set rs = New ADODB.Recordset
   q = "select * from fa"
   rs.Open q, cnib, adOpenDynamic, adLockOptimistic
  
   Open t_path For Input As #1
   While Not EOF(1)
    Label6 = t
    Label6.Refresh
    Line Input #1, l
    If Val(Mid$(l, 1, 11)) > 1000 Then
        t = t + 1
        rs.AddNew
         rs("cuit") = Val(Mid$(l, 1, 11))
         rs("fecha_deteccion") = Mid$(l, 13, 10)
         rs("fecha_publicacion") = Mid$(l, 24, 10)
        rs.Update
       rs.MoveNext
    End If
   Wend
   Set rs = Nothing
   Close #1
    
  Else
    MsgBox ("Error al abrir B.D. de Padron de Embargados IB")
  End If
 ' cnib.Close
 'End If
End If

Label5 = "Fin"

End Sub
Function verifica() As Boolean
On Error GoTo errorib
verifica = True
Open t_path For Input As #1
Line Input #1, l
Close #1
Exit Function
errorib:
  MsgBox ("Archivo Inexistente o Invalido")
  verifica = False
  Close #1
  Exit Function
  
End Function
Private Sub btnsale_Click()
Unload Me
End Sub




Private Sub Command1_Click()
x = seleccion(t_path)
End Sub

Private Sub Command2_Click()
Dim intobj As Object
Set intobj = CreateObject("InternetExplorer.Application")
intobj.Visible = -1
intobj.Navigate "http://www.afip.gob.ar/genericos/facturasApocrifas/download/facacop.txt"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If

End Sub
Function seleccion(filename As String) As Boolean
On Error GoTo err_sel
CommonDialog1.Filter = "Apps *.txt"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Selecciona Archivo"
CommonDialog1.InitDir = "C:\"
CommonDialog1.filename = filename
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
t_path = filename

Exit Function
err_sel:
t_path = filename
End Function




