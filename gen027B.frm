VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_path 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCION  DE ARCHIVO DIGITAL DE COMPROBANTE ELECTRONICO"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3270
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_origen 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comprobante"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7815
      Begin VB.TextBox t_modulo 
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox t_id 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   "Modulo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "Id."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Seleccionar"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   14
      Top             =   2160
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Archivo Digital seleccionado"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6735
      Begin VB.TextBox t_path 
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   5
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6360
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen027B.frx":0000
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
         Picture         =   "gen027B.frx":0882
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
      Top             =   3015
      Width           =   8010
      _ExtentX        =   14129
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
End
Attribute VB_Name = "gen_path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean
Dim t1 As String



Private Sub btnacepta_Click()
If t_path <> "" Then
J = MsgBox("Confirma Asignar Comprobante Electronico al comprobante registrado en el sistema", 4)
If J = 6 Then
 Select Case t_origen
  Case Is = "C"
    Set rs = New ADODB.Recordset
    q = "select * from a20 where [num_int] = " & Val(t_id)
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    If Not rs.EOF And Not rs.BOF Then
      rs("path") = t_path
    Else
      rs.AddNew
      rs("num_int") = Val(t_id)
      rs("path") = t_path
    End If
    rs.Update
    Set rs = Nothing
    MsgBox ("Comprobante Asociado")
    Unload Me
 
  Case Is = "V"
    Set rs = New ADODB.Recordset
    q = "select * from vta_014 where [num_int] = " & Val(t_id)
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    If Not rs.EOF And Not rs.BOF Then
      rs("path") = t_path
    Else
      rs.AddNew
      rs("num_int") = Val(t_id)
      rs("path") = t_path
    End If
    rs.Update
    Set rs = Nothing
    MsgBox ("Comprobante Asociado")
    Unload Me

 End Select
End If
End If
End Sub


Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub Command1_Click()
s = abrir_archivo_digital(t_path)
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1

End Sub


Private Sub File1_DblClick()
t_path = File1.Path & "\" & File1
End Sub


Function seleccion(filename As String) As Boolean
On Error GoTo err_sel
CommonDialog1.Filter = "Apps (*.pdf)|*.pdf|*.txt|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "pdf"
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

Private Sub Command2_Click()
X = seleccion(t_path)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If

End Sub
Sub camino()

End Sub



Private Sub Form_Load()
Call camino
End Sub

