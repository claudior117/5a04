VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_prov1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "BUSCADOR DE PROVEEDORES"
   ClientHeight    =   3960
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3960
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox T_PROVINCIA 
      Height          =   405
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   10
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox T_LOCALIDAD 
      Height          =   405
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   9
      Top             =   600
      Width           =   4815
   End
   Begin VB.TextBox T_CP 
      Height          =   405
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   5040
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Arch001B.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Buscar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Arch001B.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.TextBox t_DESCPROD 
      Height          =   405
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3705
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "18/12/2005"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:58 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PROVINCIA"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COD. POSTAL"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RAZON SOCIAL"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOCALIDAD"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "abm_prov1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private INDICE As Integer



Private Sub btnacepta_Click()
Call imprimir
Unload Me
End Sub

Sub imprimir()
Dim QUERY As String
Set rs = New ADODB.Recordset
QUERY = "SELECT * FROM P01   "
CONECTOR = " WHERE "

If t_DESCPROD <> "" Then
   QUERY = QUERY & CONECTOR & " [DENOMINACION] LIKE " & "'%" & t_DESCPROD & "%'"
   CONECTOR = " AND "
End If

If T_LOCALIDAD <> "" Then
   QUERY = QUERY & CONECTOR & " [LOCALIDAD] LIKE " & "'%" & T_LOCALIDAD & "%'"
   CONECTOR = " AND "
End If

If T_CP <> "" Then
   QUERY = QUERY & CONECTOR & " [CP] LIKE " & "'%" & T_CP & "%'"
   CONECTOR = " AND "
End If

If T_PROVINCIA <> "" Then
   QUERY = QUERY & CONECTOR & " [PROVINCIA] LIKE " & "'%" & T_PROVINCIA & "%'"
   CONECTOR = " AND "
End If


QUERY = QUERY & " ORDER BY [DENOMINACION]"

Call conectaradodc(pabmprov0!Adodc1, QUERY, cn1)


End Sub
Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter(Me, 3)
  
End Select


End Sub

Private Sub Form_Load()
  
'BARRA ESTADO
StatusBar1.Panels.Item(1) = ""
StatusBar1.Panels.Item(2) = "[ENTER] Avanza - [Up] Regresa - [ESC] Termina"


End Sub


Private Sub t_DESCPROD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call imprimir
End If
End Sub

