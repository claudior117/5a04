VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vet_historia 
   BackColor       =   &H00E0E0E0&
   Caption         =   "HISTORIA CLINICA"
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton Command6 
         Height          =   255
         Left            =   9600
         Picture         =   "vet004A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   11040
         Picture         =   "vet004A.frx":0105
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox c_animal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   7935
      End
      Begin VB.ComboBox C_CLIENTE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Animal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   5535
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "vet004A.frx":020A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "vet004A.frx":0514
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "vet004A.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "vet004A.frx":0B28
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10440
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vet004A.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vet004A.frx":16B4
         Style           =   1  'Graphical
         TabIndex        =   5
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
      TabIndex        =   3
      Top             =   8415
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "04/05/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:50 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9551
      _Version        =   393216
      BackColor       =   12640511
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "vet_historia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim gquery As String

Private Sub btnacepta_Click()
Call limpia
msf1.SetFocus

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub c_animal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub C_cliente_LostFocus()
If C_CLIENTE.ListIndex < 0 Then
  C_CLIENTE.ListIndex = 0
End If
Call carga_mascotas(c_animal, C_CLIENTE.ItemData(C_CLIENTE.ListIndex))
Call armagrid
End Sub

Private Sub Command1_Click()
If para.id_grupo_modulo_actual >= 5 Then
 If c_animal.ListIndex > 0 Then
  vet_historia1!t_funcion = "A"
  Call LLENACAMPOS
  vet_historia1.Show
 End If
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   vet_historia1!t_funcion = "M"
   Call LLENACAMPOS
  End If
 Else
  Call sinpermisos
 End If
End If

Exit Sub
e1:
 Exit Sub
End Sub

Sub LLENACAMPOS()
'On Error GoTo ERROR1
Set rs = New ADODB.Recordset
q = "select * from vet_02 where [id_animal] = " & c_animal.ItemData(c_animal.ListIndex)
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
 vet_historia1.t_id = rs("id_animal")
 vet_historia1.Text3 = rs("nombre")
 vet_historia1.Text4 = rs("descripcion")
 vet_historia1.Show
End If
Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Datos del Animal Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command4_Click()
Call imprime
End Sub

Sub imprime()
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    For i = 2 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE RAZAS", "Tipo Mascota: " & c_tipo, " ", " ", 60, 7, True, False, "H")
      
  End If


End Sub





Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.FixedCols = 1
msf1.SelectionMode = flexSelectionFree
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 1500
msf1.ColWidth(1) = 6000
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 1500
msf1.ColWidth(4) = 1000
msf1.ColWidth(5) = 800
msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Diagnostico Breve"
msf1.TextMatrix(0, 2) = "Peso"
msf1.TextMatrix(0, 3) = "Edad"
msf1.TextMatrix(0, 4) = "Nro. Consulta"

For i = 0 To 5
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(5) = 9 'der
End Sub

Private Sub Command5_Click()
vta_ABM_cli.Show
End Sub

Private Sub Command5_LostFocus()
C_CLIENTE.clear
Call carga_clientes(C_CLIENTE)
C_CLIENTE.ListIndex = 0

End Sub

Private Sub Command6_Click()
vet_ABM_mas.Show
End Sub

Private Sub Command6_LostFocus()
c_animal.clear
Call carga_mascotas(c_animal, C_CLIENTE.ItemData(C_CLIENTE.ListIndex))
c_animal.ListIndex = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
     
End Select
End Sub
Sub limpia()
On Error GoTo err1
Dim q As String
Call armagrid
espere.Show
espere.Label1 = "ESPERE [Leyendo Base de Datos]... "
espere.Refresh

q = "select * from vet_01, vet_02 where vet_01.[id_animal] = vet_02.[id_animal] and vet_01.[id_animal] = " & c_animal.ItemData(c_animal.ListIndex)
q = q & " order by [fecha], [id_consulta]"

Set rs = New ADODB.Recordset
'MsgBox (q)
rs.Open q, cn1
c = 0
While Not rs.EOF
 msf1.AddItem rs("fecha") & Chr$(9) & rs("diag_breve") & Chr$(9) & rs("peso") & Chr$(9) & rs("altura") & Chr$(9) & rs("largo") & Chr$(9) & rs("id_consulta")
 rs.MoveNext
 c = c + 1
Wend
 msf1.AddItem ""
 msf1.AddItem "" & Chr$(9) & "Total de Registros : " & c

Set rs = Nothing
Call INICIALIZA2(vet_historia1)
Unload espere

err1:
 Unload espere
 Exit Sub

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  
  End Select
End Sub



Private Sub Form_Load()
Call barraesag(Me)
Load vet_historia1
Call carga_clientes(C_CLIENTE)
C_CLIENTE.ListIndex = 0
Call carga_mascotas(c_animal, C_CLIENTE.ItemData(C_CLIENTE.ListIndex))
Call armagrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vet_historia1
End Sub

Private Sub msf1_GotFocus()
StatusBar1.Panels.Item(2) = "[F4] Saca - [F7] Imprime "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
 If msf1.Rows > 0 Then
   msf1.RemoveItem msf1.Row
  End If
End If


If KeyCode = vbKeyF7 Then
 If msf1.Rows > 0 Then
   
   Call imprime
    
 End If
End If

End Sub



