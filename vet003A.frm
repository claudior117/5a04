VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vet_ABM_raza 
   BackColor       =   &H00E0E0E0&
   Caption         =   "RAZAS MASCOTAS"
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   4455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id."
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Raza"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1095
      Left            =   7200
      TabIndex        =   9
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox C_TIPO 
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "vet003A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "vet003A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "vet003A.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "vet003A.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vet003A.frx":0C28
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
         Picture         =   "vet003A.frx":14AA
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
            TextSave        =   "06/04/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:27 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10186
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "vet_ABM_raza"
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



Private Sub C_TIPO_LostFocus()
If C_TIPO.ListIndex < 0 Then
  C_TIPO.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
If para.id_grupo_modulo_actual >= 5 Then
 vet_abm_raza1!t_funcion = "A"
 vet_abm_raza1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   vet_abm_raza1!t_funcion = "M"
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
On Error GoTo ERROR1
Set rs = New ADODB.Recordset
q = "select * from vet_04 where [id_raza] = " & Val(msf1.TextMatrix(msf1.Row, 0))
 rs.Open q, cn1
 vet_abm_raza1!t_id = rs("id_raza")
 vet_abm_raza1!t_descripcion = rs("raza")
 vet_abm_raza1.C_TIPO.ListIndex = buscaindice(vet_abm_raza1.C_TIPO, rs("id_tipo"))
 vet_abm_raza1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Razas. Proc.: LLENACAMPOS")
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
    Call imprimegrid(msf1, c(), "LISTADO DE RAZAS", "Tipo Mascota: " & C_TIPO, " ", " ", 60, 7, True, False, "H")
      
  End If


End Sub





Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 3
msf1.FixedCols = 1
msf1.SelectionMode = flexSelectionFree
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 1500
msf1.ColWidth(1) = 6000
msf1.ColWidth(2) = 3000
msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Raza"
msf1.TextMatrix(0, 2) = "Tipo Mascota"


For i = 1 To 2
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(0) = 9 'der
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
     
End Select
End Sub
Sub limpia()
Dim q As String
Call armagrid
espere.Show
espere.Label1 = "ESPERE [Leyendo Base de Datos]... "
espere.Refresh

q = "select * from vet_04, vet_03 where vet_04.[id_tipo] = vet_03.[id_tipo]"
c = " and "
If C_TIPO.ListIndex > 0 Then
 q = q & c & " vet_04.[id_tipo] = " & C_TIPO.ItemData(C_TIPO.ListIndex)
 c = " and "
End If

If Option2 = True Then
  q = q & " order by [raza]"
Else
   q = q & " order by [id_raza]"
End If

Set rs = New ADODB.Recordset
'MsgBox (q)
rs.Open q, cn1
c = 0
While Not rs.EOF
 msf1.AddItem rs("id_raza") & Chr$(9) & rs("raza") & Chr$(9) & rs("tipo")
 rs.MoveNext
 c = c + 1
Wend
 msf1.AddItem ""
 msf1.AddItem "" & Chr$(9) & "Total de Registros : " & c

Set rs = Nothing
Call INICIALIZA2(vet_abm_raza1)
Unload espere
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load vet_abm_raza1
Call carga_tipo_mascota(C_TIPO)
C_TIPO.AddItem "<Todos>", 0
C_TIPO.ListIndex = 0

Option2 = True
Call armagrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vet_abm_raza1
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



