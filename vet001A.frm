VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form vet_ABM_mas 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MASCOTAS"
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Maestros Relacionados"
      Height          =   735
      Left            =   4800
      TabIndex        =   20
      Top             =   7320
      Width           =   4575
      Begin VB.CommandButton Command6 
         Caption         =   "Razas"
         Height          =   375
         Left            =   2400
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Tipos Mascotas"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   7320
      Width           =   4455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id."
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1335
      Left            =   7200
      TabIndex        =   9
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox C_tipo 
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   600
         Width           =   2895
      End
      Begin VB.ComboBox C_vend 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox t_prov 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo:"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "vet001A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "vet001A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "vet001A.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "vet001A.frx":091E
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
         Picture         =   "vet001A.frx":0C28
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
         Picture         =   "vet001A.frx":14AA
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
            TextSave        =   "06:31 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   15
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
Attribute VB_Name = "vet_ABM_mas"
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


Private Sub c_vend_LostFocus()
If C_vend.ListIndex < 0 Then
   C_vend.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
If para.id_grupo_modulo_actual >= 5 Then
 vet_abm_mas1!t_funcion = "A"
 vet_abm_mas1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
'On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
   vet_abm_mas1!t_funcion = "M"
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
Set rs2 = New ADODB.Recordset
q = "select * from vet_02 where [id_animal] = " & Val(msf1.TextMatrix(msf1.Row, 0))
rs2.Open q, cn1
 vet_abm_mas1!t_id = rs2("id_animal")
 vet_abm_mas1!t_nombre = rs2("nombre")
 vet_abm_mas1!t_descripcion = rs2("descripcion")
 vet_abm_mas1!t_fechanac = rs2("fecha_nac")
 vet_abm_mas1!c_cliente.ListIndex = buscaindice(vet_abm_mas1.c_cliente, rs2("id_cliente"))
 Call carga_raza(vet_abm_mas1.c_raza, rs2("id_tipo"))
 vet_abm_mas1!C_TIPO.ListIndex = buscaindice(vet_abm_mas1.C_TIPO, rs2("id_tipo"))
 vet_abm_mas1!c_raza.ListIndex = buscaindice(vet_abm_mas1.c_raza, rs2("id_raza"))
 If rs2("sexo") = "H" Then
   vet_abm_mas1!c_sexo.ListIndex = 0
 Else
   vet_abm_mas1!c_sexo.ListIndex = 1
 End If
 vet_abm_mas1.Show

Set rs2 = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Mascota. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 7 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   vet_abm_mas1!t_funcion = "B"
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

Private Sub Command4_Click()
Call imprime
End Sub

Sub imprime()
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5
    c(6) = 6
    c(7) = 7
    For i = 8 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE MASCOTAS", "Tipo:" & C_TIPO, "Nombre: " & t_prov, "Dueño: " & C_vend, 60, 7, True, False, "H")
      
  End If


End Sub








Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.FixedCols = 2
msf1.SelectionMode = flexSelectionFree
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 600
msf1.ColWidth(1) = 2500
msf1.ColWidth(2) = 2000
msf1.ColWidth(3) = 2500
msf1.ColWidth(4) = 1500
msf1.ColWidth(5) = 600
msf1.ColWidth(6) = 2000
msf1.ColWidth(7) = 2000
msf1.ColWidth(8) = 500
msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Cliente"
msf1.TextMatrix(0, 2) = "Nombre Mascota"
msf1.TextMatrix(0, 3) = "Descipcion"
msf1.TextMatrix(0, 4) = "Fecha Nac."
msf1.TextMatrix(0, 5) = "Sexo"
msf1.TextMatrix(0, 6) = "Tipo"
msf1.TextMatrix(0, 7) = "Raza"
msf1.TextMatrix(0, 8) = ""

For i = 2 To 8
  msf1.ColAlignment(i) = 9 'der
Next i
msf1.ColAlignment(0) = 1 'izq
End Sub

Private Sub Command5_Click()
vet_ABM_tipo.Show
End Sub

Private Sub Command6_Click()
vet_ABM_raza.Show
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

q = "select * from vet_02, vta_01, vet_03, vet_04 where vet_02.[id_cliente] = vta_01.[id_cliente] and vet_02.[id_tipo] = vet_03.[id_tipo] and vet_02.[id_raza] = vet_04.[id_raza]"
c = " and "
If t_prov <> "" Then
 q = q & c & " [nombre] like '%" & t_prov & "%'"
 c = " and "
End If

If C_vend.ListIndex > 0 Then
 q = q & c & " vet_02.[id_cliente] =" & C_vend.ItemData(C_vend.ListIndex)
 c = " and "
End If

If C_TIPO.ListIndex > 0 Then
 q = q & c & " vet_02.[id_tipo] =" & C_TIPO.ItemData(C_TIPO.ListIndex)
 c = " and "
End If

If Option2 = True Then
  q = q & " order by [denominacion]"
Else
   q = q & " order by [id_animal]"
End If

Set rs = New ADODB.Recordset
rs.Open q, cn1
c = 0
While Not rs.EOF
 msf1.AddItem rs("id_animal") & Chr$(9) & rs("denominacion") & Chr$(9) & rs("nombre") & Chr$(9) & rs("descripcion") & Chr$(9) & rs("fecha_nac") & Chr$(9) & rs("sexo") & Chr$(9) & rs("tipo") & Chr$(9) & rs("raza")
 rs.MoveNext
 c = c + 1
Wend
 msf1.AddItem ""
 msf1.AddItem "" & Chr$(9) & "Total de Registros : " & c

Set rs = Nothing
Call INICIALIZA2(vet_abm_mas1)
Unload espere
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load vet_abm_mas1

Call carga_clientes(C_vend)
C_vend.AddItem "<Todos>", 0
C_vend.ListIndex = 0

Call carga_tipo_mascota(C_TIPO)
C_TIPO.AddItem "<Todos>", 0
C_TIPO.ListIndex = 0

Option2 = True
Call armagrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vet_abm_mas1
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

Private Sub t_prov_GotFocus()
t_prov = ""
End Sub


