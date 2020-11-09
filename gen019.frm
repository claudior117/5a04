VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GEN_ABMCAMION 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CAMIONES"
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
      Caption         =   "Filtrar"
      Height          =   1335
      Left            =   5760
      TabIndex        =   9
      Top             =   0
      Width           =   6375
      Begin VB.TextBox t_chofer 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   16
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox t_unidad 
         Height          =   315
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   14
         Top             =   600
         Width           =   5055
      End
      Begin VB.ComboBox C_CLIENTE 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C000C0&
         Caption         =   "Chofer:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C000C0&
         Caption         =   "Unidad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C000C0&
         Caption         =   "Transporte:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "gen019.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "gen019.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "gen019.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "gen019.frx":091E
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
      Left            =   10320
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen019.frx":0C28
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
         Picture         =   "gen019.frx":14AA
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
            TextSave        =   "09/08/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:21 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10186
      _Version        =   393216
      BackColor       =   12640511
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "GEN_ABMCAMION"
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



Private Sub C_cliente_LostFocus()
If C_CLIENTE.ListIndex < 0 Then
  C_CLIENTE.ListIndex = 0
End If
Call armagrid
End Sub

Private Sub Command1_Click()
If para.id_grupo_modulo_actual >= 5 Then
 gen_abmcamion1!t_funcion = "A"
 gen_abmcamion1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
'On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   gen_abmcamion1!t_funcion = "M"
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
q = "select * from a17 where [id_camion] = " & Val(msf1.TextMatrix(msf1.Row, 0))
 rs.Open q, cn1
 gen_abmcamion1!t_id = rs("id_camion")
 gen_abmcamion1!t_descripcion = rs("camion")
 gen_abmcamion1!t_dominio = rs("dominio")
 gen_abmcamion1!t_dominioa = rs("dominio_acoplado")
 gen_abmcamion1.c_tipo.ListIndex = buscaindice(gen_abmcamion1.c_tipo, rs("id_transporte"))
 gen_abmcamion1!t_chofer = rs("chofer")
 gen_abmcamion1!t_dni = rs("dni")
 gen_abmcamion1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Camiones. Proc.: LLENACAMPOS")
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
    Call imprimegrid(msf1, c(), "LISTADO DE CAMIONES", "Transporte: " & c_tipo, " ", " ", 60, 7, True, False, "H")
      
  End If


End Sub





Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
msf1.FixedCols = 1
msf1.SelectionMode = flexSelectionFree
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 4000
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 2500
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 2000
msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Camion"
msf1.TextMatrix(0, 2) = "Dominio"
msf1.TextMatrix(0, 3) = "Dom. Acop."
msf1.TextMatrix(0, 4) = "Chofer"
msf1.TextMatrix(0, 5) = "DNI"
msf1.TextMatrix(0, 6) = "Transporte"

For i = 0 To 5
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(5) = 9 'der
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

q = "select * from a17, a1 where [id_transporte] = [id_proveedor] "
c = " and "
If C_CLIENTE.ListIndex > 0 Then
  q = q & c & " [id_transporte] = " & C_CLIENTE.ItemData(C_CLIENTE.ListIndex)
End If

If t_unidad <> "" Then
  q = q & c & " [camion] like '%" & t_unidad & "%'"
End If

If t_chofer <> "" Then
  q = q & c & " [chofer] like '%" & t_chofer & "%'"
End If


Set rs = New ADODB.Recordset
rs.Open q, cn1
c = 0
While Not rs.EOF
 msf1.AddItem rs("id_camion") & Chr$(9) & rs("camion") & Chr$(9) & rs("dominio") & Chr$(9) & rs("dominio_acoplado") & Chr$(9) & rs("chofer") & Chr$(9) & rs("dni") & Chr$(9) & rs("denominacion")
 rs.MoveNext
 c = c + 1
Wend
 msf1.AddItem ""
 msf1.AddItem "" & Chr$(9) & "Total de Registros : " & c

Set rs = Nothing
Call INICIALIZA2(gen_abmcamion1)
Unload espere

err1:
 Unload espere
 Exit Sub

End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load gen_abmcamion1
Call carga_transporte(C_CLIENTE)
C_CLIENTE.AddItem "<Todos>", 0
C_CLIENTE.ListIndex = 0
Call armagrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload gen_abmcamion1
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



Private Sub t_chofer_GotFocus()
t_chofer = ""
End Sub

Private Sub t_unidad_GotFocus()
t_unidad = ""
End Sub
