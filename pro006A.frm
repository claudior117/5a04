VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form pro_ABM_pieza 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ABM PIEZAS ESTRUCTURA DE PRODUCCION"
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
      TabIndex        =   13
      Top             =   7320
      Width           =   4455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id."
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DETALLE"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   735
      Left            =   7200
      TabIndex        =   9
      Top             =   120
      Width           =   4575
      Begin VB.TextBox t_prov 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Detalle:"
         Height          =   255
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
      Width           =   5535
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "pro006A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "pro006A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "pro006A.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "pro006A.frx":091E
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
         Picture         =   "pro006A.frx":0C28
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
         Picture         =   "pro006A.frx":14AA
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
            TextSave        =   "22/02/2011"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:56 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10186
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "pro_ABM_pieza"
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




Private Sub Command1_Click()
If para.id_grupo_modulo_actual >= 5 Then
 pro_abm_pieza1!t_funcion = "A"
 pro_abm_pieza1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   pro_abm_pieza1!t_funcion = "M"
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
q = "select * from pro_06 where [id_pieza] = " & Val(msf1.TextMatrix(msf1.Row, 0))
rs.Open q, cn1
 pro_abm_pieza1!t_id = rs("id_pieza")
 pro_abm_pieza1!t_descripcion = rs("descripcion")
 pro_abm_pieza1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Piezas. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 7 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   pro_abm_pieza1!t_funcion = "B"
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
    For i = 4 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE PIEZAS", "", "Detalle: " & t_detalle, " ", 60, 7, True, False, "H")
      
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
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 10000
msf1.ColWidth(2) = 2000
msf1.TextMatrix(0, 0) = "Id. Pieza"
msf1.TextMatrix(0, 1) = "Pieza"
msf1.TextMatrix(0, 2) = ""

For i = 1 To 2
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(0) = 9 'der
End Sub

Sub limpia()
Dim q As String
Call armagrid
espere.Show
espere.Label1 = "ESPERE [Leyendo Base de Datos]... "
espere.Refresh

q = "select * from pro_06 "
c = " where "
If t_prov <> "" Then
 q = q & c & " [descripcion] like '%" & t_prov & "%'"
 c = " and "
End If

If Option2 = True Then
  q = q & " order by [descripcion]"
Else
   q = q & " order by [id_pieza]"
End If

Set rs = New ADODB.Recordset
rs.Open q, cn1
c = 0
While Not rs.EOF
 msf1.AddItem rs("id_pieza") & Chr$(9) & rs("descripcion")
 rs.MoveNext
 c = c + 1
Wend
msf1.AddItem ""
msf1.AddItem "" & Chr$(9) & "Total de Registros : " & c

Set rs = Nothing
Call INICIALIZA2(pro_abm_pieza1)
Unload espere
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load pro_abm_pieza1
Option2 = True
Call armagrid
Load pro_estructura
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload pro_abm_pieza1
Unload pro_estructura
End Sub

Private Sub msf1_GotFocus()
StatusBar1.Panels.Item(2) = "[F7] Imprime - [F5] Muestra estructura "

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

If KeyCode = vbKeyF5 Then
If msf1.Rows > 0 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
     pro_estructura.c_prov.ListIndex = buscaindice(pro_estructura.c_prov, Val(msf1.TextMatrix(msf1.Row, 0)))
     pro_estructura.Show
  End If
End If


End If

End Sub

Private Sub t_prov_GotFocus()
t_prov = ""
End Sub


