VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form cgr_plancuentas 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PLAN DE CUENTAS"
   ClientHeight    =   8835
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "CGR001.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "CGR001.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "CGR001.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "CGR001.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Width           =   8535
      Begin VB.ComboBox c_nivel3 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Text            =   "c_nivel3"
         Top             =   960
         Width           =   4095
      End
      Begin VB.ComboBox c_nivel2 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Text            =   "c_nivel2"
         Top             =   600
         Width           =   4095
      End
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "CGR001.frx":0C28
         Left            =   6240
         List            =   "CGR001.frx":0C35
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox c_nivel1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nivel 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nivel 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tipo :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nivel 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR001.frx":0C64
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
         Picture         =   "CGR001.frx":14E6
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
      Top             =   8580
      Width           =   12135
      _ExtentX        =   21405
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
            TextSave        =   "04/04/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:08 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_plancuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer




Private Sub btnacepta_Click()
Call armar

End Sub
Sub armar()
Call armagrid
Set rs = New ADODB.Recordset
q = "select * from c_01"
c = " where "
If c_nivel1.ListIndex > 0 Then
  q = q & c & " [pos1] = " & c_nivel1.ItemData(c_nivel1.ListIndex)
  c = " and "
End If

If c_nivel2.ListIndex > 0 Then
  q = q & c & " [pos2] = " & c_nivel2.ItemData(c_nivel2.ListIndex)
  c = " and "
End If

If c_nivel3.ListIndex > 0 Then
  q = q & c & " [pos3] = " & c_nivel3.ItemData(c_nivel3.ListIndex)
  c = " and "
End If

If c_tipo.ListIndex > 0 Then
  q = q & c & " [tipo] = '" & Mid$(c_tipo, 2, 1) & "'"
End If

q = q & " order by [id_cuenta]"
rs.Open q, cn1
While Not rs.EOF
  c = ""
  If rs("pos4") > 0 Then 'cuenta
    c = Format$(rs("id_cuenta"), "000000")
    p4 = rs("descripcion")
    p1 = ""
    p2 = ""
    p3 = ""
    t = "C"
  Else
    If rs("pos3") > 0 Then 'pos3
        p1 = ""
        p2 = ""
        p3 = rs("descripcion")
        p4 = ""
         t = "T"
    Else
       If rs("pos2") > 0 Then 'pos2
          p1 = ""
          p2 = rs("descripcion")
          p3 = ""
          p4 = ""
           t = "T"
       Else 'pos1
          p1 = rs("descripcion")
          p2 = ""
          p3 = ""
          p4 = ""
           t = "T"
       End If
    End If
  End If
  msf1.AddItem c & Chr$(9) & p1 & Chr$(9) & p2 & Chr$(9) & p3 & Chr$(9) & p4 & Chr$(9) & c
  rs.MoveNext
Wend
Set rs = Nothing

End Sub
Private Sub btnsale_Click()
Unload Me
End Sub

'FIXIT: Declare 'n' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE
Sub carga(ByVal n)
'n es nivel
Set rs = New ADODB.Recordset
Select Case n
Case Is = 1 'nivel1
  q = "select * from c_01 where [pos2] = 0 and [pos3] = 0 and [pos4] = 0"
  rs.Open q, cn1
  c_nivel1.clear
  While Not rs.EOF
    c_nivel1.AddItem rs("Descripcion")
'FIXIT: c_nivel1.ItemData(c_nivel1.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    c_nivel1.ItemData(c_nivel1.NewIndex) = rs("pos1")
    rs.MoveNext
  Wend
  c_nivel1.AddItem "<Todas>", 0
  c_nivel1.ListIndex = 0
  
Case Is = 2 'nivel2
  If c_nivel1.ItemData(c_nivel1.ListIndex) > 0 Then
     q = "select * from c_01 where [pos1] = " & c_nivel1.ItemData(c_nivel1.ListIndex) & " and [pos2] > 0 and [pos3] = 0 and [pos4] = 0"
  Else
   q = "select * from c_01 where  [pos2] > 0 and [pos3] = 0 and [pos4] = 0"
  End If
  rs.Open q, cn1
  c_nivel2.clear
  While Not rs.EOF
    c_nivel2.AddItem rs("Descripcion")
'FIXIT: c_nivel2.ItemData(c_nivel2.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    c_nivel2.ItemData(c_nivel2.NewIndex) = rs("pos2")
    rs.MoveNext
  Wend
  c_nivel2.AddItem "<Todas>", 0
  c_nivel2.ListIndex = 0
Case Is = 3 'nivel2
  q = "select * from c_01 "
  c = " where "
  If c_nivel1.ListIndex > 0 Then
    q = q & c & " [pos1] = " & c_nivel1.ItemData(c_nivel1.ListIndex)
    c = " and "
  End If
  
  If c_nivel2.ListIndex > 0 Then
    q = q & c & " [pos2] = " & c_nivel2.ItemData(c_nivel2.ListIndex)
    c = " and "
  End If
  q = q & c & " [pos3] > 0 and [pos4] = 0"
  rs.Open q, cn1
  c_nivel3.clear
  While Not rs.EOF
    c_nivel3.AddItem rs("Descripcion")
'FIXIT: c_nivel3.ItemData(c_nivel3.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    c_nivel3.ItemData(c_nivel3.NewIndex) = rs("pos3")
    rs.MoveNext
  Wend
  c_nivel3.AddItem "<Todas>", 0
  c_nivel3.ListIndex = 0
  
End Select
Set rs = Nothing

End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 2000 'nivel1
msf1.ColWidth(2) = 2000 'nivel2
msf1.ColWidth(3) = 2000 'nivel3
msf1.ColWidth(4) = 3000 'cuenta
msf1.ColWidth(5) = 1000

msf1.TextMatrix(0, 0) = "Cuenta"
msf1.TextMatrix(0, 1) = ""
msf1.TextMatrix(0, 2) = ""
msf1.TextMatrix(0, 3) = ""
msf1.TextMatrix(0, 4) = ""
msf1.TextMatrix(0, 5) = "Cuenta"

End Sub









Private Sub c_nivel1_LostFocus()
If c_nivel1.ListIndex < 0 Then
  c_nivel1.ListIndex = 0
End If
Call carga(2)
End Sub

Private Sub c_nivel2_LostFocus()
If c_nivel2.ListIndex < 0 Then
  c_nivel2.ListIndex = 0
End If
Call carga(3)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_Load()


Call armagrid
Call barraesag(Me)
Call carga(1)
Call carga(2)
Call carga(3)
c_tipo.ListIndex = 0


End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime -   "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5
    
    For i = 6 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "PLAN DE CUENTAS", "", "", "", 85, 7, True, False)
  End If
    
    
End If
End Sub


Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    cc_detalle.T_IDPROV = msf1.TextMatrix(msf1.Row, 1)
    cc_detalle.t_prov = msf1.TextMatrix(msf1.Row, 2)
    cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 5), 3, 4)
    cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 5), 1, 1)
    cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 5), 8, 8)
    cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 3)
    cc_detalle.t_NUMINT = msf1.TextMatrix(msf1.Row, 7)
    cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub

