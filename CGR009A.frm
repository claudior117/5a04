VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form CGR_CUENTAS0 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CUENTAS CONTABLES"
   ClientHeight    =   9030
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo de relacion de la cuenta con Caja"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   8040
      Width           =   8535
      Begin VB.Label Label3 
         Caption         =   "I -> Ingreso   *   E -> Egreso   *   A -> Ingreso y Egreso (Ambas)   *   N -> Ninguna"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   8535
      Begin VB.ComboBox c_nivel1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   4095
      End
      Begin VB.ComboBox c_tipo 
         Height          =   315
         ItemData        =   "CGR009A.frx":0000
         Left            =   6240
         List            =   "CGR009A.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox c_nivel2 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Text            =   "c_nivel2"
         Top             =   600
         Width           =   4095
      End
      Begin VB.ComboBox c_nivel3 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Text            =   "c_nivel3"
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nivel 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Tipo :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nivel 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nivel 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5175
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command5 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   5400
         Picture         =   "CGR009A.frx":003C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "CGR009A.frx":0346
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "CGR009A.frx":0650
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "CGR009A.frx":095A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "CGR009A.frx":0C64
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
         Picture         =   "CGR009A.frx":0F6E
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
         Picture         =   "CGR009A.frx":17F0
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
      Top             =   8775
      Width           =   11955
      _ExtentX        =   21087
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "CGR_CUENTAS0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
Call limpia
End Sub

Private Sub Command4_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 3 Then
 Call imprimir
Else
 Call sinpermisos
End If

End Sub
Sub imprimir()
Dim c(15) As Double
  
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 1
    c(1) = 2
    c(2) = 3
    c(3) = 4
    c(4) = 5
    For i = 5 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), Space(75) & "PLAN DE CUENTAS", "", "", "", 71, 9, True, False)
  End If

End Sub
Private Sub btnsale_Click()
Unload Me
End Sub


Private Sub Command1_Click()
Call nivel_acceso(7)
If para.id_grupo_modulo_actual >= 5 Then
  cgr_cuentas2.limpia
  cgr_cuentas2.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
 Call nivel_acceso(7)
 If para.id_grupo_modulo_actual >= 5 Then
   cgr_cuentas!t_funcion = "M"
   Call LLENACAMPOS
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
q = "select * from c_01 where [id_cuenta] = " & Val(msf1.TextMatrix(msf1.Row, 0))
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  Select Case rs("tipo_cuentacaja")
  Case Is = "A"
    cgr_cuentas!c_tipocaja.ListIndex = 0
  Case Is = "I"
    cgr_cuentas!c_tipocaja.ListIndex = 1
  Case Is = "E"
    cgr_cuentas!c_tipocaja.ListIndex = 2
  Case Is = "N"
    cgr_cuentas!c_tipocaja.ListIndex = 3
  End Select
  
    
  
  cgr_cuentas!t_id = Val(msf1.TextMatrix(msf1.Row, 0))
  cgr_cuentas!t_descripcion = rs("descripcion")
  cgr_cuentas.Show
End If
Set rs = Nothing
Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Cuenta. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
 Call nivel_acceso(7)
 If para.id_grupo_modulo_actual >= 8 Then
   J = MsgBox("Borrar una Cuenta o Titulo implica borrar todos los movimientos cargados y las Subcuentas Asociadas. Confirma borrar Cuenta: " & msf1.TextMatrix(msf1.Row, 0), 4)
   If J = 6 Then
      Set rs = New ADODB.Recordset
      q = "select * from c_01 where [id_cuenta] = " & Val(msf1.TextMatrix(msf1.Row, 0))
      rs.Open q, cn1
      ci = 0
      cf = 0
      If Not rs.EOF And Not rs.BOF Then
        If rs("pos4") <> 0 Then
          ci = rs("id_cuenta")
          cf = rs("id_cuenta")
        Else
          If rs("pos3") <> 0 Then
             ci = Val(Mid$(Format$(rs("id_cuenta"), "000000"), 1, 4) & "00")
             cf = Val(Mid$(Format$(rs("id_cuenta"), "000000"), 1, 4) & "99")
          Else
            If rs("pos2") <> 0 Then
             ci = Val(Mid$(Format$(rs("id_cuenta"), "000000"), 1, 2) & "0000")
             cf = Val(Mid$(Format$(rs("id_cuenta"), "000000"), 1, 2) & "9999")
            Else
             ci = Val(Mid$(Format$(rs("id_cuenta"), "000000"), 1, 1) & "00000")
             cf = Val(Mid$(Format$(rs("id_cuenta"), "000000"), 1, 1) & "99999")
            End If
          End If
        End If
      End If
      Set rs = Nothing
      
      Set rs = New ADODB.Recordset
      q = "select * from c_12 where [id_cuenta] >= " & ci & " and [id_cuenta] <= " & cf
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
         MsgBox ("La cuenta o Titulo que desea Eliminar tiene asientos asignados. Es imposible Eliminarla")
      Else
         QUERY = "DELETE FROM C_01 WHERE [id_cuenta] >= " & ci & " and [id_cuenta] <= " & cf
         cn1.BeginTrans
         cn1.Execute QUERY
         cn1.CommitTrans
      End If
   End If
 Else
  Call sinpermisos
 End If
End If

Exit Sub
e1:
 MsgBox ("Error al Borrar Cuenta")
 cn1.RollbackTrans
 Exit Sub
End Sub

Private Sub Command5_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Form_Activate()
Call limpia
msf1.SetFocus
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
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
  c1 = Format$(rs("id_cuenta"), "000000")
  If rs("pos4") > 0 Then 'cuenta
    c = Format$(rs("id_cuenta"), "000000")
    p4 = rs("descripcion")
    t = "C"
    If c_tipo.ListIndex = 1 Then
       'PARA CADA CUENTA BUSCO P1, P2 P3
        'P1
         Set rs1 = New ADODB.Recordset
         c1 = Val(Mid$(c, 1, 1) & "00000")
         q = "select * from c_01 where [id_cuenta] = " & c1
         rs1.Open q, cn1
         If Not rs1.EOF And Not rs1.BOF Then
           p1 = rs1("descripcion")
         Else
           p1 = "Error"
         End If
         Set rs1 = Nothing
           
        'p2
         Set rs1 = New ADODB.Recordset
         c1 = Val(Mid$(c, 1, 2) & "0000")
         q = "select * from c_01 where [id_cuenta] = " & c1
         rs1.Open q, cn1
         If Not rs1.EOF And Not rs1.BOF Then
           p2 = rs1("descripcion")
         Else
           p2 = "Error"
         End If
         Set rs1 = Nothing
         
          'P3
         Set rs1 = New ADODB.Recordset
         c1 = Val(Mid$(c, 1, 4) & "00")
         q = "select * from c_01 where [id_cuenta] = " & c1
         rs1.Open q, cn1
         If Not rs1.EOF And Not rs1.BOF Then
           p3 = rs1("descripcion")
         Else
           p3 = "Error"
         End If
         Set rs1 = Nothing
       
       
           
    Else
        p1 = ""
        p2 = ""
        p3 = ""
    
    End If
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
  msf1.AddItem c1 & Chr$(9) & p1 & Chr$(9) & p2 & Chr$(9) & p3 & Chr$(9) & p4 & Chr$(9) & c & Chr$(9) & rs("tipo_cuentacaja")
  rs.MoveNext
Wend
Set rs = Nothing
Call INICIALIZA2(cgr_cuentas)
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone

msf1.Rows = 1
msf1.Cols = 7
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 1800 'nivel1
msf1.ColWidth(2) = 1800 'nivel2
msf1.ColWidth(3) = 2000 'nivel3
msf1.ColWidth(4) = 3000 'cuenta
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 700

msf1.TextMatrix(0, 0) = "Cuenta"
msf1.TextMatrix(0, 1) = "Titulo"
msf1.TextMatrix(0, 2) = "Subnivel 1"
msf1.TextMatrix(0, 3) = "Subnivel 2"
msf1.TextMatrix(0, 4) = "Descirpcion"
msf1.TextMatrix(0, 5) = "Cuenta"
msf1.TextMatrix(0, 6) = "Caja"

For i = 0 To 6
    msf1.ColAlignment(i) = 1 'izq
Next i
'msf1.ColAlignment(i) = 9 'der


End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load cgr_cuentas
Call carga(1)
Call carga(2)
Call carga(3)
c_tipo.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload cgr_cuentas
End Sub


Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Call imprimir
End If
End Sub
