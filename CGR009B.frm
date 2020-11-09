VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_cuentas2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUENTAS CONTABLES"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6165
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   8295
      Begin VB.ComboBox c_tipocaja 
         Height          =   315
         ItemData        =   "CGR009B.frx":0000
         Left            =   2040
         List            =   "CGR009B.frx":0010
         TabIndex        =   30
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2040
         MaxLength       =   49
         TabIndex        =   26
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Tipo cuenta caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Nombre Cuenta o Rubro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ingreso Manual"
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   5535
      Begin VB.CommandButton Command4 
         Caption         =   "Cargar"
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox t_p4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox t_p3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox t_p2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox t_p1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Id. Cuenta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ingreso Asistido"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   8295
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   7320
         Picture         =   "CGR009B.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   7320
         Picture         =   "CGR009B.frx":0159
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   7320
         Picture         =   "CGR009B.frx":025E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox c_n3 
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   1440
         Width           =   5055
      End
      Begin VB.ComboBox c_n2 
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   1080
         Width           =   5055
      End
      Begin VB.ComboBox c_n1 
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Nivel 3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Nivel 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Nivel 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Id. Cuenta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6840
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "CGR009B.frx":0363
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
         Picture         =   "CGR009B.frx":0BE5
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
      Top             =   5910
      Width           =   8745
      _ExtentX        =   15425
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
Attribute VB_Name = "cgr_cuentas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String

Sub limpia()
t_id = ""
c_n1.ListIndex = 0
c_n2.ListIndex = 0
c_n3.ListIndex = 0
t_descripcion = ""
End Sub

Private Sub btnacepta_Click()
If t_p1 <> "0" And t_p2 <> "0" And t_p3 <> "00" And t_p4 <> "00" Then
  Call graba
End If
End Sub

Sub graba()
If Val(t_id) > 0 Then
 J = MsgBox("Confirma Valores para Grabar", 4)
 If J = 6 Then
   On Error GoTo ERRORGRABA
   Call NULOS(t_descripcion)
    Select Case c_tipocaja.ListIndex
     Case Is = 0
       tc = "A"
     Case Is = 1
       tc = "I"
     Case Is = 2
       tc = "E"
     Case Is = 3
       tc = "N"
    End Select
       
      
      
      QUERY = "INSERT INTO c_01([id_cuenta], [DEscripcion], [pos1], [pos2], [pos3], [pos4], [tipo], [tipo_cuentacaja])"
      QUERY = QUERY & " VALUES (" & Val(t_id) & ", '" & t_descripcion & "', " & Val(Mid$(t_id, 1, 1)) & ", " & Val(Mid$(t_id, 2, 1)) & ", " & Val(Mid$(t_id, 3, 2)) & ", " & Val(Mid$(t_id, 5, 2)) & ", 'C', '" & tc & "')"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      Call CGR_CUENTAS0.limpia
      CGR_CUENTAS0.Show
      Me.Hide
 End If
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub







Private Sub c_n1_LostFocus()
If c_n1.ListIndex < 0 Then
  c_n1.ListIndex = 0
  t_p1 = 0
Else
  t_p1 = c_n1.ItemData(c_n1.ListIndex)
End If
Call cargan2
Call cargan3
End Sub

Private Sub c_n2_LostFocus()
If c_n2.ListIndex < 0 Then
  c_n2.ListIndex = 0
  t_p2 = "0"
Else
  t_p2 = c_n2.ItemData(c_n2.ListIndex)
End If
Call cargan3
End Sub

Private Sub c_n3_LostFocus()
If c_n3.ListIndex < 0 Then
  c_n3.ListIndex = 0
  t_p3 = "00"
Else
  t_p3 = Format$(c_n3.ItemData(c_n3.ListIndex), "00")
  Call busca
End If
End Sub

Private Sub c_tipocaja_LostFocus()
If c_tipocaja.ListIndex < 0 Then
  c_tipocaja.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
Set rs = New ADODB.Recordset
q = "select * from c_01 where [pos1] <> 0 and [pos2] = 0 and [pos3] = 0 and [pos4] = 0 order by [pos1]"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
   rs.MoveLast
   c = rs("pos1") + 1
Else
   c = 1
End If
Set rs = Nothing
Load cgr_cuentas
cgr_cuentas.t_funcion = "A"
cgr_cuentas.t_id = Format$(c, "0") & "00000"
cgr_cuentas.Show

End Sub

Private Sub Command1_LostFocus()
Call cargan1
Call cargan2
Call cargan3
End Sub

Private Sub Command2_Click()
If c_n1.ListIndex > 0 Then
 Set rs = New ADODB.Recordset
 q = "select * from c_01 where [pos1] = " & c_n1.ItemData(c_n1.ListIndex) & " and [pos2] <> 0 and [pos3] = 0 and [pos4] = 0 order by [pos2]"
 rs.Open q, cn1, adOpenDynamic, adLockOptimistic
 If Not rs.EOF And Not rs.BOF Then
   rs.MoveLast
   c = rs("pos2") + 1
 Else
   c = 1
 End If
 Set rs = Nothing
 Load cgr_cuentas
 cgr_cuentas.t_funcion = "A"
 cgr_cuentas.t_id = Format$(c_n1.ItemData(c_n1.ListIndex), "0") & Format$(c, "0") & "0000"
 cgr_cuentas.Show
Else
 MsgBox ("Antes de dar de Alta un Titulo de Nivel 2 debe seleccionar el Nivel1")
End If
End Sub

Private Sub Command2_LostFocus()
Call cargan2
Call cargan3
End Sub

Private Sub Command3_Click()
If c_n1.ListIndex > 0 And c_n2.ListIndex > 0 Then
 Set rs = New ADODB.Recordset
 q = "select * from c_01 where [pos1] = " & c_n1.ItemData(c_n1.ListIndex) & " and [pos2] = " & c_n2.ItemData(c_n2.ListIndex) & " and [pos3] <> 0 and [pos4] = 0 order by [pos3]"
 rs.Open q, cn1, adOpenDynamic, adLockOptimistic
 If Not rs.EOF And Not rs.BOF Then
   rs.MoveLast
   c = rs("pos3") + 1
 Else
   c = 1
 End If
 Set rs = Nothing
 Load cgr_cuentas
 cgr_cuentas.t_funcion = "A"
 cgr_cuentas.t_id = Format$(c_n1.ItemData(c_n1.ListIndex), "0") & Format$(c_n2.ItemData(c_n2.ListIndex), "0") & Format$(c, "00") & "00"
 cgr_cuentas.Show
Else
 MsgBox ("Antes de dar de Alta un Titulo de Nivel 3 debe seleccionar el Nivel 1 y el Nivel 2")
End If

End Sub

Private Sub Command4_Click()
t_p1 = Format$(Val(t_p1), "0")
t_p2 = Format$(Val(t_p2), "0")
t_p3 = Format$(Val(t_p3), "00")
t_p4 = Format$(Val(t_p4), "00")
t_id = Format$(Val(t_p1 & t_p2 & t_p3 & t_p4), "000000")
Set rs = New ADODB.Recordset
q = "select * from c_01 where [id_cuenta] = " & Val(t_id)
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  MsgBox ("La cuenta existe No se puede agregar")
Else
  c_n1.ListIndex = buscaindice(c_n1, Val(t_p1))
  c_n2.ListIndex = buscaindice(c_n2, Val(t_p2))
  c_n3.ListIndex = buscaindice(c_n3, Val(t_p3))
End If




End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  t_descripcion.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   Case Is = vbKeyF9
     Call graba
         
End Select

End Sub
Sub cargan1()
'n es nivel
  Set rs = New ADODB.Recordset
  q = "select * from c_01 where [pos2] = 0 and [pos3] = 0 and [pos4] = 0"
  rs.Open q, cn1
  c_n1.clear
  While Not rs.EOF
    c_n1.AddItem rs("Descripcion")
'FIXIT: c_n1.ItemData(c_n1.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    c_n1.ItemData(c_n1.NewIndex) = rs("pos1")
    rs.MoveNext
  Wend
  c_n1.AddItem "<Sin Seleccionar>", 0
  c_n1.ListIndex = 0
  Set rs = Nothing

  t_p1 = "0"
  t_p2 = "0"
  t_p3 = "00"
  t_p4 = "00"
End Sub
Sub cargan2()
'n es nivel
 c_n2.clear
 n = c_n1.ItemData(c_n1.ListIndex)
 If n <> 0 Then
  Set rs = New ADODB.Recordset
  q = "select * from c_01 where [pos1] = " & n & " and [pos2] <> 0 and  [pos3] = 0 and [pos4] = 0"
  rs.Open q, cn1
  While Not rs.EOF
    c_n2.AddItem rs("Descripcion")
'FIXIT: c_n2.ItemData(c_n2.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    c_n2.ItemData(c_n2.NewIndex) = rs("pos2")
    rs.MoveNext
  Wend
  c_n2.AddItem "<Sin Seleccionar>", 0
  c_n2.ListIndex = 0
Else
  c_n2.AddItem "<Sin Seleccionar>", 0
  c_n2.ListIndex = 0
End If
Set rs = Nothing
c_n2.ListIndex = 0
End Sub


Sub cargan3()
'n es nivel
 c_n3.clear
 n1 = c_n1.ItemData(c_n1.ListIndex)
 n2 = c_n2.ItemData(c_n2.ListIndex)
 If n1 <> 0 And n2 <> 0 Then
  Set rs = New ADODB.Recordset
  q = "select * from c_01 where [pos1] = " & n1 & " and [pos2] = " & n2 & "  and  [pos3] <> 0 and [pos4] = 0"
  rs.Open q, cn1
  While Not rs.EOF
    c_n3.AddItem rs("Descripcion")
'FIXIT: c_n3.ItemData(c_n3.NewIndex property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    c_n3.ItemData(c_n3.NewIndex) = rs("pos3")
    rs.MoveNext
  Wend
  c_n3.AddItem "<Sin Seleccionar>", 0
  c_n3.ListIndex = 0
Else
  c_n3.AddItem "<Sin Seleccionar>", 0
  c_n3.ListIndex = 0
End If
Set rs = Nothing
c_n3.ListIndex = 0
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call cargan1
c_n2.clear
c_n2.AddItem "<Sin Seleccionar>", 0
c_n2.ListIndex = 0
c_n3.AddItem "<Sin Seleccionar>", 0
c_n3.ListIndex = 0
t_id = "000000"
c_tipocaja.ListIndex = 0
End Sub



Private Sub t_descripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call busca
   btnacepta.SetFocus
End If
End Sub

Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub
Sub busca()
b = 1
If c_n1.ListIndex = 0 Then
 b = 0
End If
If c_n2.ListIndex = 0 Then
 b = 0
End If
If c_n3.ListIndex = 0 Then
 b = 0
End If

If b = 1 Then
   c = Format$(c_n1.ItemData(c_n1.ListIndex), "0") & Format$(c_n2.ItemData(c_n2.ListIndex), "0") & Format$(c_n3.ItemData(c_n3.ListIndex), "00")
   ci = Val(c & "01")
   cf = Val(c & "99")
   Set rs = New ADODB.Recordset
   q = "select * from c_01 where [id_cuenta] >= " & ci & " and [id_cuenta] <= " & cf & " order by [id_cuenta]"
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   If Not rs.EOF And Not rs.BOF Then
     rs.MoveLast
     c = Format$(rs("id_cuenta") + 1, "000000")
     u = Val(Mid$(c, 5, 2))
   Else
     u = 1
     c = c & Format$(u, "00")
   End If
   t_id = c
   t_p4 = u
   Set rs = Nothing
Else
  t_id = "000000"
  t_p1 = "0"
  t_p2 = "0"
  t_p3 = "00"
  t_p4 = "00"
End If

End Sub





Private Sub t_p4_LostFocus()
If t_p1 <> "0" And t_p2 <> "0" And t_p3 <> "00" And Val(t_p4) > 0 Then
  t_id = t_p1 & t_p1 & Format$(Val(t_p3), "00") & Format$(Val(t_p4), "00")
  Call busca2
End If
End Sub
Sub busca2()
Set rs = New ADODB.Recordset
q = "select * from c_01 where [id_cuenta] = " & Val(t_id)
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  MsgBox ("La Cuenta Existe y no se puede crear")
  t_p1 = "0"
  t_p2 = "0"
  t_p3 = "00"
  t_p4 = "00"
  t_id = "000000"
End If
Set rs = Nothing

End Sub
