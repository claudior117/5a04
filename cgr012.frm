VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_buscacuenta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUSCADOR DE CUENTAS "
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9060
   ScaleWidth      =   12360
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar"
      Height          =   975
      Left            =   10080
      TabIndex        =   17
      Top             =   7440
      Width           =   1695
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   960
         Picture         =   "cgr012.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cgr012.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   9600
      TabIndex        =   15
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   120
         Picture         =   "cgr012.frx":1104
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   5535
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   5055
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8916
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenado por"
      Height          =   855
      Left            =   7680
      TabIndex        =   10
      Top             =   840
      Width           =   1815
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Denominacion"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Registros"
      Height          =   735
      Left            =   7680
      TabIndex        =   8
      Top             =   0
      Width           =   1815
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         MaxLength       =   13
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox c_tipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "cgr012.frx":1209
         Left            =   1680
         List            =   "cgr012.frx":1216
         TabIndex        =   20
         Top             =   1320
         Width           =   3495
      End
      Begin VB.ComboBox c_grupo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox t_detalle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox t_basico 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "Tipo Caja"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00008000&
         Caption         =   "Rubro"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
         Caption         =   "Denominacion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Codigo"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8700
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   19403
            MinWidth        =   19403
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_buscacuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub carga()
espere.Show
espere.Label1 = "Cargando Cuentas...."
espere.Refresh
 Call armagrid
 
 
 ct = Space$(10)
 Set rs = New ADODB.Recordset
 q = "select * from c_01 where [tipo] = 'C'"
 c = " and "
 
 If t_basico <> "" Then
   If Len(t_basico) >= 6 Then
     q = q & c & "[id_cuenta] = " & Val(t_basico)
   Else
     m = 6 - Len(t_basico)
     ci = t_basico
     cf = t_basico
     For X = 1 To m
       ci = ci & "0"
       cf = cf & "9"
     Next X
     q = q & c & "[id_cuenta] >= " & ci & " and [id_cuenta] <= " & cf
   End If
   c = " and "
 End If
 
 If t_detalle <> "" Then
   q = q & c & "[descripcion] like  '%" & t_detalle & "%'"
   c = " and "
 End If
 
 If c_grupo.ListIndex > 0 Then
      Set rs2 = New ADODB.Recordset
       k = "select * from c_01 where [id_cuenta] = " & c_grupo.ItemData(c_grupo.ListIndex)
       rs2.Open k, cn1
       If Not rs2.EOF And Not rs2.BOF Then
        ci = Format$(rs2("pos1"), "0")
        cf = Format$(rs2("pos1"), "0")
        If rs2("pos2") > 0 Then
         ci = ci & Format$(rs2("pos2"), "0")
         cf = cf & Format$(rs2("pos2"), "0")
         If rs2("pos3") > 0 Then
           ci = ci & Format$(rs2("pos3"), "00")
           cf = cf & Format$(rs2("pos3"), "00")
         Else
           ci = ci & "00"
           cf = cf & "99"
         End If
        Else
         ci = ci & "000"
         cf = cf & "999"
        End If
      End If
      ci = ci & "00"
      cf = cf & "99"
      Set rs2 = Nothing
  
   
   
   
   q = q & c & "[id_cuenta] >= " & ci & " and [id_cuenta] <= " & cf
   c = " and "
 End If
 
If c_tipo.ListIndex > 0 Then
 
 Select Case c_tipo.ListIndex
  Case Is = 1 'ingresos
    q = q & c & " ([tipo_cuentacaja] = 'I' or [tipo_cuentacaja] = 'A') "
  Case Is = 2
    q = q & c & " ([tipo_cuentacaja] = 'E' or [tipo_cuentacaja] = 'A')"
 End Select
 c = " and "
End If
 
 If Option1 = True Then
   q = q & " order by [id_cuenta]"
 Else
   q = q & " order by [descripcion]"
 End If

rs.Open q, cn1
 t_encontrados = 0
 While Not rs.EOF
    b = Format$(rs("id_cuenta"), "000000")
    d = rs("descripcion")
    
    msf1.AddItem b & Chr$(9) & d
    t_encontrados = Val(t_encontrados) + 1
    rs.MoveNext
 Wend
msf1.SetFocus
Set rs = Nothing
Unload espere
End Sub


Private Sub btnacepta_Click()

Call carga

End Sub


Private Sub btnsale_Click()
Unload Me
End Sub

Private Sub c_grupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub




Private Sub c_grupo_LostFocus()
If c_grupo.ListIndex < 0 Then
  c_grupo.ListIndex = 0
End If
End Sub

Private Sub c_tipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If
End Sub

Private Sub Command2_Click()
CGR_CUENTAS0.Show
End Sub

Private Sub Form_Activate()
para.cuenta_sel = 0
t_detalle.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Me.Hide
End If

End Sub
Sub cargan1()
'n es nivel
  Set rs = New ADODB.Recordset
  q = "select * from c_01 where [tipo] = 'T'"
  rs.Open q, cn1
  c_grupo.clear
  While Not rs.EOF
    c_grupo.AddItem rs("Descripcion")
    c_grupo.ItemData(c_grupo.NewIndex) = rs("id_cuenta")
    rs.MoveNext
  Wend
  c_grupo.AddItem "<Todas>", 0
  c_grupo.ListIndex = 0
  Set rs = Nothing

  End Sub
Private Sub Form_Load()
  Call cargan1
  Call armagrid
  c_tipo.ListIndex = 0
  End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 4
msf1.ColWidth(0) = 1200
msf1.ColWidth(1) = 8000
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 1000



msf1.TextMatrix(0, 0) = "Codigo"
msf1.TextMatrix(0, 1) = "Denominacion"
msf1.TextMatrix(0, 2) = "Tipo"
msf1.TextMatrix(0, 3) = "Caja"

For i = 0 To 3
  msf1.ColAlignment(i) = 1 'izq
Next i
'For i = 2 To 6
'  msf1.ColAlignment(i) = 9 'der
'Next i

End Sub

  




Private Sub Form_Unload(Cancel As Integer)
Unload vta_listaprecios2
Unload vta_listaprecios3
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F2] Selecciona -  [F4] Saca - [F7] Imprime - [Esc] Cancela"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    para.cuenta_sel = p
    Me.Hide
  Else
    para.cuenta_sel = 0
  End If
End If

If KeyCode = vbKeyF4 Then
  r = msf1.Row
  p = Val(msf1.TextMatrix(r, 0))
  If p > 1 Then
    msf1.RemoveItem r
    t_encontrados = Val(t_encontrados) - 1
  End If
End If

If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
      
    For i = 3 To 14
      c(i) = -1
    Next i
    
    If t_detalle <> "" Then
      t = "Detalle: " & t_detalle
    Else
      t = ""
    End If
    
    If c_grupo.ListIndex > 0 Then
       t1 = "Grupo: " & c_grupo
    End If
    Call imprimegrid(msf1, c(), "Cuentas Contables", "", t, t1, 80, 8, True, False, "V")
  End If
End If


End Sub
Private Sub msf1_LostFocus()
Call barra(Me)
End Sub

Private Sub t_basico_GotFocus()
t_basico = ""
End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub


Private Sub T_detalle_GotFocus()
t_detalle = ""
End Sub

Private Sub t_detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call carga
End If
End Sub

