VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form log_verlogs 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ADMINISTRADOR DE LOGS"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   18165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   18165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambiar"
      Height          =   975
      Left            =   13200
      TabIndex        =   22
      Top             =   8040
      Width           =   1095
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   120
         Picture         =   "gen_043.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por:"
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   8160
      Width           =   2295
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   2880
      TabIndex        =   14
      Top             =   1800
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   180027393
      CurrentDate     =   38754
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   6015
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   10610
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   17655
      Begin VB.ComboBox c_ente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12000
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   5415
      End
      Begin VB.ComboBox c_op 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   16
         Text            =   "c"
         Top             =   1320
         Width           =   6255
      End
      Begin VB.ComboBox c_modulo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "gen_043.frx":030A
         Left            =   12000
         List            =   "gen_043.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox c_usuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   0
         Text            =   "c_prov"
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Ente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10440
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Operacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   9960
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Módulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   10440
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   14760
      TabIndex        =   6
      Top             =   8040
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen_043.frx":0335
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen_043.frx":0BB7
         Style           =   1  'Graphical
         TabIndex        =   7
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
      TabIndex        =   5
      Top             =   9180
      Width           =   18165
      _ExtentX        =   32041
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
            TextSave        =   "22/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:42 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "log_verlogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer


Sub carga()
  espere.Show
  espere.Label1 = "Cargando logs...."
  espere.Refresh
  Call armagrid
  
  
  q = "select * from g11,g15, g1 where g11.[id_usuario] = g1.[id_usuario] and [id_operacion] = [id_oplog] "
  c = " and "
  If c_usuario.ListIndex > 0 Then
     q = q & c & " g11.[id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
  End If
  
  If c_op.ListIndex > 0 Then
    q = q & c & " [id_operacion] = " & c_op.ItemData(c_op.ListIndex)
  End If
  
  If t_fecha <> "" And IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha_hora]) >= datevalue('" & t_fecha & "')"
  End If
  
  If t_fecha2 <> "" And IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha_hora]) <= datevalue('" & t_fecha2 & "')"
  End If
    
   If c_modulo.ListIndex > 0 Then
        q = q & c & " g11.[modulo] = '" & Mid$(c_modulo, 1, 1) & "'"
   End If
  
    
   If c_ente.ListIndex > 0 Then
    q = q & c & " [id_clipro] = " & c_ente.ItemData(c_ente.ListIndex)
   End If

   
   q = q & " order by [fecha_hora], g11.[id_usuario]"
   
   
 
  Set rs = New ADODB.Recordset
  
  rs.Open q, cn1
  t = 0
  reg = 0
  While Not rs.EOF
     F = rs("fecha_hora")
   
    
     msf1.AddItem F & Chr(9) & rs("detalle") & Chr(9) & rs("g11.modulo") & Chr(9) & rs("usuario") & Chr(9) & rs("num_int_comp") & Chr(9) & rs("obs") & Chr(9) & rs("descripcion")
     reg = reg + 1
     Label5 = reg
     Label5.Refresh
    rs.MoveNext
  Wend
  Unload espere
   
End Sub

Private Sub btnacepta_Click()
If t_fecha = "" Then
    MsgBox ("Este proceso genera mucha carga de trabajo, es necesario que se indique periodos cortos(fecha desde/fecha hasta) como corte para el procesamiento ")
Else
    Call carga
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 1700
msf1.ColWidth(1) = 4000 'cod prov
msf1.ColWidth(2) = 800
msf1.ColWidth(3) = 1700
msf1.ColWidth(4) = 1700
msf1.ColWidth(5) = 4000
msf1.ColWidth(6) = 3000
msf1.ColWidth(7) = 3000


msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Detalle"
msf1.TextMatrix(0, 2) = "Modulo"
msf1.TextMatrix(0, 3) = "Usuario"
msf1.TextMatrix(0, 4) = "NI Comprobante"
msf1.TextMatrix(0, 5) = "Observaciones"
msf1.TextMatrix(0, 6) = "Operacion"
msf1.TextMatrix(0, 7) = "Ente"


For i = 0 To 7
    msf1.ColAlignment(i) = 1 'izq
Next i

End Sub









Private Sub c_modulo_LostFocus()
Call carga_ente
End Sub

Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
   t_fecha = cal1
Else
   t_fecha2 = cal1
End If
cal1.Visible = False

End Sub

Private Sub cal1_LostFocus()
If cal1.Tag = "1" Then
   t_fecha = cal1
Else
   t_fecha2 = cal1
End If
cal1.Visible = False

End Sub


Private Sub Command1_Click()
gen_seleccionarimp.Show
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_Load()
cal1.Visible = False
Call carga_usuarios(c_usuario)
c_usuario.AddItem "<Todos>", 0
c_usuario.ListIndex = 0

c_modulo.ListIndex = 0
Call carga_ente

Call armagrid

End Sub

Sub carga_operacion(m)
If m = "T" Then
   w = " "
Else
   w = " where modulo = '" & m & "'"
End If

Set rs = New ADODB.Recordset
q = "select * from g15 " & w & " order by descripcion"
rs.Open q, cn1
c_op.clear
While Not rs.EOF
  c_op.AddItem rs("Descripcion") & "  {" & rs("modulo") & "}"
  c_op.ItemData(c_op.NewIndex) = rs("id_oplog")
  rs.MoveNext
Wend

c_op.AddItem "<Todas>", 0
c_op.ListIndex = 0

Set rs = Nothing
End Sub

Sub carga_ente()
   If c_modulo.ListIndex = 0 Then
       c_ente.clear
       c_ente.AddItem "<Todos>", 0
       c_ente.ListIndex = 0
       
       Call carga_operacion("T")
    Else
        Select Case Mid$(c_modulo, 1, 1)
          Case Is = "V"
            c_ente.clear
            Call carga_clientes(c_ente)
            c_ente.AddItem "<Todos>", 0
            c_ente.ListIndex = 0
            
            Call carga_operacion("V")
          
          Case Is = "C"
           c_ente.clear
           Call carga_proveedores(c_ente)
           c_ente.AddItem "<Todos>", 0
           c_ente.ListIndex = 0
           
           Call carga_operacion("C")
           
           Case Else
           
            c_ente.clear
            c_ente.AddItem "<Todos>", 0
            c_ente.ListIndex = 0
            Call carga_operacion("T")
         
         End Select
    
   End If
   c_ente.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F1] Cliente -  [F8] Borra - [F11] Excel "
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
    c(6) = 6
    For i = 7 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LOGS", "Usuario:" & c_usuario, "Fecha desde: " & t_fecha & "  Fecha hasta: " & t_fecha2, "", 72, 8, True, False)
  End If

End If



 


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub


Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load vta_cc_detalle
    vta_cc_detalle.t_idprov = msf1.TextMatrix(msf1.Row, 1)
    vta_cc_detalle.t_prov = msf1.TextMatrix(msf1.Row, 2)
    vta_cc_detalle.t_sucursal = Mid$(msf1.TextMatrix(msf1.Row, 5), 3, 4)
    vta_cc_detalle.t_letra = Mid$(msf1.TextMatrix(msf1.Row, 5), 1, 1)
    vta_cc_detalle.t_numcomp = Mid$(msf1.TextMatrix(msf1.Row, 5), 8, 8)
    vta_cc_detalle.t_tipocomp = msf1.TextMatrix(msf1.Row, 3)
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 8)
    vta_cc_detalle.Show
  End If
End If

End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight

End Sub


Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub


