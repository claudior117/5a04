VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form com_seloc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SELECCIONA O.C. A FACTURAR"
   ClientHeight    =   6480
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   4095
      Begin VB.TextBox t_seleccionados 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "O.C. Seleciionadas:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "O.C. PENDIENTES"
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   4455
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   $"com008.frx":0000
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   7858
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FillStyle       =   1
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "com008.frx":00F6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "com008.frx":0978
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6225
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "com_seloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean
Dim grecargocc As Single

Function habilita() As Boolean
 If msf1.Rows > 1 Then
   h = 0
   For i = 1 To msf1.Rows - 1
      If msf1.TextMatrix(i, 0) = "**" Then
         h = 1
         i = msf1.Rows
      End If
   Next i
   If h = 0 Then
     habilita = False
     btnacepta.Enabled = False
   Else
     habilita = True
     btnacepta.Enabled = True
   End If
 Else
   habilita = False
   btnacepta.Enabled = False
 End If

End Function
Sub limpia()
Call armagrid
t_r1 = 0

btnacepta.Enabled = False
habilitafacturaremitos = habilita
End Sub



Private Sub btnsale_Click()
Me.Hide
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 5
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 300
msf1.ColWidth(1) = 1400
msf1.ColWidth(2) = 1400
msf1.ColWidth(3) = 800
msf1.ColWidth(4) = 1200
msf1.TextMatrix(0, 1) = "Nro. Comprobante"
msf1.TextMatrix(0, 2) = "Fecha"
msf1.TextMatrix(0, 3) = "Tipo"
msf1.TextMatrix(0, 4) = "Nro. Interno"
For i = 0 To 3
 msf1.ColAlignment(i) = 1 'izq
Next i

msf1.FocusRect = flexFocusNone

End Sub

Private Sub Form_Activate()
Call cuenta
End Sub

Private Sub Form_Load()

Call limpia
Load cc_detalle

End Sub

Sub carga()
   Call limpia
   q = "select * from a5 where [id_tipocomp] = 65 and estado = 'P' and [id_proveedor] = " & ABM_COMP_COMPRA.c_prov.ItemData(ABM_COMP_COMPRA.c_prov.ListIndex)
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   While Not rs.EOF
     nc = Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     F = Format$(rs("fecha"), "dd/mm/yyyy")
     t = rs("total")
     msf1.AddItem "" & Chr$(9) & nc & Chr$(9) & F & Chr$(9) & t & Chr$(9) & rs("num_int")
     rs.MoveNext
   Wend
   
End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload cc_detalle
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(1) = "[Barra] Marca - [F5] Todos - [F9] Arma Factura - "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF9 Then
  
  Me.Hide
  ABM_COMP_COMPRA.msf1.SetFocus
End If

If KeyCode = vbKeyF5 Then
  If msf1.Rows > 1 Then
    For i = 1 To msf1.Rows - 1
      If msf1.TextMatrix(i, 0) = "**" Then
          msf1.TextMatrix(i, 0) = ""
      Else
         msf1.TextMatrix(i, 0) = "**"
      End If
    Next i
  End If
  habilitafacturaremito = habilita
  Call cuenta
  
  
End If

End Sub
Sub cuenta()
 If msf1.Rows > 1 Then
   h = 0
   For i = 1 To msf1.Rows - 1
      If msf1.TextMatrix(i, 0) = "**" Then
         h = h + 1
      End If
   Next i
    
 Else
  h = 0
 End If
 t_seleccionados = h
End Sub


'FIXIT: Declare 'k' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE



Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Rows > 1 Then
    If Val(msf1.TextMatrix(msf1.Row, 4)) > 0 Then
       cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 4)
       cc_detalle.Show
    End If
  End If
  habilitafacturaremito = habilita
End If

If KeyAscii = vbKeySpace Then
  If Val(msf1.TextMatrix(msf1.Row, 4)) > 0 Then
      If msf1.TextMatrix(msf1.Row, 0) = "**" Then
          msf1.TextMatrix(msf1.Row, 0) = ""
      Else
         msf1.TextMatrix(msf1.Row, 0) = "**"
      End If
      
  End If
  habilitafacturaremitos = habilita
End If


End Sub

Private Sub msf1_LostFocus()
msf1.FocusRect = flexFocusNone

End Sub




