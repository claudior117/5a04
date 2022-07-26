VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_selcomp 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCION COMPROBANTES ASOCIADOS A NC"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   6210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   4095
      Begin VB.TextBox t_seleccionados 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Comprobantes Selecionados:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "REMITOS PENDIENTES"
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   3855
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Seleccione llas facturas sobre la que va a aplicar la NC "
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6800
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
         Picture         =   "vta061.frx":0000
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
         Picture         =   "vta061.frx":0882
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
Attribute VB_Name = "vta_selcomp"
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

Private Sub btnacepta_Click()
Call cuenta
Me.Hide
End Sub

Private Sub btnsale_Click()
Call cuenta
Me.Hide
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 300
msf1.ColWidth(1) = 1100
msf1.ColWidth(2) = 1100
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 500
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1200
msf1.TextMatrix(0, 1) = "Punto venta"
msf1.TextMatrix(0, 2) = "Num. Comp"
msf1.TextMatrix(0, 3) = "Fecha"
msf1.TextMatrix(0, 4) = "Tipo"
msf1.TextMatrix(0, 5) = "Total"
msf1.TextMatrix(0, 6) = "Nro. Interno"
For i = 0 To 3
 msf1.ColAlignment(i) = 1 'izq
Next i

msf1.FocusRect = flexFocusNone

End Sub

Private Sub Form_Activate()
Call cuenta
End Sub

Private Sub Form_Load()
Load vta_cc_detalle
Call limpia


End Sub

Sub carga()
   Call limpia
   q = "select * from vta_02 where ([id_tipocomp] = 1 or [id_tipocomp] = 30)  and [id_cliente] = " & vta_facturacion.c_prov.ItemData(vta_facturacion.c_prov.ListIndex)
   Set rs = New ADODB.Recordset
   rs.Open q, cn1
   While Not rs.EOF
   
     If rs("letra") = "A" Then
       If rs("id_tipocomp") = 1 Then
            t = 1
       Else
            t = 201
       End If
     Else
      If rs("id_tipocomp") = 1 Then
                t = 6
      Else
                t = 206
      End If
     End If
     
     nc = Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comp"), "00000000")
     F = Format$(rs("fecha"), "yyyymmdd")
     msf1.AddItem "" & Chr$(9) & Format$(rs("sucursal"), "0000") & Chr$(9) & Format$(rs("num_comp"), "00000000") & Chr$(9) & F & Chr$(9) & t & Chr$(9) & rs("total") & Chr$(9) & rs("num_int")
     rs.MoveNext
   Wend
   
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_cc_detalle
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(1) = "[Barra] Marca - [F5] Todos -  "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)


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


If KeyCode = vbKeyF9 Then
    Call cuenta
    Me.Hide
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








Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Rows > 1 Then
    If Val(msf1.TextMatrix(msf1.Row, 6)) > 0 Then
       vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 6)
       vta_cc_detalle.Show
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
  Call cuenta
End If


End Sub

Private Sub msf1_LostFocus()
msf1.FocusRect = flexFocusNone

End Sub




