VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_remitos_predef 
   BackColor       =   &H00E0E0E0&
   Caption         =   "REMITOS PREDEFINIDOS"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   17760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   17760
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   11245
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   12735
      Begin VB.ComboBox c_rempredef 
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
         Left            =   2280
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   240
         Width           =   8655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nuevo"
         Height          =   255
         Left            =   11280
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Remito"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   13440
      TabIndex        =   2
      Top             =   7920
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta073A.frx":0000
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
         Picture         =   "vta073A.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   3
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
      TabIndex        =   1
      Top             =   9180
      Width           =   17760
      _ExtentX        =   31327
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
            Object.Width           =   26458
            MinWidth        =   26458
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "26/07/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:01 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vta_remitos_predef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Dim calcula_perc_ib As String
Dim alicuota_perc_ib As Single
Dim minimo_perc_ib As Double

Sub renumera()
r = 1
For i = 1 To msf1.Rows - 1
 If Val(msf1.TextMatrix(i, 0)) <> 0 Then
    msf1.TextMatrix(i, 0) = r
    r = r + 1
 End If
Next i


End Sub



Sub limpia()
   Call armagrid
   
End Sub
Sub carga()
  Call armagrid
  Set rs1 = New ADODB.Recordset
  q = " select * from vta_018, a2, g5, g4 where [id_rempredef] = " & c_rempredef.ItemData(c_rempredef.ListIndex) & " and vta_018.id_producto = a2.id_producto and a2.id_unidad = g5.id_unidad and cod_tasaiva=id_tasaiva"
  rs1.Open q, cn1
  While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("vta_018.id_producto"), "00000") & Chr(9) & rs1("a2.descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("unidad") & Chr(9) & rs1("tasa")
        rs1.MoveNext
     Wend
     Set rs1 = Nothing
  
End Sub









Private Sub btnacepta_Click()
J = MsgBox("Graba Remito Predefinido", 4)
If J = 6 Then
 Call renumera
 
  
     Call graba
  

End If





End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 1800
msf1.ColWidth(2) = 6000
msf1.ColWidth(3) = 1300
msf1.ColWidth(4) = 1300
msf1.ColWidth(5) = 1300


msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "Unidad"
msf1.TextMatrix(0, 5) = "% Iva"



End Sub





Private Sub c_rempredef_GotFocus()
Call carga_rempredef(c_rempredef)

End Sub

Private Sub c_rempredef_LostFocus()
Call carga
End Sub

Private Sub Command1_Click()
vta_ABM_rempredef.Show
End Sub

Private Sub Form_Load()

Call INICIALIZA2(Me)
Call carga_rempredef(c_rempredef)

Call armagrid
Call barraesag(Me)

Load vta_remitos1_predef





End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_remitos1
Unload vta_transporte
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()

Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica  - [F5] Elimina - [F9] Graba"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If



If KeyCode = vbKeyF9 Then
  Call renumera
  btnacepta.SetFocus
 
End If

If KeyCode = vbKeyInsert Then
   vta_remitos1_predef.limpia
   vta_remitos1_predef.Show
End If


End Sub

Sub graba()
 
 Set rs1 = New ADODB.Recordset
 q = " select * from vta_017 where [id_rempredef] = " & c_rempredef.ItemData(c_rempredef.ListIndex)
 rs1.Open q, cn1
 If Not rs1.EOF And Not rs1.BOF Then
     numrem = c_rempredef.ItemData(c_rempredef.ListIndex)
     cn1.BeginTrans
     QUERY = "delete from vta_018 where id_rempredef = " & numrem
     cn1.Execute QUERY
   
     cn1.CommitTrans
     
     
     
      cn1.BeginTrans
      
      For i = 1 To msf1.Rows - 1
        
          
        QUERY = "INSERT INTO vta_018([id_rempredef], [id_producto], [cantidad])"
        QUERY = QUERY & " VALUES (" & numrem & ", " & Val(msf1.TextMatrix(i, 1)) & ", " & Val(msf1.TextMatrix(i, 3)) & ")"
        cn1.Execute QUERY
      
      Next i
        
      cn1.CommitTrans
     
 End If
  Call armagrid
  c_rempredef.ListIndex = 0
  Set rs1 = Nothing
 MsgBox ("Remito predefinido actualizado")
      

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    vta_remitos1_predef.t_renglon = msf1.Row
    vta_remitos1_predef.t_basico = msf1.TextMatrix(msf1.Row, 1)
    vta_remitos1_predef.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    vta_remitos1_predef.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
    vta_remitos1_predef.t_unidad = msf1.TextMatrix(msf1.Row, 4)
    vta_remitos1_predef.t_pu = 0
    vta_remitos1_predef.t_bultos = 0
    vta_remitos1_predef.t_importe = 0
    
    vta_remitos1_predef.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True

End Sub









