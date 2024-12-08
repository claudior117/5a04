VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_exportaplu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXPORTA PLU PARA REGISTRADORAS O BALANZAS"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox T_etiqueta 
      Height          =   285
      Left            =   7200
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox c_pesable 
      Height          =   315
      ItemData        =   "vta075.frx":0000
      Left            =   1800
      List            =   "vta075.frx":000A
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Salida:"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   6360
      Width           =   9975
      Begin VB.CommandButton Command2 
         Caption         =   "Carpeta destino:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_carpeta 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   1
      Top             =   7440
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta075.frx":0016
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
         Picture         =   "vta075.frx":0898
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
      Top             =   8550
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   14994
            MinWidth        =   14994
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "08/12/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "08:31 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Los PLU deben estar creados en la lista de precio, y el codigo de departamento debe existir en la  balanza o registradora"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   7800
      Width           =   9855
   End
   Begin VB.Label Label3 
      Caption         =   "Cod. Etiqueta / Tique"
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Pesable"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   $"vta075.frx":111A
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Width           =   9855
   End
   Begin VB.Label Label4 
      Caption         =   "IMPORTANTE: Al importar los archivos en la página del AFIP seleccionar importes expresados en $ argentinos"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   8280
      Width           =   9255
   End
End
Attribute VB_Name = "vta_exportaplu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim c5 As Double

Sub carga()
 'Dim cr(2) As Long
 Dim nic As String
 espere.Show
 espere.Label1 = "Espere...... Buscando PLU"
 espere.Refresh
 Call armagrid
 
  pesable = Mid$(c_pesable, 1, 1)
  q = "select * from A2 where [plu] <> 0 "
  q = q & " order by [plu]"
  Set rs = New ADODB.Recordset
  er = ""
  rs.Open q, cn1
  While Not rs.EOF
    msf1.AddItem "" & Chr$(9) & rs("plu") & Chr(9) & rs("id_producto") & Chr(9) & rs("descripcion") & Chr$(9) & rs("id_departamento") & Chr$(9) & rs("precio_final") & Chr$(9) & pesable & Chr$(9) & T_etiqueta
    
    rs.MoveNext
  Wend
  Unload espere
     
End Sub
Private Sub btnacepta_Click()
  Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub










Private Sub Command2_Click()
Load gen_seleccionacarpeta
gen_seleccionacarpeta.t_llamada = "7"
gen_seleccionacarpeta.Show

End Sub





Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 1000
msf1.ColWidth(3) = 5000
msf1.ColWidth(4) = 2000
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 1000
msf1.TextMatrix(0, 1) = ""
msf1.TextMatrix(0, 1) = "PLU"
msf1.TextMatrix(0, 2) = "Basico"
msf1.TextMatrix(0, 3) = "Producto"
msf1.TextMatrix(0, 4) = "Departamento"
msf1.TextMatrix(0, 5) = "Precio"
msf1.TextMatrix(0, 6) = "Pesable"
msf1.TextMatrix(0, 7) = "Cod. Etiqueta/Tique"

End Sub

Private Sub Form_Load()
Call armagrid
t_carpeta = "c:\"
c_pesable.ListIndex = 0
T_etiqueta = 1
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ESPACIO] Selecciona - [F2] Todos - [F5] Exporta -  [F11] Excel"

End Sub




Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyF11 Then
 
  Call exportaexcel(msf1)
End If


If KeyCode = vbKeyF5 Then
  J = MsgBox("Confirma genera archivos para PLU. Carpeta destino " & t_carpeta, 4)
  If J = 6 Then
    If t_carpeta <> "" Then
      Call exporta
    Else
      MsgBox ("Debe seleccionar una carpeta destino")
   End If
  End If
  
End If

If KeyCode = vbKeyF2 Then
 k = 1
 If k <= msf1.Rows - 1 Then
  ee = msf1.TextMatrix(k, 0)
  If ee = "**" Then
    ee = ""
  Else
    ee = "**"
  End If
 End If
 
 
 While k <= msf1.Rows - 1
   msf1.TextMatrix(k, 0) = ee
   k = k + 1
 Wend
End If
End Sub
Sub exporta()
Dim Detalle As String
k = 1
a1 = t_carpeta & "plu.txt"
Open a1 For Output As #1
ni15 = "000000000000000"
Detalle = String(75, " ")
cont = 0
While k <= msf1.Rows - 1
  If msf1.TextMatrix(k, 0) = "**" Then
   l = msf1.TextMatrix(k, 1) & "," & msf1.TextMatrix(k, 3) & "," & msf1.TextMatrix(k, 4) & "," & msf1.TextMatrix(k, 5) & "," & msf1.TextMatrix(k, 6) & "," & msf1.TextMatrix(k, 7)
   Print #1, l
   cont = cont + 1
  End If
  k = k + 1

Wend
Close #1

MsgBox ("Operacion Terminada. Se exportaron " & cont & " PLUs")

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 14)
    vta_cc_detalle.Show
  End If
End If

If KeyAscii = vbKeySpace Then
  r = msf1.Row
  ee = msf1.TextMatrix(r, 0)
  If ee = "**" Then
    msf1.TextMatrix(r, 0) = ""
    
  Else
    msf1.TextMatrix(r, 0) = "**"
    
  End If
End If

End Sub

