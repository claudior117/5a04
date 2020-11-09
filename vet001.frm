VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vet_abm_mas1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATOS DE LA MASCOTA"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   11790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9240
      TabIndex        =   17
      Top             =   120
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   18
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
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   8775
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   6000
         Picture         =   "vet001.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   6000
         Picture         =   "vet001.frx":0105
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   7440
         Picture         =   "vet001.frx":020A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox c_sexo 
         Height          =   315
         ItemData        =   "vet001.frx":030F
         Left            =   2160
         List            =   "vet001.frx":0319
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2880
         Width           =   2175
      End
      Begin VB.ComboBox c_cliente 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Text            =   "c_cliente"
         Top             =   1440
         Width           =   5175
      End
      Begin VB.ComboBox c_raza 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   3615
      End
      Begin VB.ComboBox c_tipo 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox t_fechanac 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_nombre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   0
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Sexo:"
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
         Height          =   495
         Index           =   6
         Left            =   480
         TabIndex        =   23
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Raza:"
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
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Tipo:"
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
         Index           =   1
         Left            =   480
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Dueño:"
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
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Nac:"
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
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "id. Mascota"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Nombre:"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Descripcion:"
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
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9360
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "vet001.frx":032C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vet001.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   9
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
      TabIndex        =   7
      Top             =   5160
      Width           =   11790
      _ExtentX        =   20796
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
            TextSave        =   "21/12/2011"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:46 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "vet_abm_mas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String



Private Sub btnacepta_Click()
Call graba
End Sub

Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   If t_fechanac = "" Then
     t_fechanac = Format$(Now, "dd/mm/yyyy")
   End If
      
   On Error GoTo ERRORGRABA
    
   Select Case t_funcion
      
   Case "A"
      QUERY = "INSERT INTO vet_02([nombre], [descripcion], [fecha_nac], [sexo], [id_cliente], [id_raza], [id_tipo] )"
      QUERY = QUERY & " VALUES ('" & t_nombre & "', '" & t_descripcion & "', '" & t_fechanac & "', '" & Mid$(c_sexo, 1, 1) & "', " & C_CLIENTE.ItemData(C_CLIENTE.ListIndex) & ", " & c_raza.ItemData(c_raza.ListIndex) & ", " & c_tipo.ItemData(c_tipo.ListIndex) & ")"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   
   
   Case "M"
   
      QUERY = "update vet_02 set  [nombre]='" & t_nombre & "' , [descripcion]='" & t_descripcion & "' , [fecha_nac]='" & t_fechanac & "' , [sexo]='" & Mid$(c_sexo, 1, 1) & _
      "' , [id_cliente]=" & C_CLIENTE.ItemData(C_CLIENTE.ListIndex) & " , [id_raza]=" & c_raza.ItemData(c_raza.ListIndex) & " , [id_tipo]=" & c_tipo.ItemData(c_tipo.ListIndex)
      
      QUERY = QUERY & " where [id_animal]= " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
      
   Case "B"
      q = "select * from vet_01 where [id_animal] = " & Val(t_id)
      Set rs = New ADODB.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
          MsgBox ("La mascota tiene Historia Clinica. No se puede Eliminar")
      Else
          QUERY = "DELETE FROM vet_02 WHERE [id_animal] = " & Val(t_id)
          cn1.BeginTrans
          cn1.Execute QUERY
          cn1.CommitTrans
      End If
      Set rs = Nothing
   
   End Select
   
   
   vet_ABM_mas.Show
   Me.Hide
    
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub

Private Sub btnsale_Click()
Me.Hide
End Sub






Private Sub C_cliente_LostFocus()
If C_CLIENTE.ListIndex < 0 Then
  C_CLIENTE.ListIndex = 0
End If

End Sub

Private Sub c_raza_LostFocus()
If c_raza.ListIndex < 0 Then
  c_raza.ListIndex = 0
End If

End Sub

Private Sub c_sexo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub c_sexo_LostFocus()
If c_sexo.ListIndex < 0 Then
  c_sexo.ListIndex = 0
End If
End Sub

Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If
Call carga_raza(c_raza, c_tipo.ItemData(c_tipo.ListIndex))
c_raza.ListIndex = 0
End Sub

Private Sub Command1_Click()
vet_ABM_tipo.Show

End Sub

Private Sub Command1_LostFocus()
Call carga_tipo_mascota(c_tipo)
c_tipo.ListIndex = 0
Call carga_raza(c_raza, c_tipo.ItemData(c_tipo.ListIndex))
c_raza.ListIndex = 0

End Sub

Private Sub Command2_Click()
vta_ABM_cli.Show
End Sub

Private Sub Command2_LostFocus()
Call carga_clientes(C_CLIENTE)
C_CLIENTE.ListIndex = 0

End Sub

Private Sub Command3_Click()
vet_abm_razas.Show
End Sub

Private Sub Command3_LostFocus()
Call carga_raza(c_raza, c_tipo.ItemData(c_tipo.ListIndex))
c_raza.ListIndex = 0
End Sub

Private Sub Form_Activate()
If t_funcion = "B" Then
  btnacepta.Enabled = True
  btnacepta.SetFocus
Else
  t_nombre.SetFocus
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

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 6)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_clientes(C_CLIENTE)
C_CLIENTE.ListIndex = 0
Call carga_tipo_mascota(c_tipo)
c_tipo.ListIndex = 0
Call carga_raza(c_raza, c_tipo.ItemData(c_tipo.ListIndex))
c_raza.ListIndex = 0
c_sexo.ListIndex = 0
End Sub



Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub



Private Sub t_fechanac_LostFocus()
If t_fechanac <> "" Then
  If Not IsDate(t_fechanac) Then
    t_fechanac = Format$(Now, "dd/mm/yyyy")
  Else
    t_fechanac = Format$(t_fechanac, "dd/mm/yyyy")
  End If
Else
  t_fechanac = Format$(Now, "dd/mm/yyyy")
End If
End Sub

Private Sub t_nombre_LostFocus()
Call NULOS(t_nombre)
End Sub




Private Sub t_sexo_LostFocus()
t_sexo = Format$(t_sexo, ">@")
If t_sexo <> "H" And t_sexo <> "M" Then
  t_sexo = "M"
End If

  
End Sub
