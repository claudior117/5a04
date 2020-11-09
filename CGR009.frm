VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_cuentas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODIFICA CUENTAS CONTABLES"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   11
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
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   8295
      Begin VB.ComboBox c_tipocaja 
         Height          =   315
         ItemData        =   "CGR009.frx":0000
         Left            =   2160
         List            =   "CGR009.frx":0010
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   0
         Top             =   720
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
         Left            =   480
         TabIndex        =   13
         Top             =   1080
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
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6840
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "CGR009.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CGR009.frx":08D6
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   2
      Top             =   3210
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
Attribute VB_Name = "cgr_cuentas"
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
   
   If t_funcion = "M" Then
      
      QUERY = "update C_01 set  [descripcion]='" & Left$(RTrim$(t_descripcion), 40) & "', [tipo_cuentacaja]='" & tc & "'"
      QUERY = QUERY & " where [id_Cuenta]= " & Val(t_id)
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans
   Else
      'altas
      QUERY = "INSERT INTO c_01([id_cuenta], [DEscripcion], [pos1], [pos2], [pos3], [pos4], [tipo],[tipo_cuentacaja])"
      QUERY = QUERY & " VALUES (" & Val(t_id) & ", '" & t_descripcion & "', " & Val(Mid$(t_id, 1, 1)) & ", " & Val(Mid$(t_id, 2, 1)) & ", " & Val(Mid$(t_id, 3, 2)) & ", " & Val(Mid$(t_id, 5, 2)) & ", 'T', 'N')"
      cn1.BeginTrans
      cn1.Execute QUERY
      cn1.CommitTrans

   
   End If
   
   Call CGR_CUENTAS0.limpia
   CGR_CUENTAS0.Show
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







Private Sub c_tipocaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   btnacepta.SetFocus
End If
End Sub

Private Sub c_tipocaja_LostFocus()
If c_tipocaja.ListIndex < 0 Then
  c_tipocaja.ListIndex = 0
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

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 1)
  Case Is = 27
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
c_tipocaja.ListIndex = 0
End Sub



Private Sub t_descripcion_LostFocus()
Call NULOS(t_descripcion)
End Sub






