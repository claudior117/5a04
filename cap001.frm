VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form com_cierremes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CIERRE DE FIN DE MES"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3945
   ScaleWidth      =   7170
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox c_estado 
         Height          =   315
         ItemData        =   "cap001.frx":0000
         Left            =   1560
         List            =   "cap001.frx":000A
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox c_año 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox c_mes 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Este proceso permite cerrar un determinado periodo para evitar que se agreguen movimientos al mismo una vez finalizado."
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Estado:"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Año:"
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
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Mes:"
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
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   5280
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "cap001.frx":0020
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
         Picture         =   "cap001.frx":08A2
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
      Top             =   3690
      Width           =   7170
      _ExtentX        =   12647
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
            TextSave        =   "09:43"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "com_cierremes"
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
J = MsgBox("Confirma Cambiar estado para el Periodo seleccionado", 4)
If J = 6 Then
   On Error GoTo ERRORGRABA
   periodo = Val(Format$(c_año, "0000") & Format$(c_mes, "00"))
   q = "select * from  a14 where [id_periodo] = " & periodo
   Set rs = New ADODB.Recordset
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   If Not rs.EOF And Not rs.BOF Then
     'existe
      rs("estado") = Mid$(c_estado, 1, 1)
      rs.Update
   Else
      rs.AddNew
       rs("id_periodo") = periodo
       rs("mes") = Val(c_mes)
       rs("año") = Val(c_año)
       rs("estado") = Mid$(c_estado, 1, 1)
       rs.Update
   End If
   MsgBox ("El estado del periodo ha sido cambiado")
   Unload Me
End If

Exit Sub
ERRORGRABA:
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos. Modulo: Graba")
  
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub c_cuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  c_cuenta.ListIndex = 0
End If
End Sub

Private Sub c_año_LostFocus()
If c_año.ListIndex < 0 Then
  Call actual
End If

End Sub

Private Sub c_estado_Change()
If c_estado.ListIndex < 0 Then
  c_estado.ListIndex = 0
End If

End Sub

Private Sub c_mes_LostFocus()
If c_mes.ListIndex < 0 Then
  Call actual
End If
End Sub

Private Sub Form_Load()
Call barraesag(Me)
c_mes.clear
c_año.clear
For i = 1 To 12
   c_mes.AddItem i, i - 1
Next i

For i = 0 To 100
  c_año.AddItem i + 1990, i
Next i
c_año.ListIndex = 0
c_mes.ListIndex = 0
c_estado.ListIndex = 0
Call actual

End Sub
Sub actual()

m = Val(Mid$(Format$(Date, "dd/mm/yyyy"), 4, 2))
a = Val(Mid$(Format$(Date, "dd/mm/yyyy"), 7, 7))
c_mes.ListIndex = m - 1
c_año.ListIndex = (a - 1990)



End Sub



