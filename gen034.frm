VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_padronibp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PROCESO DE ACTUALIZACION DEL PADRON DE PERCEPCIONES DE INGRESOS BRUTOS"
   ClientHeight    =   5160
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5160
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estado del Proceso"
      Height          =   1695
      Left            =   5400
      TabIndex        =   10
      Top             =   240
      Width           =   1935
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label2"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ubicacion definitiva del archivo de origen de retenciones"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   4335
      Begin VB.TextBox t_camino 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SELECCIONE ARCHIVO ORIGEN RETENCIONES"
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   240
         Pattern         =   "*.txt"
         TabIndex        =   6
         Top             =   2880
         Width           =   3135
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   5640
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "gen034.frx":0000
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
         Picture         =   "gen034.frx":0882
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
      Top             =   4905
      Width           =   8820
      _ExtentX        =   15558
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
Attribute VB_Name = "gen_padronibp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim habilitafacturaremito As Boolean
Dim t1 As String



Private Sub btnacepta_Click()
Dim l As String
Dim l2 As String

If verifica Then
 J = MsgBox("Confirma Actualizacion del Padron de Percepciones. ", 4)
 If J = 6 Then
  'borrar bd.actual
  'If abrirconexionib Then
  
   Label1 = "Borrando padron actual.."
   t = 0
   Set rs = New ADODB.Recordset
   q = "select * from i3"
   rs.Open q, cnib, adOpenDynamic, adLockOptimistic
   While Not rs.EOF
    t = t + 1
    Label2 = t
    Label2.Refresh
    rs.Delete
    rs.MoveNext
   Wend
   Set rs = Nothing
 
 
 
   
   t = 0
   Set rs = New ADODB.Recordset
   q = "select * from i3"
   rs.Open q, cnib, adOpenDynamic, adLockOptimistic
   Label1 = "Actualizando padron Percpeciones de IB.."
   Label1.Refresh
   Open t_camino For Input As #1
   While Not EOF(1)
     
    Label2 = t
    Label2.Refresh
    Line Input #1, l

    If t = 0 Then
          Set rs2 = New ADODB.Recordset
          q = "select * from i4"
          rs2.Open q, cnib, adOpenDynamic, adLockOptimistic
         ' rs2.AddNew
            rs2("id_padron") = 1
            rs2("fecha_actualizacion") = Mid$(l, 3, 2) & "/" & Mid$(l, 5, 2) & "/" & Mid$(l, 7, 4)
            rs2("fecha_desde") = Mid$(l, 12, 2) & "/" & Mid$(l, 14, 2) & "/" & Mid$(l, 16, 4)
            rs2("fecha_hasta") = Mid$(l, 21, 2) & "/" & Mid$(l, 23, 2) & "/" & Mid$(l, 25, 4)
            rs2("id_usuario") = para.id_usuario
          rs2.Update
          Set rs2 = Nothing
    End If
    t = t + 1
    rs.AddNew
    rs("id_padron") = 1
    rs("cuit") = Val(Mid$(l, 30, 11))
    rs("tasa_perc") = Val(Mid$(l, 48, 1) & "." & Mid$(l, 50, 2))
    rs("tasa_ret") = 0
    rs("tipo_contribuyente") = Mid$(l, 42, 1)
    rs.Update
    rs.MoveNext
   Wend
   Set rs = Nothing
   Close #1
    
   
    
    
  Else
    MsgBox ("Error al abrir B.D. de Padron IB")
  End If
 ' cnib.Close
 'End If
End If

Label1 = "Fin"
Label2 = ""
End Sub
Function verifica() As Boolean
verifica = True
'On Error GoTo errorib
Open t_camino For Input As #1
Line Input #1, l
If Len(l) <> 55 Then
   MsgBox ("El Archivo DE PERCEPCIONES no parece tener el formato del padron. Verifiquelo")
   verifica = False
   Exit Function
End If
Close #1



Exit Function
errorib:
  MsgBox ("Archivo de PERCEPCIONES Inexistente o Invalido")
  verifica = False
  Close #1
  Exit Function

  
  
End Function
Private Sub btnsale_Click()
Unload Me
End Sub


Sub limpia()
 t_cli = " "
 t_direccion = " "
 t_localidad = " "
 t_cuit = " "
 t_iva = " "
 t_provincia = " "
 t_cp = " "
 t_te = " "
 t_email = " "
 t_saldo1 = " "
 t_saldo2 = " "

 End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
End Sub



Private Sub Drive1_Change()
Dir1.Path = Drive1
Call camino
End Sub



Private Sub File1_Click()
Call camino
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If

End Sub
Sub camino()
t_camino = Dir1 & "\" & File1
Label1 = "En espera..."
Label2 = ""
End Sub


Sub carga()
Call limpia
If Val(t_id) > 0 Then
   Set cl_cli = New Clientes
   cl_cli.carga (Val(t_id))
   If cl_cli.id > 0 Then
     t_cli = cl_cli.razonsocial
     t_direccion = cl_cli.direccion
     t_localidad = cl_cli.localidad
     t_cuit = cl_cli.CUIT
     t_cp = cl_cli.cp
     t_provincia = cl_cli.provincia
     t_email = cl_cli.email
     t_te = cl_cli.te
     t_saldo1 = Format$(cl_cli.saldo(True, Now, True), "######0.00")
     t_saldo2 = Format$(cl_cli.saldo(True, Now, False), "######0.00")
     t_iva = cl_cli.abreviatura_tipoiva
   End If
   Set cl_cli = Nothing
End If
End Sub



Private Sub Form_Load()
Call camino
End Sub
