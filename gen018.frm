VERSION 5.00
Begin VB.Form gen_duplica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicado Sistema"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Duplicar"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unidad Destino"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Este modulo genera un duplicado exacto del sistema en otra unidad para que pueda ser resguradado o trasnportado a otra maquina."
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "gen_duplica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
J = MsgBox("Cierre el sistema en todas las terminales de la red y confirme", 4)
If J = 6 Then
  Dir1 = Drive1 & "\"
  g = MsgBox("Confirma transferir duplicar el sistema de la unidad [" & App.Path & "] a la unidad [" & Dir1 & "]", 4)
  If g = 6 Then
    'On Error GoTo errbak
    espere.Show
    espere.ProgressBar1.Min = 0
    espere.ProgressBar1.Max = 3
    espere.Label1 = "Espere... Borrando Unidad destino"
    espere.ProgressBar1.Value = 1
    espere.Refresh
    F = Eliminar_Directorio
    espere.Label1 = "Espere... Duplicando Sistema"
    espere.Refresh
    espere.ProgressBar1.Value = 2
    Dim resguardo As New Scripting.FileSystemObject
    resguardo.CopyFolder App.Path, Dir1, True
    Unload espere
    MsgBox ("Operacion terminada con Exito!!!")
    
    Set cn1 = New ADODB.Connection
    gconexion = "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & App.Path & "\dat\5a11.mdb;User id=" & "claudio" & ";password=0969" & ";" & "Jet OLEDB:System database=" & App.Path & "\SEG\system2.mdw;"
    cn1.Open gconexion
  
    Set rs = New ADODB.Recordset
    q = "select * from g0 where [id_sucursal] = 0 "
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    rs("ultima_copia") = Format$(Now, "dd/mm/yyyy")
    rs.Update
    Set rs = Nothing
    cn1.Close
    
    End
  End If
End If
Exit Sub
errbak:
 MsgBox ("¡¡¡Error!!! El Duplicado no pudo realizarse")
 End
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
 On Error GoTo ERRDSK
 Dir1 = Drive1.Drive

Exit Sub
ERRDSK:
 MsgBox ("Error en la Unidad Seleccionada")
 Drive1.Drive = "C:"
 
End Sub

Private Sub Form_Load()
Drive1.Drive = "C:"

End Sub
Function Eliminar_Directorio() As Boolean
  
On Error GoTo Error_Sub
 
'Variable de tipo file System Object
Dim fso As FileSystemObject
  
'Creamos la Nueva referencia Fso
Set fso = New FileSystemObject

'Le pasamos a DeleTeFolder el Path a eliminar
 fso.DeleteFolder Drive1.Drive & "\5A11", True
  
 If Err.Number = 0 Then
       ' Ok
     Eliminar_Directorio = True
     Set fso = Nothing
  
 End If
       
Exit Function
Error_Sub:
 Resume Next
  
End Function

