VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form vta_importalistaprecios 
   Caption         =   "Importa lista de precios de otro sistema GESTIONE"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Importar"
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seleccionar Base de Datos GESTIONE origen"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.TextBox t_path 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   6255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Examinar"
         Height          =   375
         Left            =   6600
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "ATENCION!!!: SALGA DEL SISTEMA y VUELVA A INGRESAR PARA QUE LOS CAMBIOS TENGAN EFECTO"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Width           =   8895
   End
   Begin VB.Label Label2 
      Caption         =   "Tenga en cuenta que tambien migra Grupos, Departamentos, Marcas, y  Proveedores. "
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   8895
   End
   Begin VB.Label Label1 
      Caption         =   $"vta054.frx":0000
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   8895
   End
End
Attribute VB_Name = "vta_importalistaprecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
X = seleccion(t_path)
End Sub

Private Sub Command2_Click()
J = MsgBox("Confirma ejecutar este proceso", 4)
If J = 6 Then
 k = MsgBox("Por seguridad le preguntamos otra vez, Confirma ejecutar este proceso", 4)
If k = 6 Then
  If para.id_grupo_modulo_actual >= 9 Then
    MsgBox ("Cierre todas las terminales, cierre todos los procesos en ejecucion y acepte")
    Call importa
  Else
    Call sinpermisos
  End If
End If
End If

End Sub

Function seleccion(filename As String) As Boolean
On Error GoTo err_sel
CommonDialog1.Filter = "Apps *.mdb"
CommonDialog1.DefaultExt = "mdb"
CommonDialog1.DialogTitle = "Selecciona Archivo"
CommonDialog1.InitDir = "C:\"
CommonDialog1.filename = filename
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
t_path = filename

Exit Function
err_sel:
t_path = filename
End Function

Sub importa()
If t_path <> "" Then
  On Error GoTo err1
  Load espere
  
  'PRECIOS
  espere.Label1 = "Espere.... Borrando lista de precios actual"
  espere.Show

   cn1.BeginTrans
   
   q = "delete from a2"
   cn1.Execute q

   espere.Label1 = "Espere.... Importando lista de precios actual"
   espere.Refresh

   q = "INSERT INTO A2 SELECT * FROM [" & t_path & "].A2"
   cn1.Execute q
   
   cn1.CommitTrans
   
   
   cn1.BeginTrans
   'GRUPOS
   espere.Label1 = "Espere.... Borrando Grupos"
   espere.Label1.Refresh
   q = "delete from a8"
   cn1.Execute q

   espere.Label1 = "Espere.... Importando Grupos actual"
   espere.Refresh

   q = "INSERT INTO A8 SELECT * FROM [" & t_path & "].A8"
   cn1.Execute q
   cn1.CommitTrans
   
   
   
   cn1.BeginTrans
   
   'departametos
   espere.Label1 = "Espere.... Borrando Departamentos"
   espere.Label1.Refresh
   q = "delete from a9"
   cn1.Execute q

   espere.Label1 = "Espere.... Importando Departamentos actual"
   espere.Refresh

   q = "INSERT INTO A9 SELECT * FROM [" & t_path & "].A9"
   cn1.Execute q
   
   'marcas
   espere.Label1 = "Espere.... Borrando Marcas"
   espere.Label1.Refresh
   q = "delete from a10"
   cn1.Execute q

   espere.Label1 = "Espere.... Importando Marcas actual"
   espere.Refresh

   q = "INSERT INTO A10 SELECT * FROM [" & t_path & "].A10"
   cn1.Execute q
   
   'proveedores
   espere.Label1 = "Espere.... Borrando Proveedores"
   espere.Label1.Refresh
   q = "delete from a1"
   cn1.Execute q

   espere.Label1 = "Espere.... Importando Proveedores actual"
   espere.Refresh

   q = "INSERT INTO A1 SELECT * FROM [" & t_path & "].A1"
   cn1.Execute q
   
   
      
   cn1.CommitTrans
   
   
   
   
   
   
   
   
   Unload espere

   MsgBox ("Proceso terminado")
End If
  
Exit Sub

err1:
   MsgBox ("Error en el proceso de importacion. Verifique la BD origen")
   Exit Sub
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

