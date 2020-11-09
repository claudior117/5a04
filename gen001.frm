VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gen_borradatos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BORRA DATOS DEL SISTEMA"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4560
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   4095
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox Check10 
         Caption         =   "Contabilidad"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Produccion"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   495
         Left            =   3120
         TabIndex        =   11
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Mov. Compras"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Mov. Ventas"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Caja"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Cheques/Bancos"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Stock"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Proveedores"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Productos"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Clientes"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4305
      Width           =   6705
      _ExtentX        =   11827
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
            TextSave        =   "21/01/2010"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "06:33 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "gen_borradatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984



Private Sub Command1_Click()
J = InputBox$("Ingrese Clave de Administrador")
If J = "0969" Then
   
   If Check1 = 1 Then
      'clientes
      q = "delete * from vta_01"
      cn1.Execute q
      
      q = "delete * from vta_02"
      cn1.Execute q
      
      q = "delete * from vta_03"
      cn1.Execute q

      q = "delete * from vta_04"
      cn1.Execute q

   End If
   
   
   If Check7 = 1 Then
      q = "delete * from vta_02"
      cn1.Execute q
      
      q = "delete * from vta_03"
      cn1.Execute q

      q = "delete * from vta_04"
      cn1.Execute q
      
      
      q = "delete * from vta_07"
      cn1.Execute q
      
      q = "delete * from vta_08"
      cn1.Execute q
            
   End If
   
   
   If Check2 = 1 Then
      q = "delete * from a2"
      cn1.Execute q
      
      q = "delete * from vta_03"
      cn1.Execute q

      q = "delete * from a6"
      cn1.Execute q
            
      q = "delete * from stk_01"
      cn1.Execute q
   End If


   If Check3 = 1 Then 'proveedores
      
      q = "delete * from a1"
      cn1.Execute q
      
      q = "delete * from a5"
      cn1.Execute q
      
      q = "delete * from a6"
      cn1.Execute q
            
      q = "delete * from a7"
      cn1.Execute q
               
   End If
  
   If Check8 = 1 Then 'mov. compras
      
      
      q = "delete * from a5"
      cn1.Execute q
      
      q = "delete * from a6"
      cn1.Execute q
            
      q = "delete * from a7"
      cn1.Execute q
      
      q = "delete * from ret_01"
      cn1.Execute q
      
               
   End If
  
   
   If Check4 = 1 Then 'stock
      
      q = "delete * from stk_01"
      cn1.Execute q

   End If

   If Check5 = 1 Then 'cheques
      
      q = "delete * from cyb_02"
      cn1.Execute q
      
      q = "delete * from cyb_03"
      cn1.Execute q
      
      q = "delete * from cyb_04"
      cn1.Execute q
      
      q = "delete * from cyb_05"
      cn1.Execute q
      
      
   End If

 
   If Check6 = 1 Then 'caja
      q = "delete * from cyb_05"
      cn1.Execute q
   End If

   If Check9 = 1 Then 'produccion
      q = "delete * from a11"
      cn1.Execute q
   
      q = "delete * from a13"
      cn1.Execute q
      
      q = "delete * from a14"
      cn1.Execute q
      
      q = "delete * from pro_01"
      cn1.Execute q
      
      q = "delete * from pro_02"
      cn1.Execute q
  
      q = "delete * from pro_04"
      cn1.Execute q
      
      q = "delete * from pro_05"
      cn1.Execute q
  
  
   End If

  If Check10 = 1 Then 'contabilidad
      q = "delete * from c_02"
      cn1.Execute q
   
      q = "delete * from c_03"
      cn1.Execute q
  End If

  MsgBox ("Proceso Terminado")
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 27
        
        Me.Hide
End Select
End Sub
Sub carga()
vta_recibo.armagrid
k = 0
r = 1
While k < List1.ListCount
  If List1.Selected(k) = True Then
   f = Mid$(List1.List(k), 1, 10)
   c = Mid$(List1.List(k), 11, 21)
   vta_recibo.msf1.AddItem f & Chr(9) & c & Chr(9) & Mid$(List1.List(k), 33, 10) & Chr(9) & Mid$(List1.List(k), 45, 3) & Chr(9) & Mid$(List1.List(k), 51, 8) & Chr(9) & Mid$(List1.List(k), 61, 8)
   r = r + 1
  End If
   k = k + 1
Wend

   
End Sub
Private Sub Form_Load()
Call barraesag(Me)

End Sub

  
Private Sub List1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[ESC] Termina - [F9] Agrega "

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
  Call carga
  Me.Hide
End If
End Sub
