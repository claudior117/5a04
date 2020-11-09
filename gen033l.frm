VERSION 5.00
Begin VB.Form gen_cambiamemoriafiscal 
   Caption         =   "Cambio de Memoria Fiscal"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_pv2 
      Height          =   285
      Left            =   7200
      MaxLength       =   4
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox t_pv1 
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9000
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "gen033l.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "gen033l.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Punto de venta Destino"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Punto de venta Origen"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"gen033l.frx":1104
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   7695
   End
End
Attribute VB_Name = "gen_cambiamemoriafiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnacepta_Click()
If verifica Then
 J = MsgBox("Confirma realizar esta operacion", 4)
 If J = 6 Then
  espere.Show
  Set rs = New ADODB.Recordset
    q = "select * from vta_02 where [sucursal_ingreso] = " & Val(t_pv1)
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
      'rs("sucursal_ingreso") = Val(t_pv2)
      If rs("id_tipocomp") <= 10 Then
        rs("sucursal") = Val(t_pv2)
      End If
      rs.MoveNext
  Wend
  Set rs = Nothing
  Unload espere
  MsgBox ("Proceso terminado")
 End If
End If
End Sub

Function verifica() As Boolean
v = True
Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & Val(t_pv1)
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
     
Else
   MsgBox ("El punto de venta origen no esta creado.Imposible realizar la operacion ")
   v = False
End If
Set rs = Nothing
Set rs = New ADODB.Recordset
q = "select * from vta_06 where [sucursal] = " & Val(t_pv2)
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
     
Else
   MsgBox ("El punto de venta destino no esta creado.Imposible realizar la operacion ")
   v = False
End If
Set rs = Nothing

verifica = v
End Function

Private Sub btnsale_Click()
Unload Me
End Sub
