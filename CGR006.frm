VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_asientos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ASIENTOS CONTABLES"
   ClientHeight    =   8460
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8460
   ScaleWidth      =   11850
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   3600
      TabIndex        =   30
      Top             =   7440
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "Renumera Items"
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Acumula Cuentas"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1215
      Left            =   0
      TabIndex        =   23
      Top             =   6000
      Width           =   11775
      Begin VB.TextBox t_diferencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox t_tothaber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_totdebe 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Diferencia (Haber - Debe) ------->"
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
         Left            =   2040
         TabIndex        =   28
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Total HABER ------->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6600
         TabIndex        =   26
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Total DEBE ------->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "HABER"
      Height          =   4215
      Left            =   6000
      TabIndex        =   22
      Top             =   1680
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid msf2 
         Height          =   3855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6800
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DEBE"
      Height          =   4215
      Left            =   0
      TabIndex        =   21
      Top             =   1680
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6800
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   8400
      TabIndex        =   16
      Top             =   240
      Width           =   3255
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_numero 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Id. Asiento:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Numero Asiento:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   7440
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asiento"
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   9600
         Picture         =   "CGR006.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox t_f1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   34
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox t_descripciong 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   49
         TabIndex        =   1
         Top             =   960
         Width           =   5895
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
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
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Descripcion Gral:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   5
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR006.frx":0105
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CGR006.frx":0987
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   4
      Top             =   8205
      Width           =   11850
      _ExtentX        =   20902
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
            TextSave        =   "23/11/2020"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:17 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "abm_asientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Private gf1 As Date
Private gf2 As Date




Private Sub btnacepta_Click()
If verificaperiodog(t_f1) = "A" Then
 Call graba
Else
 MsgBox ("Periodo Cerrado. Imposible grabar operacion")
End If
End Sub

Sub graba()
J = MsgBox("Confirma Valores para Grabar", 4)
If J = 6 Then
   'On Error GoTo ERRORGRABA
   If verifica Then
     Select Case t_funcion
     Case Is = "A", Is = "M"
      If t_funcion = "M" Then
        QUERY = "DELETE FROM c_11 WHERE [id_asiento] = " & Val(t_id)
        cn1.BeginTrans
        cn1.Execute QUERY
        cn1.CommitTrans
 
        na = Val(t_numero)
      Else
            
        'saco numero
        a = Format$(Val(Mid$(t_f1, 7, 4)), "0000")
        m = Format$(Val(Mid$(t_f1, 4, 2)), "00")
       a1 = Val(a & m & "000")
       a2 = Val(a & m & "999")
      
       Set rs = New adodb.Recordset
       q = "select * from c_11 where [año] = " & Val(a) & " and [mes] = " & Val(m)
       rs.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs.EOF And Not rs.BOF Then
         rs.MoveLast
         na = rs("num_asiento") + 1
       Else
         na = Val(a & m & "001")
       End If
       Set r5s = Nothing
      End If
      
      cn1.BeginTrans
      QUERY = "INSERT INTO c_11([num_asiento], [fecha], [descripcion], [id_periodo], [importe], [año], [mes])"
      QUERY = QUERY & " VALUES (" & na & ", '" & t_f1 & "', '" & t_descripciong & "', " & para.id_periodo_contable & ", " & Val(t_totdebe) & ", " & Val(a) & ", " & Val(m) & ")"
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs = cn1.Execute(qr)
      nic = rs.Fields("NewID").Value

      
      s = 1
      For i = 1 To msf1.Rows - 1
        QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
        QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & Val(msf1.TextMatrix(i, 1)) & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & Left$(msf1.TextMatrix(i, 2) & " " & msf1.TextMatrix(i, 5), 80) & "', 'D')"
        cn1.Execute QUERY
        s = s + 1
      Next i
      
      For i = 1 To msf2.Rows - 1
        QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
        QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & Val(msf2.TextMatrix(i, 1)) & ", " & Val(msf2.TextMatrix(i, 3)) & ", '" & Left$(msf2.TextMatrix(i, 2) & " " & msf2.TextMatrix(i, 5), 80) & "', 'H')"
        cn1.Execute QUERY
        s = s + 1
      Next i
      cn1.CommitTrans
      
   End Select
   Call limpia
    
  Else
    MsgBox ("Los datos no estan completos. No se pudo actualizar")
  
  End If
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub
Sub limpia()
Call armagrid
Call armagrid2
t_descripciong = ""
t_totdebe = ""
t_tothaber = ""
t_diferencia = ""
't_f1.SetFocus

End Sub
Function verifica() As Boolean
  v = True
  'verifco fechas
  If DateValue(t_f1) < DateValue(gf1) Or DateValue(t_f1) > DateValue(gf2) Then
    v = False
    MsgBox ("La fecha ingresada no esta dentro del periodo contable")
  End If
  
  If Val(t_diferencia) <> 0 Then
    v = False
    MsgBox ("El asient0 NO cumple con la partida doble")
  End If
  verifica = v
End Function
Private Sub btnsale_Click()
Me.Hide
End Sub









Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 400
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 2200
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 1600
msf1.ColWidth(5) = 1600
msf1.TextMatrix(0, 0) = ""
msf1.TextMatrix(0, 1) = "Id.Cuenta"
msf1.TextMatrix(0, 2) = "Descripcion"
msf1.TextMatrix(0, 3) = "Importe"
msf1.TextMatrix(0, 4) = "Desc. Cuenta"
msf1.TextMatrix(0, 5) = "Obs"
For i = 0 To 4
 msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(3) = 9

msf1.FocusRect = flexFocusNone

End Sub
Sub acumula()
 'debe
 
 If msf1.Rows > 2 Then
   For i = 1 To msf1.Rows - 2
      cb = Val(msf1.TextMatrix(i, 1))
      If cb > 0 Then
        tc = Val(msf1.TextMatrix(i, 3))
        msf1.TextMatrix(i, 2) = msf1.TextMatrix(i, 4)
        'busco si la cuenta cb se repite
        For J = i + 1 To msf1.Rows - 1
           If cb = Val(msf1.TextMatrix(J, 1)) Then
             'si encuentro la cuenta sumo y pongo idcuenta en 0
             tc = tc + Val(msf1.TextMatrix(J, 3))
             msf1.TextMatrix(J, 1) = 0 'para saber que cuenta modifique
           End If
        Next J
        msf1.TextMatrix(i, 3) = tc
      End If
    
   Next i
   
   'saca las cuenta en 0
   s = 0
   While s = 0
    s = 1
    For i = 1 To msf1.Rows - 1
      cb = Val(msf1.TextMatrix(i, 1))
      If cb = 0 Then
        msf1.RemoveItem i
        i = msf1.Rows
        s = 0
      End If
    Next i
   Wend
   
 End If
 
 
 
 'haber
 
 If msf2.Rows > 2 Then
   For i = 1 To msf2.Rows - 2
      cb = Val(msf2.TextMatrix(i, 1))
      If cb > 0 Then
        tc = Val(msf2.TextMatrix(i, 3))
         msf2.TextMatrix(i, 2) = msf2.TextMatrix(i, 4)
        'busco si la cuenta cb se repite
        For J = i + 1 To msf2.Rows - 1
           If cb = Val(msf2.TextMatrix(J, 1)) Then
             'si encuentro la cuenta sumo y pongo idcuenta en 0
             tc = tc + Val(msf2.TextMatrix(J, 3))
             msf2.TextMatrix(J, 1) = 0 'para saber que cuenta modifique
           End If
        Next J
        msf2.TextMatrix(i, 3) = tc
      End If
    
   Next i
   
   'saca las cuenta en 0
   s = 0
   While s = 0
    s = 1
    For i = 1 To msf2.Rows - 1
      cb = Val(msf2.TextMatrix(i, 1))
      If cb = 0 Then
        msf2.RemoveItem i
        i = msf2.Rows
        s = 0
      End If
    Next i
   Wend
   
   Call renumera
 End If
End Sub
Sub armagrid2()
msf2.clear
msf2.Rows = 1
msf2.Cols = 6
msf2.AllowUserResizing = flexResizeNone
msf2.FixedCols = 0
msf2.SelectionMode = flexSelectionByRow
msf2.FocusRect = flexFocusNone
msf2.ColWidth(0) = 400
msf2.ColWidth(1) = 800
msf2.ColWidth(2) = 2200
msf2.ColWidth(3) = 1000
msf2.ColWidth(4) = 1600
msf2.ColWidth(5) = 1600
msf2.TextMatrix(0, 0) = ""
msf2.TextMatrix(0, 1) = "Cuenta"
msf2.TextMatrix(0, 2) = "Descripcion"
msf2.TextMatrix(0, 3) = "Importe"
msf2.TextMatrix(0, 4) = "Desc.Cuenta"
msf2.TextMatrix(0, 5) = "Obs"
For i = 0 To 4
 msf2.ColAlignment(i) = 1 'izq
Next i
msf2.ColAlignment(3) = 9

msf2.FocusRect = flexFocusNone

End Sub

Private Sub Command1_Click()
Call acumula

End Sub

Private Sub Command3_Click()
Call renumera
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   Case Is = vbKeyF9
     Call graba
         
End Select

End Sub
Sub calcula_totales()
t_totdebe = Format$(suma_msflexgrid(msf1, 3), "######0.00")
t_tothaber = Format$(suma_msflexgrid(msf2, 3), "######0.00")
t_diferencia = Format$(Val(t_tothaber) - Val(t_totdebe), "######0.00")

End Sub
Sub renumera()
 If msf1.Rows > 1 Then
   For i = 1 To msf1.Rows - 1
      msf1.TextMatrix(i, 0) = i
   Next i
 End If
 
 If msf2.Rows > 1 Then
   For i = 1 To msf2.Rows - 1
      msf2.TextMatrix(i, 0) = i
   Next i
 End If
 
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  
End Select
End Sub

Private Sub Form_Load()
Call barracgr(Me)
Call armagrid
Call armagrid2
Load abm_asientos2

'saco fechas del periodo
Set rs = New adodb.Recordset
q = "select * from c_10 where [id_periodo] = " & para.id_periodo_contable
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  gf1 = rs("fecha_inicio")
  gf2 = rs("fecha_cierre")
Else
  MsgBox ("Error en el periodo contable")
  Unload Me
End If
Set rs = Nothing
End Sub





Private Sub Form_Unload(Cancel As Integer)
Unload abm_asientos2
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba - "
msf1.FocusRect = flexFocusHeavy
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
  abm_asientos2.Show
  abm_asientos2.limpia
  'abm_asientos2.t_renglon = ""
  abm_asientos2.c_ubica.ListIndex = 0
End If

If KeyCode = vbKeyF9 Then
  Call graba
End If


If KeyCode = vbKeyF5 Then
  r = msf1.Row
  p = msf1.Rows
  J = MsgBox("confirma eliminar renglon actual", 4)
  If J = 6 Then
   If p > 2 Then
    msf1.RemoveItem r
    
   Else
    Call armagrid
   End If
    Call renumera
    Call calcula_totales
  End If
End If
End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    abm_asientos2.limpia
    abm_asientos2.t_renglon = msf1.Row
    abm_asientos2.t_cod = msf1.TextMatrix(msf1.Row, 1)
    abm_asientos2.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    abm_asientos2.c_cuenta.ListIndex = buscaindice(abm_asientos2.c_cuenta, Val(msf1.TextMatrix(msf1.Row, 1)))
    abm_asientos2.c_ubica.ListIndex = 0
    abm_asientos2.t_ubicaanterior = "D"
    abm_asientos2.t_importe = msf1.TextMatrix(msf1.Row, 3)
    
    abm_asientos2.Show
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barracgr(Me)
msf1.FocusRect = flexFocusNone

End Sub

Private Sub msf2_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba - "
msf1.FocusRect = flexFocusHeavy
End Sub

Private Sub msf2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
  abm_asientos2.Show
  abm_asientos2.limpia
  abm_asientos2.c_ubica.ListIndex = 1
End If

If KeyCode = vbKeyF5 Then
  r = msf2.Row
  p = msf2.Rows
  J = MsgBox("confirma eliminar renglon actual", 4)
  If J = 6 Then
   If p > 2 Then
    msf2.RemoveItem r
    
   Else
    Call armagrid
   End If
    Call renumera
    Call calcula_totales
  End If
End If

End Sub

Private Sub msf2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf2.Row > 0 Then
    abm_asientos2.limpia
    abm_asientos2.t_renglon = msf2.Row
    abm_asientos2.t_cod = msf2.TextMatrix(msf2.Row, 1)
    abm_asientos2.t_detalle = msf2.TextMatrix(msf2.Row, 2)
    abm_asientos2.c_cuenta.ListIndex = buscaindice(abm_asientos2.c_cuenta, Val(msf2.TextMatrix(msf2.Row, 1)))
    abm_asientos2.c_ubica.ListIndex = 1
    abm_asientos2.t_ubicaanterior = "H"
    abm_asientos2.t_importe = msf2.TextMatrix(msf2.Row, 3)
    
    abm_asientos2.Show
  End If
End If

End Sub

Private Sub msf2_LostFocus()
Call barracgr(Me)
msf2.FocusRect = flexFocusNone
End Sub

Private Sub t_descripciong_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  msf1.SetFocus
End If

End Sub

Private Sub t_descripciong_LostFocus()
Call NULOS(t_descripciong)
End Sub

Private Sub t_f1_LostFocus()
If t_f1 <> "" Then
  If Not IsDate(t_f1) Then
    t_f1 = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_f1 = Format$(Now, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_f1)
End Sub


Private Sub UpDown1_DownClick()
t_f1 = DateValue(t_f1) - 1
End Sub

Private Sub UpDown1_UpClick()
t_f1 = DateValue(t_f1) + 1
End Sub

