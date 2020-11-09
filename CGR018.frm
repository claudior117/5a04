VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_cierreyapertura 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ASIENTOS CONTABLES DE CIERRE y APERTURA"
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
   Begin VB.Frame Frame7 
      Height          =   1215
      Left            =   240
      TabIndex        =   20
      Top             =   6840
      Width           =   9735
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   26
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1815
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1935
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
         Left            =   1560
         TabIndex        =   25
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
         Left            =   5040
         TabIndex        =   23
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "HABER"
      Height          =   4215
      Left            =   6000
      TabIndex        =   19
      Top             =   2520
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid msf2 
         Height          =   3855
         Left            =   120
         TabIndex        =   6
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
      TabIndex        =   18
      Top             =   2520
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   3855
         Left            =   120
         TabIndex        =   5
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
      Height          =   1815
      Left            =   8640
      TabIndex        =   13
      Top             =   120
      Width           =   3255
      Begin VB.TextBox t_fechaa 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   31
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox t_fechac 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Apert.:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Fecha Cierre:"
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
         Left            =   120
         TabIndex        =   30
         Top             =   960
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
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
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asiento"
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8535
      Begin VB.TextBox t_descapertura 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   49
         TabIndex        =   3
         Text            =   " Apertura Ejercicio"
         Top             =   1440
         Width           =   5895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GENERA ASIENTO CIERRE"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   7935
      End
      Begin VB.ComboBox c_apertura 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   600
         Width           =   6255
      End
      Begin VB.ComboBox c_cierre 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   240
         Width           =   6255
      End
      Begin VB.TextBox t_descripciong 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   49
         TabIndex        =   2
         Text            =   "Cierre  Ejercicio"
         Top             =   1080
         Width           =   5895
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Descrip. Apertura:"
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
         TabIndex        =   33
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Periodo Apertura:"
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
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Periodo Cierre:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Descrip. Cierre:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10080
      TabIndex        =   8
      Top             =   7080
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR018.frx":0000
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
         Picture         =   "CGR018.frx":0882
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:39"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cgr_cierreyapertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Private gfcierreinicio As Date
Private gfcierrefin As Date
Private gfaperturafin As Date
Private gfaperturainicio As Date




Private Sub btnacepta_Click()
If Val(t_diferencia) = 0 Then
 Call graba
Else
 MsgBox ("El asiento tiene diferencias. Imposible grabar operacion")
End If
End Sub

Sub graba()
J = MsgBox("Confirma Valores para Grabar. Recuerde que se realizan ambos asientos Cierre/Apertura y se cerrará el periodo ", 4)

If J = 6 Then
     
   On Error GoTo ERRORGRABA
            
  espere.Show
  espere.Label1 = "Espere Grabando Asientos...."
          
       'saco numero asiento cierre
        a = Format$(Val(Mid$(gfcierrefin, 7, 4)), "0000")
        m = Format$(Val(Mid$(gfcierrefin, 4, 2)), "00")
       a1 = Val(a & m & "000")
       a2 = Val(a & m & "999")
      
       Set rs = New ADODB.Recordset
       q = "select * from c_11 where [año] = " & Val(a) & " and [mes] = " & Val(m)
       rs.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs.EOF And Not rs.BOF Then
         rs.MoveLast
         na = rs("num_asiento") + 1
       Else
         na = Val(a & m & "001")
       End If
       Set rs = Nothing
      
      cn1.BeginTrans
      QUERY = "INSERT INTO c_11([num_asiento], [fecha], [descripcion], [id_periodo], [importe], [año], [mes])"
      QUERY = QUERY & " VALUES (" & na & ", '" & gfcierrefin & "', '" & t_descripciong & "', " & c_cierre.ItemData(c_cierre.ListIndex) & ", " & Val(t_totdebe) & ", " & Val(a) & ", " & Val(m) & ")"
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs = cn1.Execute(qr)
      nic = rs.Fields("NewID").Value

      
      s = 1
      For i = 1 To msf1.Rows - 1
        QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
        QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & Val(msf1.TextMatrix(i, 1)) & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & msf1.TextMatrix(i, 2) & "', 'D')"
        cn1.Execute QUERY
        s = s + 1
      Next i
      
      For i = 1 To msf2.Rows - 1
        QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
        QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & Val(msf2.TextMatrix(i, 1)) & ", " & Val(msf2.TextMatrix(i, 3)) & ", '" & msf2.TextMatrix(i, 2) & "', 'H')"
        cn1.Execute QUERY
        s = s + 1
      Next i
      
      QUERY = "update c_10 set  [estado]='C'"
      QUERY = QUERY & " where [id_periodo]= " & c_cierre.ItemData(c_cierre.ListIndex)
      cn1.Execute QUERY
      
      cn1.CommitTrans
   
   
     'saco numero asiento apertura
        a = Format$(Val(Mid$(gfaperturainicio, 7, 4)), "0000")
        m = Format$(Val(Mid$(gfaperturainicio, 4, 2)), "00")
       a1 = Val(a & m & "000")
       a2 = Val(a & m & "999")
      
       Set rs = New ADODB.Recordset
       q = "select * from c_11 where [año] = " & Val(a) & " and [mes] = " & Val(m)
       rs.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs.EOF And Not rs.BOF Then
         rs.MoveLast
         na = rs("num_asiento") + 1
       Else
         na = Val(a & m & "001")
       End If
       Set rs = Nothing
      
      
      cn1.BeginTrans
      QUERY = "INSERT INTO c_11([num_asiento], [fecha], [descripcion], [id_periodo], [importe], [año], [mes])"
      QUERY = QUERY & " VALUES (" & na & ", '" & gfaperturainicio & "', '" & t_descapertura & "', " & c_apertura.ItemData(c_apertura.ListIndex) & ", " & Val(t_totdebe) & ", " & Val(a) & ", " & Val(m) & ")"
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs = cn1.Execute(qr)
      nic = rs.Fields("NewID").Value

      
      s = 1
      For i = 1 To msf1.Rows - 1
        QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
        QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & Val(msf1.TextMatrix(i, 1)) & ", " & Val(msf1.TextMatrix(i, 3)) & ", '" & t_descapertura & "', 'H')"
        cn1.Execute QUERY
        s = s + 1
      Next i
      
      For i = 1 To msf2.Rows - 1
        QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
        QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & Val(msf2.TextMatrix(i, 1)) & ", " & Val(msf2.TextMatrix(i, 3)) & ", '" & t_descapertura & "', 'D')"
        cn1.Execute QUERY
        s = s + 1
      Next i
      cn1.CommitTrans
      
      Unload espere
  End If
  Unload Me
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
msf1.Cols = 5
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 400
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 2200
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 1600
msf1.TextMatrix(0, 0) = ""
msf1.TextMatrix(0, 1) = "Id.Cuenta"
msf1.TextMatrix(0, 2) = "Descripcion"
msf1.TextMatrix(0, 3) = "Importe"
msf1.TextMatrix(0, 4) = "Desc. Cuenta"
For i = 0 To 4
 msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(3) = 9

msf1.FocusRect = flexFocusNone

End Sub

Sub armagrid2()
msf2.clear
msf2.Rows = 1
msf2.Cols = 5
msf2.AllowUserResizing = flexResizeNone
msf2.FixedCols = 0
msf2.SelectionMode = flexSelectionByRow
msf2.FocusRect = flexFocusNone
msf2.ColWidth(0) = 400
msf2.ColWidth(1) = 800
msf2.ColWidth(2) = 2200
msf2.ColWidth(3) = 1000
msf2.ColWidth(4) = 1600
msf2.TextMatrix(0, 0) = ""
msf2.TextMatrix(0, 1) = "Cuenta"
msf2.TextMatrix(0, 2) = "Descripcion"
msf2.TextMatrix(0, 3) = "Importe"
msf2.TextMatrix(0, 4) = "Desc.Cuenta"

For i = 0 To 4
 msf2.ColAlignment(i) = 1 'izq
Next i
msf2.ColAlignment(3) = 9

msf2.FocusRect = flexFocusNone

End Sub

Private Sub c_apertura_LostFocus()
If c_apertura.ListIndex < 0 Then
  c_apertura.ListIndex = 0
End If

End Sub

Private Sub c_cierre_LostFocus()
If c_cierre.ListIndex < 0 Then
  c_cierre.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
If verifica2 Then
 J = MsgBox("Confirma generar asiento de Cierre e Ejercicio", 4)
 If J = 6 Then
  Call armagrid
  Call armagrid2
  Load espere
  t = 1
  m = "Espere procesando cuentas... "
  espere.Label1 = m & t
  espere.Refresh
  q = " select * from c_01 order by [id_cuenta]"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  While Not rs.EOF
   s = sacasaldo(rs("id_cuenta"))
   If s <> 0 Then
      If s > 0 Then
        'saldo deudor va al haber
         msf2.AddItem msf2.Rows & Chr$(9) & rs("id_cuenta") & Chr$(9) & t_descripciong & Chr$(9) & Format$(s, "######0.00") & Chr$(9) & rs("descripcion")
      Else
         'saldo acreedor va al debe
         msf1.AddItem msf1.Rows & Chr$(9) & rs("id_cuenta") & Chr$(9) & t_descripciong & Chr$(9) & Format$(-s, "######0.00") & Chr$(9) & rs("descripcion")
      End If
   End If
    t = t + 1
    espere.Label1 = m & t
    espere.Label1.Refresh
  
   rs.MoveNext
  Wend

  Call calcula_totales
  Unload espere
 End If
End If
End Sub
Function verifica2() As Boolean
v = True
If c_cierre.ListIndex = c_apertura.ListIndex Then
  MsgBox ("Error. El periodo de Cierre no puede ser el mismo que el periodo de Apertura")
  v = False
End If

'saco fechas del periodo
If v = True Then
  Set rs = New ADODB.Recordset
  q = "select * from c_10 where [id_periodo] = " & c_cierre.ItemData(c_cierre.ListIndex)
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
    t_fechac = rs("fecha_cierre")
    gfcierreinicio = rs("fecha_inicio")
    gfcierrefin = rs("fecha_cierre")
    If rs("estado") = "C" Then
      MsgBox ("El periodo de Cierre ya esta cerrado")
      v = False
    End If
  Else
    MsgBox ("Error en el periodo contable de Cierre")
    v = False
  End If
  Set rs = Nothing

  Set rs = New ADODB.Recordset
  q = "select * from c_10 where [id_periodo] = " & c_apertura.ItemData(c_apertura.ListIndex)
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
    t_fechaa = rs("fecha_inicio")
    gfaperturainicio = rs("fecha_inicio")
    gfaperturafin = rs("fecha_cierre")
    If rs("estado") = "C" Then
      MsgBox ("El periodo de Apertura esta cerrado")
      v = False
    End If
  Else
    MsgBox ("Error en el periodo contable de Apertura")
    v = False
  End If
  Set rs = Nothing

  If DateValue(t_fechac) >= DateValue(t_fechaa) Then
    MsgBox ("La fecha de Apertura es Anterior a la fecha de cierre")
    v = False
  End If
End If

verifica2 = v
  

End Function
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
Function sacasaldo(ByVal c As Long) As Double
'saldo anterior
Dim q As String
d = 0
h = 0
dt = 0
ht = 0
st = 0
q = "select * from c_11, c_12 where c_11.[id_asiento] = c_12.[id_asiento] and [id_cuenta] = " & c
q = q & " and datevalue([fecha]) >= datevalue('" & gfcierreinicio & "')"
q = q & " and datevalue([fecha]) <= datevalue('" & gfcierrefin & "')"
q = q & " and  c_11.[id_periodo] = " & c_cierre.ItemData(c_cierre.ListIndex)
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
While Not rs2.EOF
 If rs2("ubicacion") = "D" Then
    d = d + rs2("c_12.importe")
 Else
    h = h + rs2("c_12.importe")
 End If
 rs2.MoveNext
 Wend
 Set rs2 = Nothing
 st = d - h
 sacasaldo = st
End Function

Sub renumera()
 If msf1.Rows > 1 Then
   For i = 1 To msf1.Rows - 1
      msf1.TextMatrix(i, 0) = i
   Next i
 End If
 
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  
End Select
End Sub

Private Sub Form_Load()
Call barracgr(Me)
Call armagrid
Call armagrid2
Load abm_asientos2

Call carga_periodos(c_cierre)
Call carga_periodos(c_apertura)
c_cierre.ListIndex = buscaindice(c_cierre, para.id_periodo_contable)
c_apertura.ListIndex = 0



End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_asientos2
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba - "
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
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba - "
msf1.FocusRect = flexFocusHeavy
End Sub

Private Sub msf2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
  abm_asientos2.Show
  abm_asientos2.limpia
  abm_asientos2.c_ubica.ListIndex = 1
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


