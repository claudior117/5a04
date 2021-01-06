VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejemplo VB6 Lib Epson Bajo Nivel | EpsonFiscalDriver"
   ClientHeight    =   9615
   ClientLeft      =   7425
   ClientTop       =   3375
   ClientWidth     =   15765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   15765
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   22
      Top             =   3000
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Operaciones"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmbOpenVoucher"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnOpen"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnItemUp"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnPayment"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "btnClose"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "btnZ"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "btnX"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Consultas"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnErrorDescription"
      Tab(1).Control(1)=   "btnLastError"
      Tab(1).Control(2)=   "btnPrinterStatus"
      Tab(1).Control(3)=   "btnFiscalStatus"
      Tab(1).Control(4)=   "btnGetFPVersion"
      Tab(1).Control(5)=   "btnDllVersion"
      Tab(1).Control(6)=   "Image5"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Contáctanos"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image9"
      Tab(2).Control(1)=   "Image8"
      Tab(2).Control(2)=   "RichTextBox1"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton btnErrorDescription 
         BackColor       =   &H80000002&
         Caption         =   "Descripción de error"
         Height          =   375
         Left            =   -70920
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton btnLastError 
         BackColor       =   &H80000002&
         Caption         =   "Último error"
         Height          =   375
         Left            =   -70920
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton btnPrinterStatus 
         BackColor       =   &H80000002&
         Caption         =   "Estado de Impresora"
         Height          =   375
         Left            =   -72960
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton btnFiscalStatus 
         BackColor       =   &H80000002&
         Caption         =   "Estado Fiscal"
         Height          =   375
         Left            =   -72960
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton btnGetFPVersion 
         BackColor       =   &H80000002&
         Caption         =   "Versión de IF"
         Height          =   375
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton btnDllVersion 
         BackColor       =   &H80000002&
         Caption         =   "Versión de DLL"
         Height          =   375
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton btnX 
         BackColor       =   &H80000002&
         Caption         =   "X"
         Height          =   375
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton btnZ 
         BackColor       =   &H80000002&
         Caption         =   "Z"
         Height          =   375
         Left            =   5460
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton btnClose 
         BackColor       =   &H80000002&
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton btnPayment 
         BackColor       =   &H80000002&
         Caption         =   "Pago"
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton btnItemUp 
         BackColor       =   &H80000002&
         Caption         =   "Item+"
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton btnOpen 
         BackColor       =   &H80000002&
         Caption         =   "Abrir"
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cmbOpenVoucher 
         Height          =   315
         Left            =   180
         TabIndex        =   23
         Text            =   "Tique"
         Top             =   840
         Width           =   3255
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   38
         Top             =   2280
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6165
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"main.frx":0000
      End
      Begin VB.Image Image8 
         Height          =   1770
         Left            =   -75000
         Top             =   360
         Width           =   12270
      End
      Begin VB.Image Image5 
         Height          =   5700
         Left            =   -74940
         Stretch         =   -1  'True
         Top             =   360
         Width           =   11145
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes de cierre"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   31
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de comprobante"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   480
         Width           =   2175
      End
      Begin VB.Image Image4 
         Height          =   5700
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   11265
      End
      Begin VB.Image Image9 
         Height          =   5700
         Left            =   -74940
         Stretch         =   -1  'True
         Top             =   360
         Width           =   11145
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log"
      Height          =   7695
      Left            =   11400
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
      Begin VB.CommandButton btnClearLog 
         BackColor       =   &H80000002&
         Caption         =   "Borrar log"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7200
         Width           =   3855
      End
      Begin RichTextLib.RichTextBox ctrlRichTextBoxLog 
         Height          =   6855
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   12091
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"main.frx":0082
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Configuración"
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   11295
      Begin VB.ComboBox cmbProtocolo 
         Height          =   315
         Left            =   4080
         TabIndex        =   19
         Text            =   "Extendido"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cmbEquipo 
         Height          =   315
         Left            =   4080
         TabIndex        =   18
         Text            =   "Ceres - Vesta"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton btnDisconnect 
         BackColor       =   &H80000002&
         Caption         =   "Desconectar"
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cmbBaudRate 
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Text            =   "9600"
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cmbPort 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Text            =   "XXX.XXX.XXX.XXX (IP)"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnConnect 
         BackColor       =   &H80000002&
         Caption         =   "Conectar"
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Protocolo:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "v.001.2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10320
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "(EpsonFiscalDriver.dll)"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   8160
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "EPSON LATIN AMERICA"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   8040
         TabIndex        =   12
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15495
      Begin VB.CommandButton btnContactenos 
         BackColor       =   &H80000002&
         Caption         =   "Contáctenos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   10175
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton btnConsultas 
         BackColor       =   &H80000002&
         Caption         =   "Consultas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   7845
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton btnOperaciones 
         BackColor       =   &H80000002&
         Caption         =   "Operaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   5520
         MaskColor       =   &H80000002&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   360
         Top             =   360
         Width           =   2250
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9240
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Empresa"
            TextSave        =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Puerto"
            TextSave        =   "Puerto"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Baudios"
            TextSave        =   "Baudios"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Versión DLL"
            TextSave        =   "Versión DLL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Versión IF"
            TextSave        =   "Versión IF"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/******************************************************************************
'*   Autor: Business Support And Development Unit                              *
'*                                                                             *
'*   Este código es gratuito y permite integrar impresoras fiscales EPSON      *
'*   usando la librería de bajo nivel (EpsonFiscalDriver.dll).                 *
'*                                                                             *
'*   Para implementarla las funciones de la libreria (.dll) usted debe incluir *
'*   el modulo 'EpsonFiscalDriver' en su proyecto.                             *
'*                                                                             *
'*   Se ofrecen funciones desarrollados para una rapida integración las cuales *
'*   demuestran la forma correcta de usar las funciones de la librería.        *
'*                                                                             *
'*   Este software se entrega con fines didácticos y sin garantía alguna.      *
'*   EPSON  NO ASUME responsabilidad legal alguna.                             *
'*   El programador usa este software bajo su propio riesgo y responsabilidad. *
'******************************************************************************/



Private Function cleanAllBoton()
    btnOperaciones.BackColor = &H80000002
    btnConsultas.BackColor = &H80000002
    btnContactenos.BackColor = &H80000002
End Function

Private Sub HabilitarPaneles()
    SSTab1.Enabled = True
End Sub

Private Sub DeshabilitarPaneles()
    SSTab1.Enabled = False
End Sub

Private Sub btnClose_Click()
    Select Case cmbOpenVoucher.ListIndex
        Case 0  ' - Tique
            Call cerrarTique
        Case 1  ' - NOTA DE CREDITO
            Call cerrarTiqueNC
        Case 2  ' - DNF
            Call cerrarDNF
    End Select
End Sub

Private Sub cerrarTique()
    Dim retorno As Long
    Dim cmd As String

    cmd = TICKET_CLOSE & TICKET_CLOSE_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Tiquet cerrado.")

End Sub

Private Sub cerrarTiqueNC()
    Dim retorno As Long
    Dim cmd As String

    cmd = TICKET_NC_CLOSE & TICKET_NC_CLOSE_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Tiquet nota de credito cerrado.")

End Sub

Private Sub cerrarDNF()
    Dim retorno As Long
    Dim cmd As String

    cmd = DNF_CLOSE & DNF_CLOSE_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Tiquet DNF cerrado.")

End Sub



Private Sub btnConsultas_Click()
    Call cleanAllBoton
    btnConsultas.BackColor = vbWhite
    SSTab1.Tab = 1
End Sub

Private Sub btnContactenos_Click()
    Call cleanAllBoton
    btnContactenos.BackColor = vbWhite
    SSTab1.Tab = 2
End Sub

Private Sub btnDisconnect_Click()
    Dim retorno As Long

    retorno = FP.Desconectar
    If retorno = ERROR_NONE Then
        Call DeshabilitarPaneles
        Call ReiniciarBarraDeEstado
        
        btnDisconnect.Enabled = False
        btnConnect.Enabled = True
        
        Call MostrarMensaje(retorno, "Host desconectado de la impresora fiscal.")
    Else
        Call MostrarMensaje(retorno, "")
    End If
End Sub

Private Sub btnDllVersion_Click()
    Call ObtenerVersionDll
End Sub

Private Sub btnErrorDescription_Click()
    Dim respuesta As String

    respuesta = FP.ConsultarDescripcionDeError()
    Call MostrarMensaje(ERROR_NONE, "Descripción de error: " & respuesta)
End Sub

Private Sub btnFiscalStatus_Click()
    Dim retorno As Long

    retorno = FP.ConsultarEstadoFiscal()
    Call MostrarMensaje(ERROR_NONE, "Estado fiscal: " & Hex(retorno))
End Sub

Private Sub btnGetFPVersion_Click()
    Call GetFPVersion
End Sub

Private Sub btnItemUp_Click()
    Select Case cmbOpenVoucher.ListIndex
        Case 0  ' - Tique
            Call ItemTique
        Case 1  ' - NOTA DE CREDITO
            Call ItemTiqueNC
        Case 2  ' - DNF
            Call ItemDNF
        End Select
End Sub


Private Sub ItemTique()
    Dim retorno As Long
    Dim cmd As String

    cmd = TICKET_ITEM & TICKET_ITEM_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Item enviado")

End Sub

Private Sub ItemTiqueNC()
    Dim retorno As Long
    Dim cmd As String

    cmd = TICKET_NC_ITEM & TICKET_NC_ITEM_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Item nota de credito enviado")

End Sub

Private Sub ItemDNF()
    Dim retorno As Long
    Dim cmd As String

    cmd = DNF_ITEM & DNF_ITEM_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Item DNF enviado")

End Sub


Private Sub btnLastError_Click()
    Dim retorno As Long

    retorno = FP.ultimoError()
    Call MostrarMensaje(ERROR_NONE, "Ultimo error: " & retorno)
End Sub

Private Sub btnOpen_Click()
    Select Case cmbOpenVoucher.ListIndex
        Case 0  ' - Tique
            Call AbrirTique
        Case 1  ' - NOTA DE CREDITO
            Call AbrirTiqueNC
        Case 2  ' - DNF
            Call AbrirDNF
    End Select
End Sub

Private Sub AbrirTique()
    Dim retorno As Long
    Dim cmd As String


    cmd = TICKET_OPEN
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Tiquet abierto.")

End Sub

Private Sub AbrirTiqueNC()
    Dim retorno As Long
    Dim cmd As String

    cmd = TICKET_NC_OPEN & TICKET_NC_OPEN_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Tiquet nota de credito abierto.")

End Sub

Private Sub AbrirDNF()
    Dim retorno As Long
    Dim cmd As String

    cmd = DNF_OPEN
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Tiquet DNF abierto.")
End Sub


Private Sub btnOperaciones_Click()
    Call cleanAllBoton
    btnOperaciones.BackColor = vbWhite
    SSTab1.Tab = 0
End Sub

Private Sub btnPayment_Click()
    Select Case cmbOpenVoucher.ListIndex
        Case 0  ' - Tique
            Call pagoTique
        Case 1  ' - NOTA DE CREDITO
            Call pagoTiqueNC
    End Select
End Sub

Private Sub pagoTique()
    Dim retorno As Long
    Dim cmd As String

    cmd = TICKET_NC_PAYMENT & TICKET_NC_PAYMENT_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Pago tiquet enviado")

End Sub

Private Sub pagoTiqueNC()
    Dim retorno As Long
    Dim cmd As String

    cmd = TICKET_NC_PAYMENT & TICKET_NC_PAYMENT_FIELDS
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Pago tiquet nota de credito enviado")

End Sub

Private Sub btnPrinterStatus_Click()
    Dim retorno As Long

    retorno = FP.ConsultarEstadoImpresora()
    Call MostrarMensaje(ERROR_NONE, "Estado de impresora: " & Hex(retorno))
End Sub

Private Sub btnX_Click()
    Dim retorno As Long
    Dim cmd As String

    cmd = X_REPORT
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Cambio de cajero.")

End Sub

Private Sub btnZ_Click()
    Dim retorno As Long
    Dim cmd As String

    cmd = Z_REPORT
    retorno = FP.EnviarComando(cmd)
    Call MostrarMensaje(retorno, "Jornada cerrada.")
End Sub




Private Sub Form_Load()
    Call setCommands_init
    
    Call btnOperaciones_Click
    RichTextBox1.Text = "Este software es un ejemplo de uso de las funciones expuestas por la librería de bajo nivel 'EpsonFiscalDriver'." & vbCrLf & vbCrLf & "Utilizando el Manual de Especificaciones de la impresora, se arma la trama para ser enviada. Ver ejemplos en este demo." & vbCrLf & vbCrLf & "En caso que necesite contactarnos puede hacerlo al correo: soporte_sd_argentina@epson.com.ar."
    
    cmbPort.ListIndex = 0
    cmbBaudRate.ListIndex = 0
    btnDisconnect.Enabled = False
    
    cmbProtocolo.ListIndex = 1
    cmbEquipo.ListIndex = 1
    cmbOpenVoucher.ListIndex = 0
    
    
    Call DeshabilitarPaneles
    Call ReiniciarBarraDeEstado
    
End Sub


Private Sub ReiniciarBarraDeEstado()
    StatusBar1.Panels.item(1).Text = "Epson Latin America"
    StatusBar1.Panels.item(2).Text = "     ---     "
    StatusBar1.Panels.item(3).Text = "     ---     "
    StatusBar1.Panels.item(4).Text = "     ---     "
    StatusBar1.Panels.item(5).Text = "     ---     "
End Sub


 Private Sub MostrarMensaje(ByVal retorno, ByVal mensaje)
    Dim respuesta As String

    If retorno = ERROR_NONE Then
        Call MsgLogAgregar(mensaje, True, "Información")
    Else
        respuesta = FP.ConsultarDescripcionDeError()
        If (respuesta = "") Then
            Call MsgLogAgregar("Error: " & retorno & vbCrLf & "Descripción: " & "No definido en el listado de errores", False, "Error!")
        Else
            Call MsgLogAgregar("Error: " & retorno & vbCrLf & "Descripción: " & respuesta, False, "Error!")
        End If
    End If

End Sub



Private Sub MsgLogAgregar(ByVal mensaje As String, ByVal tipo As Boolean, ByVal titulo As String)
    Dim default_color As ColorConstants
    Dim information_color As ColorConstants
    Dim error_color As ColorConstants

    ' init
    default_color = ctrlRichTextBoxLog.SelColor
    information_color = vbBlue
    error_color = vbRed


    ' ir al final del texto
    ctrlRichTextBoxLog.SelStart = ctrlRichTextBoxLog.SelLength


    ' titulo del mensaje
    If tipo Then 'MsgBoxStyle.Information
        ctrlRichTextBoxLog.SelColor = information_color
    Else
        ctrlRichTextBoxLog.SelColor = error_color
    End If
    
    ctrlRichTextBoxLog.Text = ctrlRichTextBoxLog.Text & titulo & vbCrLf

    ' cuerpo del mensaje
    ctrlRichTextBoxLog.SelColor = default_color
    ' ir al final del texto
    
    ctrlRichTextBoxLog.Text = ctrlRichTextBoxLog.Text & vbCrLf & mensaje & vbCrLf & vbCrLf
    
    ctrlRichTextBoxLog.SelStart = ctrlRichTextBoxLog.SelLength
    ctrlRichTextBoxLog.SetFocus

End Sub


Function ConectarFiscal() As Long
    Dim retorno As Long
    Dim Port As Integer
    Dim myPort As String
        

    Call SetData(cmbProtocolo.ListIndex, cmbEquipo.ListIndex) 'Opcional solo para efectos demostrativos en el ejemplo


    Do While (True)

        retorno = FP.ConfigurarVelocidad(Val(cmbBaudRate.Text))
        If Not (retorno = ERROR_NONE) Then
            Exit Do
        End If


        Port = cmbPort.ListIndex - 1

        If Port >= 0 Then
            myPort = Str(Port)
        Else
            myPort = cmbPort.Text
        End If
        
        retorno = FP.ConfigurarPuerto(myPort)
        If Not (retorno = ERROR_NONE) Then
            Exit Do
        End If



        retorno = FP.ConfigurarProtocolo(cmbProtocolo.ListIndex)
        If Not (retorno = ERROR_NONE) Then
            Exit Do
        End If


        retorno = FP.NewConectar()
        If Not (retorno = ERROR_NONE) Then
            Exit Do
        End If

        
        Select Case cmbPort.ListIndex
            Case 0
                StatusBar1.Panels.item(2).Text = "     USB     "
            Case 1
                StatusBar1.Panels.item(2).Text = "     COM1     "
            Case 2
                StatusBar1.Panels.item(2).Text = "     COM2     "
            Case 3
                StatusBar1.Panels.item(2).Text = "     COM3     "
            Case 4
                StatusBar1.Panels.item(2).Text = "     COM4     "
            Case 5
                StatusBar1.Panels.item(2).Text = "     COM5     "
            Case 6
                StatusBar1.Panels.item(2).Text = "     COM6     "
            Case 7
                StatusBar1.Panels.item(2).Text = "     COM7     "
            Case 8
                StatusBar1.Panels.item(2).Text = "     COM8     "
            Case 9
                StatusBar1.Panels.item(2).Text = "     COM9     "
            Case 10
                StatusBar1.Panels.item(2).Text = "     COM10     "
        End Select

        If cmbPort.ListIndex > 0 Then
            Select Case cmbBaudRate.ListIndex
                Case 0
                    StatusBar1.Panels.item(3).Text = "     9600     "
                Case 1
                    StatusBar1.Panels.item(3).Text = "     19200     "
                Case 2
                    StatusBar1.Panels.item(3).Text = "     38400     "
                Case 3
                    StatusBar1.Panels.item(3).Text = "     57600     "
                Case 4
                    StatusBar1.Panels.item(3).Text = "     115200     "
            End Select
        End If


        Exit Do
    Loop

    Call MostrarMensaje(retorno, "Host vinculado a la impresora fiscal.")

    ConectarFiscal = retorno
End Function

Private Sub ObtenerVersionDll()
    Dim respuesta As String
 
    respuesta = FP.ConsultarVersionDll()
    
    Call MostrarMensaje(retorno, "Versión de la dll --> " & respuesta)
    
    ' mostrar en toolbar
    respuesta = Mid(respuesta, 41, 5)
    StatusBar1.Panels.item(4).Text = "Versión Dll: v.: " + respuesta
End Sub
    

Private Function GetFPVersion() As Long
    Dim cad As String

    cad = FP.ConsultarVersionIF(GET_FIRMWARE_VERSION, NUM_CAMPO_VERSION, NUM_CAMPO_VERSION_MAYOR, NUM_CAMPO_VERSION_MENOR)
    If (cad = "") Then
        Call MostrarMensaje(1, "") ' el valor 1 es para indicar cualquier error ya que ne la funcion mostrar mensaje se detecta el error
        GetFPVersion = 1
    Else
        Call MostrarMensaje(ERROR_NONE, "Versión de la impresora fiscal --> " & cad)
        StatusBar1.Panels.item(5).Text = cad
        GetFPVersion = ERROR_NONE
    End If
End Function


Private Sub btnConnect_Click()
    Dim retorno As Long

    Do While (True)
        retorno = ConectarFiscal()
        If Not (retorno = ERROR_NONE) Then
            Exit Do
        End If

        Call ObtenerVersionDll

        retorno = GetFPVersion()
        If Not (retorno = ERROR_NONE) Then
            Call btnDisconnect_Click
            Exit Do
        End If

        btnDisconnect.Enabled = True
        btnConnect.Enabled = False
        Call HabilitarPaneles
        Exit Do
    Loop
        
End Sub





