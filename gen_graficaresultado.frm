VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form gen_graficaresultado 
   Caption         =   "Graficas"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   7560
      Width           =   2415
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   7335
      Left            =   720
      OleObjectBlob   =   "gen_graficaresultado.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "gen_graficaresultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
J = MsgBox("Prepare Impresora y Confirme", 4)
If J = 6 Then
         escala = 1
        'limpia portapapaetkles
        Clipboard.clear
        'copia grafico a portapapeles
        MSChart1.EditCopy
           
        ' sila imagen es válida
        If Clipboard.GetFormat(vbCFBitmap) Then
            'scale mode
             Printer.ScaleMode = vbTwips
            MSChart1.Parent.ScaleMode = vbTwips
               
           ' titulo
           ' Printer.Font.Size = 10
            'Printer.FontName = "Verdana"
               
            'Printer.Print vbNullString
            'Printer.Print titulo
            'Printer.Print vbNullString
               
            ' dibuja la imagen
            Printer.PaintPicture Clipboard.GetData(), 100, 500, _
                                 MSChart1.Width * escala, MSChart1.Height * escala, 0, 0
           
               
            Printer.EndDoc ' envía el trabajo a la impresora
        End If
   
End If

End Sub
