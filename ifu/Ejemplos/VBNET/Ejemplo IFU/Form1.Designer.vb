Partial Class Form1
	''' <summary>
	''' Required designer variable.
	''' </summary>
	Private components As System.ComponentModel.IContainer = Nothing

	''' <summary>
	''' Clean up any resources being used.
	''' </summary>
	''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
	Protected Overrides Sub Dispose(disposing As Boolean)
		If disposing AndAlso (components IsNot Nothing) Then
			components.Dispose()
		End If
		MyBase.Dispose(disposing)
	End Sub

	#Region "Windows Form Designer generated code"

	''' <summary>
	''' Required method for Designer support - do not modify
	''' the contents of this method with the code editor.
	''' </summary>
	Private Sub InitializeComponent()
		Me.Button3 = New System.Windows.Forms.Button()
		Me.button1 = New System.Windows.Forms.Button()
		Me.button2 = New System.Windows.Forms.Button()
		Me.button4 = New System.Windows.Forms.Button()
		Me.button5 = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		' 
		' Button3
		' 
		Me.Button3.Location = New System.Drawing.Point(17, 33)
		Me.Button3.Name = "Button3"
		Me.Button3.Size = New System.Drawing.Size(82, 67)
		Me.Button3.TabIndex = 3
		Me.Button3.Text = "Factura A"
		Me.Button3.UseVisualStyleBackColor = True
		AddHandler Me.Button3.Click, New System.EventHandler(AddressOf Me.Button3_Click)
		' 
		' button1
		' 
		Me.button1.Location = New System.Drawing.Point(135, 33)
		Me.button1.Name = "button1"
		Me.button1.Size = New System.Drawing.Size(82, 67)
		Me.button1.TabIndex = 4
		Me.button1.Text = "Factura B"
		Me.button1.UseVisualStyleBackColor = True
		AddHandler Me.button1.Click, New System.EventHandler(AddressOf Me.button1_Click)
		' 
		' button2
		' 
		Me.button2.Location = New System.Drawing.Point(261, 33)
		Me.button2.Name = "button2"
		Me.button2.Size = New System.Drawing.Size(82, 67)
		Me.button2.TabIndex = 5
		Me.button2.Text = "Ticket"
		Me.button2.UseVisualStyleBackColor = True
		AddHandler Me.button2.Click, New System.EventHandler(AddressOf Me.button2_Click)
		' 
		' button4
		' 
		Me.button4.Location = New System.Drawing.Point(73, 121)
		Me.button4.Name = "button4"
		Me.button4.Size = New System.Drawing.Size(82, 67)
		Me.button4.TabIndex = 6
		Me.button4.Text = "Cierre Z"
		Me.button4.UseVisualStyleBackColor = True
		AddHandler Me.button4.Click, New System.EventHandler(AddressOf Me.button4_Click)
		' 
		' button5
		' 
		Me.button5.Location = New System.Drawing.Point(201, 121)
		Me.button5.Name = "button5"
		Me.button5.Size = New System.Drawing.Size(82, 67)
		Me.button5.TabIndex = 7
		Me.button5.Text = "Cierre X"
		Me.button5.UseVisualStyleBackColor = True
		AddHandler Me.button5.Click, New System.EventHandler(AddressOf Me.button5_Click)
		' 
		' Form1
		' 
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(362, 233)
		Me.Controls.Add(Me.button5)
		Me.Controls.Add(Me.button4)
		Me.Controls.Add(Me.button2)
		Me.Controls.Add(Me.button1)
		Me.Controls.Add(Me.Button3)
		Me.Name = "Form1"
		Me.Text = "Form1"
		Me.ResumeLayout(False)

	End Sub

	#End Region

	Friend Button3 As System.Windows.Forms.Button
	Friend button1 As System.Windows.Forms.Button
	Friend button2 As System.Windows.Forms.Button
	Friend button4 As System.Windows.Forms.Button
	Friend button5 As System.Windows.Forms.Button
End Class

