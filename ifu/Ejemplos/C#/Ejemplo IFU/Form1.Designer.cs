﻿namespace Ejemplo_IFU
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Button3
            // 
            this.Button3.Location = new System.Drawing.Point(17, 33);
            this.Button3.Name = "Button3";
            this.Button3.Size = new System.Drawing.Size(82, 67);
            this.Button3.TabIndex = 3;
            this.Button3.Text = "Factura A";
            this.Button3.UseVisualStyleBackColor = true;
            this.Button3.Click += new System.EventHandler(this.Button3_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(135, 33);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(82, 67);
            this.button1.TabIndex = 4;
            this.button1.Text = "Factura B";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(261, 33);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(82, 67);
            this.button2.TabIndex = 5;
            this.button2.Text = "Ticket";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(73, 121);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(82, 67);
            this.button4.TabIndex = 6;
            this.button4.Text = "Cierre Z";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(201, 121);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(82, 67);
            this.button5.TabIndex = 7;
            this.button5.Text = "Cierre X";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(362, 233);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.Button3);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Button Button3;
        internal System.Windows.Forms.Button button1;
        internal System.Windows.Forms.Button button2;
        internal System.Windows.Forms.Button button4;
        internal System.Windows.Forms.Button button5;
    }
}

