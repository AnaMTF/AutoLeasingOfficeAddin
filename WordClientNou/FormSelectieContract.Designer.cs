namespace WordClientNou
{
    partial class FormSelectieContract
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
            this.labelTitlu = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBoxClienti = new System.Windows.Forms.ComboBox();
            this.buttonConfirma = new System.Windows.Forms.Button();
            this.buttonRenunta = new System.Windows.Forms.Button();
            this.comboBoxOferta = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // labelTitlu
            // 
            this.labelTitlu.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.labelTitlu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelTitlu.Dock = System.Windows.Forms.DockStyle.Top;
            this.labelTitlu.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.labelTitlu.Font = new System.Drawing.Font("Monotype Corsiva", 26.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTitlu.Location = new System.Drawing.Point(0, 0);
            this.labelTitlu.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelTitlu.Name = "labelTitlu";
            this.labelTitlu.Size = new System.Drawing.Size(888, 69);
            this.labelTitlu.TabIndex = 0;
            this.labelTitlu.Text = "Generare Contract";
            this.labelTitlu.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.label1.Font = new System.Drawing.Font("MV Boli", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(36, 145);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(826, 39);
            this.label1.TabIndex = 1;
            this.label1.Text = "Va rugam sa selectati clientul si sa confirmati oferta dorita";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.label2.Font = new System.Drawing.Font("MV Boli", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(88, 263);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(129, 48);
            this.label2.TabIndex = 2;
            this.label2.Text = "Clientul:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.label3.Font = new System.Drawing.Font("MV Boli", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(88, 380);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(129, 48);
            this.label3.TabIndex = 3;
            this.label3.Text = "Oferta:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBoxClienti
            // 
            this.comboBoxClienti.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.comboBoxClienti.FormattingEnabled = true;
            this.comboBoxClienti.Location = new System.Drawing.Point(302, 263);
            this.comboBoxClienti.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.comboBoxClienti.Name = "comboBoxClienti";
            this.comboBoxClienti.Size = new System.Drawing.Size(442, 28);
            this.comboBoxClienti.TabIndex = 4;
            // 
            // buttonConfirma
            // 
            this.buttonConfirma.BackColor = System.Drawing.Color.Thistle;
            this.buttonConfirma.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonConfirma.Location = new System.Drawing.Point(190, 458);
            this.buttonConfirma.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonConfirma.Name = "buttonConfirma";
            this.buttonConfirma.Size = new System.Drawing.Size(130, 55);
            this.buttonConfirma.TabIndex = 5;
            this.buttonConfirma.Text = "Confirma";
            this.buttonConfirma.UseVisualStyleBackColor = false;
            this.buttonConfirma.Click += new System.EventHandler(this.buttonConfirma_Click);
            // 
            // buttonRenunta
            // 
            this.buttonRenunta.BackColor = System.Drawing.Color.Thistle;
            this.buttonRenunta.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonRenunta.Location = new System.Drawing.Point(520, 458);
            this.buttonRenunta.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonRenunta.Name = "buttonRenunta";
            this.buttonRenunta.Size = new System.Drawing.Size(130, 55);
            this.buttonRenunta.TabIndex = 6;
            this.buttonRenunta.Text = "Renunta";
            this.buttonRenunta.UseVisualStyleBackColor = false;
            this.buttonRenunta.Click += new System.EventHandler(this.buttonRenunta_Click);
            // 
            // comboBoxOferta
            // 
            this.comboBoxOferta.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.comboBoxOferta.FormattingEnabled = true;
            this.comboBoxOferta.Location = new System.Drawing.Point(302, 380);
            this.comboBoxOferta.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.comboBoxOferta.Name = "comboBoxOferta";
            this.comboBoxOferta.Size = new System.Drawing.Size(442, 28);
            this.comboBoxOferta.TabIndex = 7;
            // 
            // FormSelectieContract
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.BackgroundImage = global::WordClientNou.Properties.Resources.images__1_;
            this.ClientSize = new System.Drawing.Size(888, 523);
            this.Controls.Add(this.comboBoxOferta);
            this.Controls.Add(this.buttonRenunta);
            this.Controls.Add(this.buttonConfirma);
            this.Controls.Add(this.comboBoxClienti);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.labelTitlu);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormSelectieContract";
            this.Text = "FormSelectieContract";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelTitlu;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBoxClienti;
        private System.Windows.Forms.Button buttonConfirma;
        private System.Windows.Forms.Button buttonRenunta;
        private System.Windows.Forms.ComboBox comboBoxOferta;
    }
}