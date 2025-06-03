namespace ExcelPlanFinanciar
{
    partial class FormGraficSelector
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
            this.radioLinie = new System.Windows.Forms.RadioButton();
            this.radioStructura = new System.Windows.Forms.RadioButton();
            this.radioCombo = new System.Windows.Forms.RadioButton();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // radioLinie
            // 
            this.radioLinie.AutoSize = true;
            this.radioLinie.Location = new System.Drawing.Point(30, 20);
            this.radioLinie.Name = "radioLinie";
            this.radioLinie.Size = new System.Drawing.Size(186, 19);
            this.radioLinie.TabIndex = 0;
            this.radioLinie.TabStop = true;
            this.radioLinie.Text = "Evoluţie Sold + Dobândă (linie)";
            this.radioLinie.UseVisualStyleBackColor = true;
            // 
            // radioStructura
            // 
            this.radioStructura.AutoSize = true;
            this.radioStructura.Location = new System.Drawing.Point(30, 50);
            this.radioStructura.Name = "radioStructura";
            this.radioStructura.Size = new System.Drawing.Size(210, 19);
            this.radioStructura.TabIndex = 1;
            this.radioStructura.TabStop = true;
            this.radioStructura.Text = "Structură lunară (% principal/dobândă)";
            this.radioStructura.UseVisualStyleBackColor = true;
            // 
            // radioCombo
            // 
            this.radioCombo.AutoSize = true;
            this.radioCombo.Location = new System.Drawing.Point(30, 80);
            this.radioCombo.Name = "radioCombo";
            this.radioCombo.Size = new System.Drawing.Size(190, 19);
            this.radioCombo.TabIndex = 2;
            this.radioCombo.TabStop = true;
            this.radioCombo.Text = "Dobândă + Sold (grafic combinat)";
            this.radioCombo.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(90, 120);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(90, 27);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "Generează";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // FormGraficSelector
            // 
            this.ClientSize = new System.Drawing.Size(280, 170);
            this.Controls.Add(this.radioLinie);
            this.Controls.Add(this.radioStructura);
            this.Controls.Add(this.radioCombo);
            this.Controls.Add(this.btnOK);
            this.Name = "FormGraficSelector";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Alege tipul de grafic";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.RadioButton radioLinie;
        private System.Windows.Forms.RadioButton radioStructura;
        private System.Windows.Forms.RadioButton radioCombo;
        private System.Windows.Forms.Button btnOK;
    }
}