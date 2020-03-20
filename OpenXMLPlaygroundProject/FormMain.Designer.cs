namespace OpenXMLPlaygroundProject
{
    partial class FormMain
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSimpleWordText = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSimpleWordText
            // 
            this.btnSimpleWordText.Location = new System.Drawing.Point(0, 0);
            this.btnSimpleWordText.Name = "btnSimpleWordText";
            this.btnSimpleWordText.Size = new System.Drawing.Size(488, 23);
            this.btnSimpleWordText.TabIndex = 0;
            this.btnSimpleWordText.Text = "SIMPLE WORD DOCUMENT TEST";
            this.btnSimpleWordText.UseVisualStyleBackColor = true;
            this.btnSimpleWordText.Click += new System.EventHandler(this.btnSimpleWordText_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 361);
            this.Controls.Add(this.btnSimpleWordText);
            this.Name = "FormMain";
            this.ShowIcon = false;
            this.Text = "OpenXML Playground";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnSimpleWordText;
    }
}

