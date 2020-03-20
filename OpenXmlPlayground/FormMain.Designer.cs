namespace OpenXmlPlayground
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
            this.btnSimpleWordTest = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSimpleWordTest
            // 
            this.btnSimpleWordTest.Location = new System.Drawing.Point(12, 12);
            this.btnSimpleWordTest.Name = "btnSimpleWordTest";
            this.btnSimpleWordTest.Size = new System.Drawing.Size(460, 23);
            this.btnSimpleWordTest.TabIndex = 0;
            this.btnSimpleWordTest.Text = "SIMPLE WORD DOCUMENT TEST";
            this.btnSimpleWordTest.UseVisualStyleBackColor = true;
            this.btnSimpleWordTest.Click += new System.EventHandler(this.btnSimpleWordTest_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 361);
            this.Controls.Add(this.btnSimpleWordTest);
            this.Name = "FormMain";
            this.Text = "OpenXML Playground";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnSimpleWordTest;
    }
}

