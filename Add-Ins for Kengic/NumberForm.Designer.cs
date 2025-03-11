
namespace Kengic
{
    partial class NumberForm
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
            this.numericUpDownStarting = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownIncrement = new System.Windows.Forms.NumericUpDown();
            this.labelStartingNumber = new System.Windows.Forms.Label();
            this.labelIncrement = new System.Windows.Forms.Label();
            this.buttonOk = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStarting)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownIncrement)).BeginInit();
            this.SuspendLayout();
            // 
            // numericUpDownStarting
            // 
            this.numericUpDownStarting.BackColor = System.Drawing.Color.White;
            this.numericUpDownStarting.Location = new System.Drawing.Point(108, 12);
            this.numericUpDownStarting.Maximum = new decimal(new int[] {
            59999,
            0,
            0,
            0});
            this.numericUpDownStarting.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownStarting.Name = "numericUpDownStarting";
            this.numericUpDownStarting.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.numericUpDownStarting.Size = new System.Drawing.Size(120, 20);
            this.numericUpDownStarting.TabIndex = 0;
            this.numericUpDownStarting.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownStarting.Enter += new System.EventHandler(this.numericUpDownStarting_Enter);
            // 
            // numericUpDownIncrement
            // 
            this.numericUpDownIncrement.Location = new System.Drawing.Point(108, 38);
            this.numericUpDownIncrement.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownIncrement.Name = "numericUpDownIncrement";
            this.numericUpDownIncrement.Size = new System.Drawing.Size(120, 20);
            this.numericUpDownIncrement.TabIndex = 1;
            this.numericUpDownIncrement.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownIncrement.Enter += new System.EventHandler(this.numericUpDownIncrement_Enter);
            // 
            // labelStartingNumber
            // 
            this.labelStartingNumber.AutoSize = true;
            this.labelStartingNumber.Location = new System.Drawing.Point(12, 14);
            this.labelStartingNumber.Name = "labelStartingNumber";
            this.labelStartingNumber.Size = new System.Drawing.Size(86, 13);
            this.labelStartingNumber.TabIndex = 2;
            this.labelStartingNumber.Text = "起始编号:";
            // 
            // labelIncrement
            // 
            this.labelIncrement.AutoSize = true;
            this.labelIncrement.Location = new System.Drawing.Point(12, 40);
            this.labelIncrement.Name = "labelIncrement";
            this.labelIncrement.Size = new System.Drawing.Size(57, 13);
            this.labelIncrement.TabIndex = 3;
            this.labelIncrement.Text = "递增值:";
            // 
            // buttonOk
            // 
            this.buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOk.Location = new System.Drawing.Point(56, 75);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(75, 23);
            this.buttonOk.TabIndex = 4;
            this.buttonOk.Text = "好";
            this.buttonOk.UseVisualStyleBackColor = true;
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(153, 75);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 5;
            this.buttonCancel.Text = "取消";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // NumberForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(240, 110);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.labelIncrement);
            this.Controls.Add(this.labelStartingNumber);
            this.Controls.Add(this.numericUpDownIncrement);
            this.Controls.Add(this.numericUpDownStarting);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "NumberForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "程序块自动编号";
            this.TopMost = true;
            this.Deactivate += new System.EventHandler(this.NumberForm_Deactivate);
            this.Load += new System.EventHandler(this.NumberForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStarting)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownIncrement)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.NumericUpDown numericUpDownStarting;
        private System.Windows.Forms.NumericUpDown numericUpDownIncrement;
        private System.Windows.Forms.Label labelStartingNumber;
        private System.Windows.Forms.Label labelIncrement;
        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.Button buttonCancel;
    }
}