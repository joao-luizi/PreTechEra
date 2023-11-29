namespace SignConf
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btFolder = new System.Windows.Forms.Button();
            this.btSig = new System.Windows.Forms.Button();
            this.btDoc = new System.Windows.Forms.Button();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.txtSig = new System.Windows.Forms.TextBox();
            this.txtDoc = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.panel5 = new System.Windows.Forms.Panel();
            this.button3 = new System.Windows.Forms.Button();
            this.panel6 = new System.Windows.Forms.Panel();
            this.nmSigRedux = new System.Windows.Forms.NumericUpDown();
            this.nmPosY = new System.Windows.Forms.NumericUpDown();
            this.nmPosX = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.bckDocumentLoader = new System.ComponentModel.BackgroundWorker();
            this.bckSigLoader = new System.ComponentModel.BackgroundWorker();
            this.label7 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nmSigRedux)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmPosY)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmPosX)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.Controls.Add(this.panel4, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel3, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel5, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel6, 2, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 148F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 41F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(784, 562);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // panel4
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.panel4, 3);
            this.panel4.Controls.Add(this.pictureBox1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(3, 192);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(778, 367);
            this.panel4.TabIndex = 3;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(778, 367);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // panel1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.panel1, 2);
            this.panel1.Controls.Add(this.button4);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.btFolder);
            this.panel1.Controls.Add(this.btSig);
            this.panel1.Controls.Add(this.btDoc);
            this.panel1.Controls.Add(this.txtFolder);
            this.panel1.Controls.Add(this.txtSig);
            this.panel1.Controls.Add(this.txtDoc);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(516, 142);
            this.panel1.TabIndex = 0;
            // 
            // btFolder
            // 
            this.btFolder.Location = new System.Drawing.Point(470, 112);
            this.btFolder.Name = "btFolder";
            this.btFolder.Size = new System.Drawing.Size(29, 20);
            this.btFolder.TabIndex = 8;
            this.btFolder.Text = "...";
            this.btFolder.UseVisualStyleBackColor = true;
            this.btFolder.Click += new System.EventHandler(this.btFolder_Click);
            // 
            // btSig
            // 
            this.btSig.Location = new System.Drawing.Point(470, 36);
            this.btSig.Name = "btSig";
            this.btSig.Size = new System.Drawing.Size(29, 20);
            this.btSig.TabIndex = 7;
            this.btSig.Text = "...";
            this.btSig.UseVisualStyleBackColor = true;
            this.btSig.Click += new System.EventHandler(this.btSig_Click);
            // 
            // btDoc
            // 
            this.btDoc.Location = new System.Drawing.Point(470, 7);
            this.btDoc.Name = "btDoc";
            this.btDoc.Size = new System.Drawing.Size(29, 20);
            this.btDoc.TabIndex = 6;
            this.btDoc.Text = "...";
            this.btDoc.UseVisualStyleBackColor = true;
            this.btDoc.Click += new System.EventHandler(this.btDoc_Click);
            // 
            // txtFolder
            // 
            this.txtFolder.Location = new System.Drawing.Point(80, 112);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(384, 20);
            this.txtFolder.TabIndex = 5;
            // 
            // txtSig
            // 
            this.txtSig.Location = new System.Drawing.Point(80, 36);
            this.txtSig.Name = "txtSig";
            this.txtSig.Size = new System.Drawing.Size(384, 20);
            this.txtSig.TabIndex = 4;
            // 
            // txtDoc
            // 
            this.txtDoc.Location = new System.Drawing.Point(80, 7);
            this.txtDoc.Name = "txtDoc";
            this.txtDoc.Size = new System.Drawing.Size(384, 20);
            this.txtDoc.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(-2, 116);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Pasta de Saida";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Assinatura";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Documento";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.button1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 151);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(255, 35);
            this.panel2.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button1.Location = new System.Drawing.Point(0, 0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(255, 35);
            this.button1.TabIndex = 0;
            this.button1.Text = "Carregar Documentos";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.button2);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(264, 151);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(255, 35);
            this.panel3.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button2.Location = new System.Drawing.Point(0, 0);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(255, 35);
            this.button2.TabIndex = 1;
            this.button2.Text = "Refrescar";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.button3);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(525, 151);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(256, 35);
            this.panel5.TabIndex = 4;
            // 
            // button3
            // 
            this.button3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button3.Location = new System.Drawing.Point(0, 0);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(256, 35);
            this.button3.TabIndex = 2;
            this.button3.Text = "Dar Saida";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.nmSigRedux);
            this.panel6.Controls.Add(this.nmPosY);
            this.panel6.Controls.Add(this.nmPosX);
            this.panel6.Controls.Add(this.label6);
            this.panel6.Controls.Add(this.label5);
            this.panel6.Controls.Add(this.label4);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel6.Location = new System.Drawing.Point(525, 3);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(256, 142);
            this.panel6.TabIndex = 5;
            // 
            // nmSigRedux
            // 
            this.nmSigRedux.Location = new System.Drawing.Point(79, 71);
            this.nmSigRedux.Name = "nmSigRedux";
            this.nmSigRedux.Size = new System.Drawing.Size(77, 20);
            this.nmSigRedux.TabIndex = 5;
            // 
            // nmPosY
            // 
            this.nmPosY.Location = new System.Drawing.Point(79, 39);
            this.nmPosY.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.nmPosY.Name = "nmPosY";
            this.nmPosY.Size = new System.Drawing.Size(77, 20);
            this.nmPosY.TabIndex = 4;
            // 
            // nmPosX
            // 
            this.nmPosX.Location = new System.Drawing.Point(79, 7);
            this.nmPosX.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.nmPosX.Name = "nmPosX";
            this.nmPosX.Size = new System.Drawing.Size(77, 20);
            this.nmPosX.TabIndex = 3;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(14, 11);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(55, 13);
            this.label6.TabIndex = 2;
            this.label6.Text = "Posição X";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 43);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(55, 13);
            this.label5.TabIndex = 1;
            this.label5.Text = "Posição Y";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(1, 75);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(68, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Redução (%)";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(15, 71);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(118, 13);
            this.label7.TabIndex = 9;
            this.label7.Text = "Documentos Adicionais";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(139, 66);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(325, 22);
            this.button4.TabIndex = 10;
            this.button4.Text = "...";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 562);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "Form1";
            this.Text = "Conferido Sign";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nmSigRedux)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmPosY)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nmPosX)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.TextBox txtSig;
        private System.Windows.Forms.TextBox txtDoc;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.NumericUpDown nmSigRedux;
        private System.Windows.Forms.NumericUpDown nmPosY;
        private System.Windows.Forms.NumericUpDown nmPosX;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btFolder;
        private System.Windows.Forms.Button btSig;
        private System.Windows.Forms.Button btDoc;
        private System.ComponentModel.BackgroundWorker bckDocumentLoader;
        private System.ComponentModel.BackgroundWorker bckSigLoader;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label7;
    }
}

