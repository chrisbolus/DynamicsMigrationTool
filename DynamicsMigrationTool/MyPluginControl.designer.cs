namespace DynamicsMigrationTool
{
    partial class MyPluginControl
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.CreateStgTbl_Btn = new System.Windows.Forms.Button();
            this.EntityCmb = new System.Windows.Forms.ComboBox();
            this.stagingDBConnection_txtb = new System.Windows.Forms.TextBox();
            this.EntityName_Lbl = new System.Windows.Forms.Label();
            this.StagingDatabase_Lbl = new System.Windows.Forms.Label();
            this.About_btn = new System.Windows.Forms.Button();
            this.CreateSrcVwTmpl_Btn = new System.Windows.Forms.Button();
            this.sourceDBConnection_txtb = new System.Windows.Forms.TextBox();
            this.SourceDatabase_Lbl = new System.Windows.Forms.Label();
            this.test_btn = new System.Windows.Forms.Button();
            this.sourceToStagingLocation_txtb = new System.Windows.Forms.TextBox();
            this.sourceToStagingLocation_Lbl = new System.Windows.Forms.Label();
            this.CreateS2SPackage_Btn = new System.Windows.Forms.Button();
            this.sourceDBSchema_txtb = new System.Windows.Forms.TextBox();
            this.stagingDBSchema_txtb = new System.Windows.Forms.TextBox();
            this.sourceDBSchema_Lbl = new System.Windows.Forms.Label();
            this.stagingDBSchema_Lbl = new System.Windows.Forms.Label();
            this.dataTransforms_chbx = new System.Windows.Forms.CheckBox();
            this.crmToStagingLocation_txtb = new System.Windows.Forms.TextBox();
            this.crmToStagingLocation_lbl = new System.Windows.Forms.Label();
            this.stagingDBImportSchema_txtb = new System.Windows.Forms.TextBox();
            this.stagingDBImportSchema_lbl = new System.Windows.Forms.Label();
            this.createCRMToStagingPackage_btn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CreateStgTbl_Btn
            // 
            this.CreateStgTbl_Btn.Location = new System.Drawing.Point(41, 152);
            this.CreateStgTbl_Btn.Name = "CreateStgTbl_Btn";
            this.CreateStgTbl_Btn.Size = new System.Drawing.Size(247, 32);
            this.CreateStgTbl_Btn.TabIndex = 30;
            this.CreateStgTbl_Btn.Text = "Create Staging Table";
            this.CreateStgTbl_Btn.UseVisualStyleBackColor = true;
            this.CreateStgTbl_Btn.Click += new System.EventHandler(this.CreateStgTbl_Btn_Click);
            // 
            // EntityCmb
            // 
            this.EntityCmb.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.EntityCmb.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.EntityCmb.FormattingEnabled = true;
            this.EntityCmb.Location = new System.Drawing.Point(42, 70);
            this.EntityCmb.Name = "EntityCmb";
            this.EntityCmb.Size = new System.Drawing.Size(265, 24);
            this.EntityCmb.TabIndex = 10;
            // 
            // stagingDBConnection_txtb
            // 
            this.stagingDBConnection_txtb.Location = new System.Drawing.Point(42, 332);
            this.stagingDBConnection_txtb.Name = "stagingDBConnection_txtb";
            this.stagingDBConnection_txtb.Size = new System.Drawing.Size(505, 22);
            this.stagingDBConnection_txtb.TabIndex = 60;
            this.stagingDBConnection_txtb.TextChanged += new System.EventHandler(this.stagingDBConnection_txtb_TextChanged);
            // 
            // EntityName_Lbl
            // 
            this.EntityName_Lbl.AutoSize = true;
            this.EntityName_Lbl.Location = new System.Drawing.Point(42, 48);
            this.EntityName_Lbl.Name = "EntityName_Lbl";
            this.EntityName_Lbl.Size = new System.Drawing.Size(39, 16);
            this.EntityName_Lbl.TabIndex = 11;
            this.EntityName_Lbl.Text = "Entity";
            // 
            // StagingDatabase_Lbl
            // 
            this.StagingDatabase_Lbl.AutoSize = true;
            this.StagingDatabase_Lbl.Location = new System.Drawing.Point(42, 313);
            this.StagingDatabase_Lbl.Name = "StagingDatabase_Lbl";
            this.StagingDatabase_Lbl.Size = new System.Drawing.Size(223, 16);
            this.StagingDatabase_Lbl.TabIndex = 9;
            this.StagingDatabase_Lbl.Text = "Staging Database Connection String";
            // 
            // About_btn
            // 
            this.About_btn.Location = new System.Drawing.Point(650, 48);
            this.About_btn.Name = "About_btn";
            this.About_btn.Size = new System.Drawing.Size(75, 29);
            this.About_btn.TabIndex = 80;
            this.About_btn.Text = "About";
            this.About_btn.UseVisualStyleBackColor = true;
            this.About_btn.Click += new System.EventHandler(this.About_btn_Click);
            // 
            // CreateSrcVwTmpl_Btn
            // 
            this.CreateSrcVwTmpl_Btn.Location = new System.Drawing.Point(42, 114);
            this.CreateSrcVwTmpl_Btn.Name = "CreateSrcVwTmpl_Btn";
            this.CreateSrcVwTmpl_Btn.Size = new System.Drawing.Size(246, 32);
            this.CreateSrcVwTmpl_Btn.TabIndex = 20;
            this.CreateSrcVwTmpl_Btn.Text = "Create Source View Template";
            this.CreateSrcVwTmpl_Btn.UseVisualStyleBackColor = true;
            this.CreateSrcVwTmpl_Btn.Click += new System.EventHandler(this.CreateSrcVwTmpl_Btn_Click);
            // 
            // sourceDBConnection_txtb
            // 
            this.sourceDBConnection_txtb.Location = new System.Drawing.Point(42, 276);
            this.sourceDBConnection_txtb.Name = "sourceDBConnection_txtb";
            this.sourceDBConnection_txtb.Size = new System.Drawing.Size(505, 22);
            this.sourceDBConnection_txtb.TabIndex = 50;
            this.sourceDBConnection_txtb.TextChanged += new System.EventHandler(this.sourceDBConnection_txtb_TextChanged);
            // 
            // SourceDatabase_Lbl
            // 
            this.SourceDatabase_Lbl.AutoSize = true;
            this.SourceDatabase_Lbl.Location = new System.Drawing.Point(42, 257);
            this.SourceDatabase_Lbl.Name = "SourceDatabase_Lbl";
            this.SourceDatabase_Lbl.Size = new System.Drawing.Size(220, 16);
            this.SourceDatabase_Lbl.TabIndex = 13;
            this.SourceDatabase_Lbl.Text = "Source Database Connection String";
            // 
            // test_btn
            // 
            this.test_btn.Location = new System.Drawing.Point(653, 100);
            this.test_btn.Name = "test_btn";
            this.test_btn.Size = new System.Drawing.Size(71, 29);
            this.test_btn.TabIndex = 1000;
            this.test_btn.Text = "test";
            this.test_btn.UseVisualStyleBackColor = true;
            this.test_btn.Visible = false;
            this.test_btn.Click += new System.EventHandler(this.test_btn_Click);
            // 
            // sourceToStagingLocation_txtb
            // 
            this.sourceToStagingLocation_txtb.Location = new System.Drawing.Point(41, 393);
            this.sourceToStagingLocation_txtb.Name = "sourceToStagingLocation_txtb";
            this.sourceToStagingLocation_txtb.Size = new System.Drawing.Size(665, 22);
            this.sourceToStagingLocation_txtb.TabIndex = 70;
            this.sourceToStagingLocation_txtb.TextChanged += new System.EventHandler(this.sourceToStagingLocation_txtb_TextChanged);
            // 
            // sourceToStagingLocation_Lbl
            // 
            this.sourceToStagingLocation_Lbl.AutoSize = true;
            this.sourceToStagingLocation_Lbl.Location = new System.Drawing.Point(42, 374);
            this.sourceToStagingLocation_Lbl.Name = "sourceToStagingLocation_Lbl";
            this.sourceToStagingLocation_Lbl.Size = new System.Drawing.Size(251, 16);
            this.sourceToStagingLocation_Lbl.TabIndex = 15;
            this.sourceToStagingLocation_Lbl.Text = "Source To Staging SSIS Project Location";
            // 
            // CreateS2SPackage_Btn
            // 
            this.CreateS2SPackage_Btn.Location = new System.Drawing.Point(41, 190);
            this.CreateS2SPackage_Btn.Name = "CreateS2SPackage_Btn";
            this.CreateS2SPackage_Btn.Size = new System.Drawing.Size(247, 32);
            this.CreateS2SPackage_Btn.TabIndex = 40;
            this.CreateS2SPackage_Btn.Text = "Create Source To Staging Package";
            this.CreateS2SPackage_Btn.UseVisualStyleBackColor = true;
            this.CreateS2SPackage_Btn.Click += new System.EventHandler(this.CreateS2SPackage_Btn_Click);
            // 
            // sourceDBSchema_txtb
            // 
            this.sourceDBSchema_txtb.Location = new System.Drawing.Point(575, 276);
            this.sourceDBSchema_txtb.Name = "sourceDBSchema_txtb";
            this.sourceDBSchema_txtb.Size = new System.Drawing.Size(131, 22);
            this.sourceDBSchema_txtb.TabIndex = 1001;
            this.sourceDBSchema_txtb.TextChanged += new System.EventHandler(this.sourceDBSchema_txtb_TextChanged);
            // 
            // stagingDBSchema_txtb
            // 
            this.stagingDBSchema_txtb.Location = new System.Drawing.Point(575, 332);
            this.stagingDBSchema_txtb.Name = "stagingDBSchema_txtb";
            this.stagingDBSchema_txtb.Size = new System.Drawing.Size(131, 22);
            this.stagingDBSchema_txtb.TabIndex = 1002;
            this.stagingDBSchema_txtb.TextChanged += new System.EventHandler(this.stagingDBSchema_txtb_TextChanged);
            // 
            // sourceDBSchema_Lbl
            // 
            this.sourceDBSchema_Lbl.AutoSize = true;
            this.sourceDBSchema_Lbl.Location = new System.Drawing.Point(572, 257);
            this.sourceDBSchema_Lbl.Name = "sourceDBSchema_Lbl";
            this.sourceDBSchema_Lbl.Size = new System.Drawing.Size(125, 16);
            this.sourceDBSchema_Lbl.TabIndex = 1003;
            this.sourceDBSchema_Lbl.Text = "Source DB Schema";
            // 
            // stagingDBSchema_Lbl
            // 
            this.stagingDBSchema_Lbl.AutoSize = true;
            this.stagingDBSchema_Lbl.Location = new System.Drawing.Point(572, 313);
            this.stagingDBSchema_Lbl.Name = "stagingDBSchema_Lbl";
            this.stagingDBSchema_Lbl.Size = new System.Drawing.Size(128, 16);
            this.stagingDBSchema_Lbl.TabIndex = 1004;
            this.stagingDBSchema_Lbl.Text = "Staging DB Schema";
            // 
            // dataTransforms_chbx
            // 
            this.dataTransforms_chbx.AutoSize = true;
            this.dataTransforms_chbx.Location = new System.Drawing.Point(308, 197);
            this.dataTransforms_chbx.Name = "dataTransforms_chbx";
            this.dataTransforms_chbx.Size = new System.Drawing.Size(157, 20);
            this.dataTransforms_chbx.TabIndex = 1006;
            this.dataTransforms_chbx.Text = "Use Data Transforms";
            this.dataTransforms_chbx.UseVisualStyleBackColor = true;
            this.dataTransforms_chbx.CheckedChanged += new System.EventHandler(this.dataTransforms_chbx_CheckedChanged);
            // 
            // crmToStagingLocation_txtb
            // 
            this.crmToStagingLocation_txtb.Location = new System.Drawing.Point(41, 550);
            this.crmToStagingLocation_txtb.Name = "crmToStagingLocation_txtb";
            this.crmToStagingLocation_txtb.Size = new System.Drawing.Size(505, 22);
            this.crmToStagingLocation_txtb.TabIndex = 1007;
            this.crmToStagingLocation_txtb.TextChanged += new System.EventHandler(this.crmToStagingLocation_txtb_TextChanged);
            // 
            // crmToStagingLocation_lbl
            // 
            this.crmToStagingLocation_lbl.AutoSize = true;
            this.crmToStagingLocation_lbl.Location = new System.Drawing.Point(38, 531);
            this.crmToStagingLocation_lbl.Name = "crmToStagingLocation_lbl";
            this.crmToStagingLocation_lbl.Size = new System.Drawing.Size(238, 16);
            this.crmToStagingLocation_lbl.TabIndex = 1008;
            this.crmToStagingLocation_lbl.Text = "CRM To Staging SSIS Project Location";
            // 
            // stagingDBImportSchema_txtb
            // 
            this.stagingDBImportSchema_txtb.Location = new System.Drawing.Point(575, 550);
            this.stagingDBImportSchema_txtb.Name = "stagingDBImportSchema_txtb";
            this.stagingDBImportSchema_txtb.Size = new System.Drawing.Size(131, 22);
            this.stagingDBImportSchema_txtb.TabIndex = 1009;
            this.stagingDBImportSchema_txtb.TextChanged += new System.EventHandler(this.stagingDBImportSchema_txtb_TextChanged);
            // 
            // stagingDBImportSchema_lbl
            // 
            this.stagingDBImportSchema_lbl.AutoSize = true;
            this.stagingDBImportSchema_lbl.Location = new System.Drawing.Point(572, 531);
            this.stagingDBImportSchema_lbl.Name = "stagingDBImportSchema_lbl";
            this.stagingDBImportSchema_lbl.Size = new System.Drawing.Size(168, 16);
            this.stagingDBImportSchema_lbl.TabIndex = 1010;
            this.stagingDBImportSchema_lbl.Text = "Staging DB Import Schema";
            // 
            // createCRMToStagingPackage_btn
            // 
            this.createCRMToStagingPackage_btn.Location = new System.Drawing.Point(41, 487);
            this.createCRMToStagingPackage_btn.Name = "createCRMToStagingPackage_btn";
            this.createCRMToStagingPackage_btn.Size = new System.Drawing.Size(247, 32);
            this.createCRMToStagingPackage_btn.TabIndex = 1011;
            this.createCRMToStagingPackage_btn.Text = "Create CRM To Staging Package";
            this.createCRMToStagingPackage_btn.UseVisualStyleBackColor = true;
            this.createCRMToStagingPackage_btn.Click += new System.EventHandler(this.createCRMToStagingPackage_btn_Click);
            // 
            // MyPluginControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.createCRMToStagingPackage_btn);
            this.Controls.Add(this.stagingDBImportSchema_lbl);
            this.Controls.Add(this.stagingDBImportSchema_txtb);
            this.Controls.Add(this.crmToStagingLocation_lbl);
            this.Controls.Add(this.crmToStagingLocation_txtb);
            this.Controls.Add(this.dataTransforms_chbx);
            this.Controls.Add(this.stagingDBSchema_Lbl);
            this.Controls.Add(this.sourceDBSchema_Lbl);
            this.Controls.Add(this.stagingDBSchema_txtb);
            this.Controls.Add(this.sourceDBSchema_txtb);
            this.Controls.Add(this.CreateS2SPackage_Btn);
            this.Controls.Add(this.sourceToStagingLocation_Lbl);
            this.Controls.Add(this.sourceToStagingLocation_txtb);
            this.Controls.Add(this.SourceDatabase_Lbl);
            this.Controls.Add(this.sourceDBConnection_txtb);
            this.Controls.Add(this.CreateSrcVwTmpl_Btn);
            this.Controls.Add(this.test_btn);
            this.Controls.Add(this.About_btn);
            this.Controls.Add(this.StagingDatabase_Lbl);
            this.Controls.Add(this.EntityName_Lbl);
            this.Controls.Add(this.stagingDBConnection_txtb);
            this.Controls.Add(this.EntityCmb);
            this.Controls.Add(this.CreateStgTbl_Btn);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MyPluginControl";
            this.Size = new System.Drawing.Size(833, 602);
            this.Load += new System.EventHandler(this.MyPluginControl_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button CreateStgTbl_Btn;
        private System.Windows.Forms.ComboBox EntityCmb;
        private System.Windows.Forms.TextBox stagingDBConnection_txtb;
        private System.Windows.Forms.Label EntityName_Lbl;
        private System.Windows.Forms.Label StagingDatabase_Lbl;
        private System.Windows.Forms.Button About_btn;
        private System.Windows.Forms.Button CreateSrcVwTmpl_Btn;
        private System.Windows.Forms.TextBox sourceDBConnection_txtb;
        private System.Windows.Forms.Label SourceDatabase_Lbl;
        private System.Windows.Forms.Button test_btn;
        private System.Windows.Forms.TextBox sourceToStagingLocation_txtb;
        private System.Windows.Forms.Label sourceToStagingLocation_Lbl;
        private System.Windows.Forms.Button CreateS2SPackage_Btn;
        private System.Windows.Forms.TextBox sourceDBSchema_txtb;
        private System.Windows.Forms.TextBox stagingDBSchema_txtb;
        private System.Windows.Forms.Label sourceDBSchema_Lbl;
        private System.Windows.Forms.Label stagingDBSchema_Lbl;
        private System.Windows.Forms.CheckBox dataTransforms_chbx;
        private System.Windows.Forms.TextBox crmToStagingLocation_txtb;
        private System.Windows.Forms.Label crmToStagingLocation_lbl;
        private System.Windows.Forms.TextBox stagingDBImportSchema_txtb;
        private System.Windows.Forms.Label stagingDBImportSchema_lbl;
        private System.Windows.Forms.Button createCRMToStagingPackage_btn;
    }
}
