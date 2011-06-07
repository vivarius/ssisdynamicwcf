namespace SSISWCFTask100
{
    partial class frmEditProperties
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmEditProperties));
            this.cmbURL = new System.Windows.Forms.ComboBox();
            this.btGO = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbServices = new System.Windows.Forms.ComboBox();
            this.grdParameters = new System.Windows.Forms.DataGridView();
            this.grdColParams = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.grdColDirection = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.grdColVars = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.grdColExpression = new System.Windows.Forms.DataGridViewButtonColumn();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbMethods = new System.Windows.Forms.ComboBox();
            this.cmbReturnVariable = new System.Windows.Forms.ComboBox();
            this.lbOutputValue = new System.Windows.Forms.Label();
            this.btSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.lstCol1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lstCol2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            ((System.ComponentModel.ISupportInitialize)(this.grdParameters)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbURL
            // 
            this.cmbURL.FormattingEnabled = true;
            this.cmbURL.Location = new System.Drawing.Point(114, 12);
            this.cmbURL.Name = "cmbURL";
            this.cmbURL.Size = new System.Drawing.Size(399, 21);
            this.cmbURL.TabIndex = 0;
            this.cmbURL.SelectedIndexChanged += new System.EventHandler(this.cmbURL_SelectedIndexChanged);
            // 
            // btGO
            // 
            this.btGO.Location = new System.Drawing.Point(519, 12);
            this.btGO.Name = "btGO";
            this.btGO.Size = new System.Drawing.Size(42, 23);
            this.btGO.TabIndex = 1;
            this.btGO.Text = "&Go";
            this.btGO.UseVisualStyleBackColor = true;
            this.btGO.Click += new System.EventHandler(this.btGO_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Service URL:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Services Contracts:";
            // 
            // cmbServices
            // 
            this.cmbServices.FormattingEnabled = true;
            this.cmbServices.Location = new System.Drawing.Point(114, 36);
            this.cmbServices.Name = "cmbServices";
            this.cmbServices.Size = new System.Drawing.Size(448, 21);
            this.cmbServices.TabIndex = 3;
            this.cmbServices.SelectedIndexChanged += new System.EventHandler(this.cmbServices_SelectedIndexChanged);
            // 
            // grdParameters
            // 
            this.grdParameters.AllowUserToAddRows = false;
            this.grdParameters.AllowUserToDeleteRows = false;
            this.grdParameters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdParameters.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.grdColParams,
            this.grdColDirection,
            this.grdColVars,
            this.grdColExpression});
            this.grdParameters.Location = new System.Drawing.Point(11, 104);
            this.grdParameters.Name = "grdParameters";
            this.grdParameters.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.grdParameters.Size = new System.Drawing.Size(550, 173);
            this.grdParameters.TabIndex = 25;
            this.grdParameters.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.grdParameters_CellContentClick);
            // 
            // grdColParams
            // 
            this.grdColParams.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.grdColParams.Frozen = true;
            this.grdColParams.HeaderText = "Parameters";
            this.grdColParams.Name = "grdColParams";
            this.grdColParams.ReadOnly = true;
            this.grdColParams.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.grdColParams.Width = 66;
            // 
            // grdColDirection
            // 
            this.grdColDirection.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.grdColDirection.FillWeight = 40F;
            this.grdColDirection.HeaderText = "Param Type";
            this.grdColDirection.Name = "grdColDirection";
            this.grdColDirection.ReadOnly = true;
            this.grdColDirection.Width = 89;
            // 
            // grdColVars
            // 
            this.grdColVars.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.grdColVars.DropDownWidth = 300;
            this.grdColVars.HeaderText = "Variables";
            this.grdColVars.MaxDropDownItems = 10;
            this.grdColVars.Name = "grdColVars";
            this.grdColVars.Sorted = true;
            this.grdColVars.Width = 240;
            // 
            // grdColExpression
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.Padding = new System.Windows.Forms.Padding(2);
            this.grdColExpression.DefaultCellStyle = dataGridViewCellStyle1;
            this.grdColExpression.HeaderText = "f(x)";
            this.grdColExpression.Name = "grdColExpression";
            this.grdColExpression.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.grdColExpression.Text = "f(x)";
            this.grdColExpression.ToolTipText = "Expressions...";
            this.grdColExpression.Width = 30;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 88);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 13);
            this.label3.TabIndex = 26;
            this.label3.Text = "Parameters:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 66);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 13);
            this.label4.TabIndex = 28;
            this.label4.Text = "Methods:";
            // 
            // cmbMethods
            // 
            this.cmbMethods.FormattingEnabled = true;
            this.cmbMethods.Location = new System.Drawing.Point(113, 63);
            this.cmbMethods.Name = "cmbMethods";
            this.cmbMethods.Size = new System.Drawing.Size(448, 21);
            this.cmbMethods.TabIndex = 27;
            this.cmbMethods.SelectedIndexChanged += new System.EventHandler(this.cmbMethods_SelectedIndexChanged);
            // 
            // cmbReturnVariable
            // 
            this.cmbReturnVariable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbReturnVariable.FormattingEnabled = true;
            this.cmbReturnVariable.Location = new System.Drawing.Point(113, 289);
            this.cmbReturnVariable.Name = "cmbReturnVariable";
            this.cmbReturnVariable.Size = new System.Drawing.Size(448, 21);
            this.cmbReturnVariable.TabIndex = 30;
            // 
            // lbOutputValue
            // 
            this.lbOutputValue.AutoSize = true;
            this.lbOutputValue.Location = new System.Drawing.Point(8, 292);
            this.lbOutputValue.Name = "lbOutputValue";
            this.lbOutputValue.Size = new System.Drawing.Size(69, 13);
            this.lbOutputValue.TabIndex = 29;
            this.lbOutputValue.Text = "Output Value";
            // 
            // btSave
            // 
            this.btSave.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btSave.Location = new System.Drawing.Point(405, 321);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(75, 23);
            this.btSave.TabIndex = 32;
            this.btSave.Text = "OK";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(486, 321);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 31;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.lstCol1,
            this.lstCol2});
            this.listView1.Location = new System.Drawing.Point(-2, 382);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(151, 28);
            this.listView1.TabIndex = 33;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.List;
            this.listView1.Visible = false;
            // 
            // lstCol1
            // 
            this.lstCol1.Text = "[ServiceContract]";
            // 
            // lstCol2
            // 
            this.lstCol2.Text = "Binding";
            // 
            // frmEditProperties
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(573, 352);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.btSave);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.cmbReturnVariable);
            this.Controls.Add(this.lbOutputValue);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbMethods);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.grdParameters);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbServices);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btGO);
            this.Controls.Add(this.cmbURL);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmEditProperties";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Edit task properties";
            ((System.ComponentModel.ISupportInitialize)(this.grdParameters)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbURL;
        private System.Windows.Forms.Button btGO;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbServices;
        private System.Windows.Forms.DataGridView grdParameters;
        private System.Windows.Forms.DataGridViewTextBoxColumn grdColParams;
        private System.Windows.Forms.DataGridViewTextBoxColumn grdColDirection;
        private System.Windows.Forms.DataGridViewComboBoxColumn grdColVars;
        private System.Windows.Forms.DataGridViewButtonColumn grdColExpression;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbMethods;
        private System.Windows.Forms.ComboBox cmbReturnVariable;
        private System.Windows.Forms.Label lbOutputValue;
        private System.Windows.Forms.Button btSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader lstCol1;
        private System.Windows.Forms.ColumnHeader lstCol2;

    }
}