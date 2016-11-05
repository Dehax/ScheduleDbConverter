namespace ScheduleDbToPublisherConverter
{
    partial class MainForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.openDbFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.promptLabel = new System.Windows.Forms.Label();
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.buttonChooseFile = new System.Windows.Forms.Button();
            this.savePubFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.exportProgressBar = new System.Windows.Forms.ProgressBar();
            this.exportLabel = new System.Windows.Forms.Label();
            this.exportBackgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // openDbFileDialog
            // 
            this.openDbFileDialog.Filter = "SQLite databases|*.db";
            // 
            // promptLabel
            // 
            this.promptLabel.AutoSize = true;
            this.promptLabel.Location = new System.Drawing.Point(13, 13);
            this.promptLabel.Name = "promptLabel";
            this.promptLabel.Size = new System.Drawing.Size(230, 13);
            this.promptLabel.TabIndex = 0;
            this.promptLabel.Text = "Choose the SQLite database file with schedule:";
            // 
            // textBoxPath
            // 
            this.textBoxPath.Location = new System.Drawing.Point(16, 32);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.Size = new System.Drawing.Size(283, 20);
            this.textBoxPath.TabIndex = 1;
            // 
            // buttonChooseFile
            // 
            this.buttonChooseFile.Location = new System.Drawing.Point(305, 30);
            this.buttonChooseFile.Name = "buttonChooseFile";
            this.buttonChooseFile.Size = new System.Drawing.Size(75, 23);
            this.buttonChooseFile.TabIndex = 2;
            this.buttonChooseFile.Text = "Choose...";
            this.buttonChooseFile.UseVisualStyleBackColor = true;
            this.buttonChooseFile.Click += new System.EventHandler(this.buttonChooseFile_Click);
            // 
            // savePubFileDialog
            // 
            this.savePubFileDialog.DefaultExt = "*.pub";
            this.savePubFileDialog.Filter = "Публикация Publisher 2010|*.pub";
            // 
            // exportProgressBar
            // 
            this.exportProgressBar.Location = new System.Drawing.Point(16, 75);
            this.exportProgressBar.Name = "exportProgressBar";
            this.exportProgressBar.Size = new System.Drawing.Size(283, 23);
            this.exportProgressBar.TabIndex = 3;
            // 
            // exportLabel
            // 
            this.exportLabel.AutoSize = true;
            this.exportLabel.Location = new System.Drawing.Point(16, 59);
            this.exportLabel.Name = "exportLabel";
            this.exportLabel.Size = new System.Drawing.Size(83, 13);
            this.exportLabel.TabIndex = 4;
            this.exportLabel.Text = "Export progress:";
            // 
            // exportBackgroundWorker
            // 
            this.exportBackgroundWorker.WorkerReportsProgress = true;
            this.exportBackgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.exportBackgroundWorker_DoWork);
            this.exportBackgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.exportBackgroundWorker_ProgressChanged);
            this.exportBackgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.exportBackgroundWorker_RunWorkerCompleted);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(392, 123);
            this.Controls.Add(this.exportLabel);
            this.Controls.Add(this.exportProgressBar);
            this.Controls.Add(this.buttonChooseFile);
            this.Controls.Add(this.textBoxPath);
            this.Controls.Add(this.promptLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Расписашка (*.db) -> Publisher 2010 (*.pub)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openDbFileDialog;
        private System.Windows.Forms.Label promptLabel;
        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.Button buttonChooseFile;
        private System.Windows.Forms.SaveFileDialog savePubFileDialog;
        private System.Windows.Forms.ProgressBar exportProgressBar;
        private System.Windows.Forms.Label exportLabel;
        private System.ComponentModel.BackgroundWorker exportBackgroundWorker;
    }
}

