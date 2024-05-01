namespace PPT_Pair_GenGUI_Net {
	partial class FormMain {
		/// <summary>
		///  Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		///  Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		///  Required method for Designer support - do not modify
		///  the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
			statusStrip1 = new StatusStrip();
			toolStripProgressBar1 = new ToolStripProgressBar();
			toolStripStatusLabel1 = new ToolStripStatusLabel();
			button1 = new Button();
			openFileDialog1 = new OpenFileDialog();
			label1 = new Label();
			label2 = new Label();
			statusStrip1.SuspendLayout();
			SuspendLayout();
			// 
			// statusStrip1
			// 
			statusStrip1.Items.AddRange(new ToolStripItem[] { toolStripProgressBar1, toolStripStatusLabel1 });
			statusStrip1.Location = new Point(0, 301);
			statusStrip1.Name = "statusStrip1";
			statusStrip1.Size = new Size(452, 22);
			statusStrip1.SizingGrip = false;
			statusStrip1.TabIndex = 0;
			statusStrip1.Text = "statusStrip1";
			// 
			// toolStripProgressBar1
			// 
			toolStripProgressBar1.Name = "toolStripProgressBar1";
			toolStripProgressBar1.Size = new Size(200, 16);
			// 
			// toolStripStatusLabel1
			// 
			toolStripStatusLabel1.Name = "toolStripStatusLabel1";
			toolStripStatusLabel1.Size = new Size(131, 17);
			toolStripStatusLabel1.Text = "toolStripStatusLabel1";
			// 
			// button1
			// 
			button1.Location = new Point(18, 85);
			button1.Name = "button1";
			button1.Size = new Size(96, 39);
			button1.TabIndex = 1;
			button1.Text = "手动选择文件";
			button1.UseVisualStyleBackColor = true;
			button1.Click += Button1_Click;
			// 
			// openFileDialog1
			// 
			openFileDialog1.DefaultExt = "pptx";
			openFileDialog1.Filter = "Microsoft PowerPoint 演示文稿|*.pptx";
			openFileDialog1.Multiselect = true;
			openFileDialog1.FileOk += OpenFileDialog1_FileOk;
			// 
			// label1
			// 
			label1.AutoSize = true;
			label1.Font = new Font("Microsoft YaHei UI", 21.75F, FontStyle.Regular, GraphicsUnit.Point, 134);
			label1.Location = new Point(12, 9);
			label1.Name = "label1";
			label1.Size = new Size(220, 38);
			label1.TabIndex = 2;
			label1.Text = "拖入文件即处理";
			// 
			// label2
			// 
			label2.AutoSize = true;
			label2.Location = new Point(18, 57);
			label2.Name = "label2";
			label2.Size = new Size(211, 17);
			label2.TabIndex = 3;
			label2.Text = "文件将保存至同目录，名称后加 [pair]";
			// 
			// FormMain
			// 
			AllowDrop = true;
			AutoScaleDimensions = new SizeF(7F, 17F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(452, 323);
			Controls.Add(label2);
			Controls.Add(label1);
			Controls.Add(button1);
			Controls.Add(statusStrip1);
			FormBorderStyle = FormBorderStyle.FixedSingle;
			Icon = (Icon)resources.GetObject("$this.Icon");
			MaximizeBox = false;
			Name = "FormMain";
			Text = "PPT Pair Generator";
			Load += FormMain_Load;
			DragDrop += FormMain_DragDrop;
			DragEnter += FormMain_DragEnter;
			statusStrip1.ResumeLayout(false);
			statusStrip1.PerformLayout();
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion

		private StatusStrip statusStrip1;
		private ToolStripProgressBar toolStripProgressBar1;
		private ToolStripStatusLabel toolStripStatusLabel1;
		private Button button1;
		private OpenFileDialog openFileDialog1;
		private Label label1;
		private Label label2;
	}
}
