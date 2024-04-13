using PairGenLibrary;

namespace PPT_Pair_GenGUI_Net {
	public partial class FormMain : Form {
		public FormMain() {
			InitializeComponent();
		}

		#region StatusStrip

		private void UpdateProgressBar(int value) {
			if (InvokeRequired) {
				var r = BeginInvoke(new Action(() => { toolStripProgressBar1.Value = value; }));
				EndInvoke(r);
			}
			else {
				toolStripProgressBar1.Value = value;
			}
		}

		private void UpdateStatusStrip(string text) {
			if (InvokeRequired) {
				var r = BeginInvoke(new Action(() => { toolStripStatusLabel1.Text = text; }));
				EndInvoke(r);
			}
			else {
				toolStripStatusLabel1.Text = text;
			}
		}

		private void UpdateGuiTopMost(bool topmost) {
			if (InvokeRequired) {
				var r = BeginInvoke(new Action(() => { TopMost = topmost; }));
				EndInvoke(r);
			}
			else {
				TopMost = topmost;
			}
		}

		#endregion

		private void ProcessFiles(List<string> filePath) {
			int filei = 0, filen = filePath.Count;
			GenCore.ProgressBarUpdate = (float val) => {
				if (val > 100.0f)
					val = 100.0f;
				UpdateProgressBar(Convert.ToInt32((val + filei * 100.0f) / filen));
			};
			GenCore.StatusStripUpdate = (int t, int pi, int pn) => {
				switch (t) {
				case 0:
					UpdateStatusStrip(string.Format(UIString.Status.Reading, filei + 1, filen, pi, pn));
					break;
				case 1:
					UpdateStatusStrip(string.Format(UIString.Status.Processing, filei + 1, filen, pi, pn));
					break;
				case 2:
					UpdateStatusStrip(string.Format(UIString.Status.ReadingMaster, filei + 1, filen));
					break;
				case 3:
					UpdateStatusStrip(string.Format(UIString.Status.ProcessingMaster, filei + 1, filen));
					break;
				};
			};
			UpdateGuiTopMost(true);
			foreach (string file in filePath) {
				if (!GenCore.DoGen(file, out string errStr)) {
					MessageBox.Show(
						string.Format(UIString.MsgBox.FileError, file, errStr),
						Text
					);
				}
				filei++;
			}
			UpdateGuiTopMost(false);
			UpdateProgressBar(100);
			UpdateStatusStrip(UIString.Status.Ready);
		}

		private void FormMain_Load(object sender, EventArgs e) {
			UpdateProgressBar(100);
			UpdateStatusStrip(UIString.Status.Ready);
		}

		private void FormMain_DragEnter(object sender, DragEventArgs e) {
			if (e.Data == null)
				return;
			if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
				e.Effect = DragDropEffects.All;
			}
		}

		private async void FormMain_DragDrop(object sender, DragEventArgs e) {
			if (e.Data == null)
				return;
			var paths = (string[]?)e.Data.GetData(DataFormats.FileDrop);
			if (paths == null)
				return;
			if (paths.Length > 0) {
				await Task.Run(() => ProcessFiles(paths.ToList()));
			}
		}

		private void Button1_Click(object sender, EventArgs e) {
			openFileDialog1.ShowDialog();
		}

		private async void OpenFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e) {
			await Task.Run(() => ProcessFiles(openFileDialog1.FileNames.ToList()));
		}
	}
}
