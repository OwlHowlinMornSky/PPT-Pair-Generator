
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using PairGenLibrary;

namespace PPT_Pair_Gen {
	public partial class Form1 : Form {
		public Form1() {
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
				UpdateProgressBar(Convert.ToInt32((val + filei * 100.0f) / filen));
			};
			GenCore.StatusStripUpdate = (bool t, int pi, int pn) => {
				UpdateStatusStrip(
					string.Format(
						t ? UIString.Status.Processing : UIString.Status.Reading,
						filei + 1, filen, pi, pn
					)
				);
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

		private void Form1_Load(object sender, EventArgs e) {
			UpdateProgressBar(100);
			UpdateStatusStrip(UIString.Status.Ready);
		}

		private void Form1_DragEnter(object sender, DragEventArgs e) {
			if (e.Data == null)
				return;
			if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
				e.Effect = DragDropEffects.All;
			}
		}

		private async void Form1_DragDrop(object sender, DragEventArgs e) {
			if (e.Data == null)
				return;
			if (!e.Data.GetDataPresent(DataFormats.FileDrop)) {
				return;
			}
			var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
			if (paths == null)
				return;
			if (paths.Length > 0) {
				await Task.Run(() => ProcessFiles(paths.ToList()));
			}
		}

		private void button1_Click(object sender, EventArgs e) {
			openFileDialog1.ShowDialog();
		}

		private async void openFileDialog1_FileOk(object sender, CancelEventArgs e) {
			await Task.Run(() => ProcessFiles(openFileDialog1.FileNames.ToList()));
		}

	}
}
