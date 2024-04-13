using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace PairGenLibrary {
	public static class GenCore {

		public static Action<bool, int, int> StatusStripUpdate = (bool t, int x, int y) => { };
		public static Action<float> ProgressBarUpdate = (float x) => { };

		public static bool DoGen(string filePath, out string errorStr) {
			float RatioSum = 0.0f;
			const float OpenRatio = 5.0f;
			const float ReadOldShapeInfoTotalRatio = 20.0f;
			const float ResizeRatio = 5.0f;
			const float WriteShapeTotalRatio = 60.0f;
			float ReadRatio = 0.0f;
			float WriteRatio = 0.0f;
			try {
				ProgressBarUpdate(0.0f);

				// 创建PowerPoint应用程序实例
				var pptApp = new PPT.Application {
					Visible = MsoTriState.msoTrue
				};

				// 打开一个演示文稿文件
				var pre = pptApp.Presentations.Open(filePath);

				RatioSum += OpenRatio; // 已经打开
				ProgressBarUpdate(RatioSum);

				List<List<Tuple<float, float>>> oldShapes = new List<List<Tuple<float, float>>>();

				ReadRatio = ReadOldShapeInfoTotalRatio / pre.Slides.Count;
				WriteRatio = WriteShapeTotalRatio / pre.Slides.Count;

				for (int i = 0, n = pre.Slides.Count; i < n; i++) {
					StatusStripUpdate(false, i + 1, n); // 开始读取一页
					var slide = pre.Slides[i + 1];
					var slideShapes = new List<Tuple<float, float>>();
					for (int j = 0, m = slide.Shapes.Count; j < m; j++) {
						var shape = slide.Shapes[j + 1];
						slideShapes.Add(new Tuple<float, float>(shape.Left, shape.Width));
					}
					oldShapes.Add(slideShapes);
					RatioSum += ReadRatio; // 读完一页
					ProgressBarUpdate(RatioSum);
				}
				RatioSum = OpenRatio + ReadOldShapeInfoTotalRatio; // 读取完成
				ProgressBarUpdate(RatioSum);

				float oldWidth = pre.PageSetup.SlideWidth;
				//float halfOldWidth = oldWidth / 2.0f;
				pre.PageSetup.SlideWidth *= 2.0f;

				RatioSum += ResizeRatio; // 扩大完成
				ProgressBarUpdate(RatioSum);

				for (int i = 0, n = pre.Slides.Count; i < n; i++) {
					StatusStripUpdate(true, i + 1, n); // 开始修改一页
					var slide = pre.Slides[i + 1];
					var slideShapes = oldShapes[i];
					for (int j = 0, m = slide.Shapes.Count; j < m; j++) {
						var shape = slide.Shapes[j + 1];
						var oldShape = slideShapes[j];
						shape.Left = oldShape.Item1;
						shape.Width = oldShape.Item2;

						var range = shape.Duplicate();
						range.Left = shape.Left + oldWidth;
						range.Top = shape.Top;
					}
					RatioSum += WriteRatio; // 写完一页
					ProgressBarUpdate(RatioSum);
				}
				RatioSum = OpenRatio + ReadOldShapeInfoTotalRatio + ResizeRatio + WriteShapeTotalRatio; // 修改完成
				ProgressBarUpdate(RatioSum);


				/*foreach (PPT.Slide slide in pre.Slides) {
					List<PPT.Shape> add = new List<PPT.Shape>();
					foreach (PPT.Shape shape in slide.Shapes) {

						if (shape.Left >= halfOldWidth) {
							shape.Left -= halfOldWidth;
						}
						if(shape.Width > oldWidth) {
							shape.Width /= 2.0f;
						}

						add.Add(shape);
					}
					foreach (PPT.Shape shape in add) {
						var range = shape.Duplicate();
						range.Left = shape.Left + oldWidth;
						range.Top = shape.Top;
					}
				}*/

				string newName =
					Path.Combine(
						Path.GetDirectoryName(filePath),
						Path.GetFileNameWithoutExtension(filePath) +
						" [pair]"
					);
				string exName = Path.GetExtension(filePath);

				// 另存为
				if(File.Exists(newName + exName)) {
					ulong cnt = 0;
					while(File.Exists(newName + cnt + exName)) {
						cnt++;
					}
					newName += cnt;
				}
				pre.SaveAs(newName + exName);

				RatioSum = 100.0f; // 保存完成
				ProgressBarUpdate(RatioSum);

				// 关闭演示文稿文件
				pre.Close();

				// 关闭PowerPoint
				//pptApp.Quit();
			}
			catch (Exception ex) {
				errorStr = ex.Message;
				return false;
			}
			errorStr = "";
			return true;
		}

	}
}
