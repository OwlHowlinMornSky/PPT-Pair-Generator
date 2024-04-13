using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace PairGenLibrary {
	public static class GenCore {

		public static Action<int, int, int> StatusStripUpdate = (int t, int x, int y) => { };
		public static Action<float> ProgressBarUpdate = (float x) => { };

		public static bool DoGen(string filePath, out string errorStr) {
			float RatioSum = 0.0f;
			const float OpenRatio = 5.0f;
			const float ReadOldShapeInfoTotalRatio = 20.0f;
			const float ReadMasterRatio = 5.0f;
			const float ResizeRatio = 5.0f;
			const float WriteMasterRatio = 5.0f;
			const float WriteShapeTotalRatio = 55.0f;
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
				// 已经打开
				RatioSum += OpenRatio;
				ProgressBarUpdate(RatioSum);

				// 保存所有幻灯片的所有图形的旧位置参数 的变量
				List<List<Tuple<float, float>>> oldShapes = new List<List<Tuple<float, float>>>();

				// 根据幻灯片数量和图形数量决定progress增量
				ReadRatio = ReadOldShapeInfoTotalRatio / pre.Slides.Count;
				WriteRatio = WriteShapeTotalRatio / pre.Slides.Count;

				// 保存所有幻灯片的所有图形的旧位置参数
				for (int i = 0, n = pre.Slides.Count; i < n; i++) {
					// 开始读取一页
					StatusStripUpdate(0, i + 1, n);

					var slide = pre.Slides[i + 1];
					var slideShapes = new List<Tuple<float, float>>();
					for (int j = 0, m = slide.Shapes.Count; j < m; j++) {
						var shape = slide.Shapes[j + 1];
						slideShapes.Add(new Tuple<float, float>(shape.Left, shape.Width));
					}
					oldShapes.Add(slideShapes);

					// 读完一页
					RatioSum += ReadRatio;
					ProgressBarUpdate(RatioSum);
				}
				// 读取完成
				RatioSum = OpenRatio + ReadOldShapeInfoTotalRatio;
				ProgressBarUpdate(RatioSum);

				// 保存幻灯片母版的图形的旧位置参数 的变量
				List<Tuple<float, float>> oldMasterShapes = new List<Tuple<float, float>>();
				// 保存母版子项...
				List<List<Tuple<float, float>>> oldMasterLayoutShapes = new List<List<Tuple<float, float>>>();

				// 开始读取母版
				StatusStripUpdate(2, 0, 0);
				// 保存幻灯片母版的图形的旧位置参数
				{
					var master = pre.SlideMaster;
					for (int j = 0, m = master.Shapes.Count; j < m; j++) {
						var shape = master.Shapes[j + 1];
						oldMasterShapes.Add(new Tuple<float, float>(shape.Left, shape.Width));
					}
				}
				// 保存母版子项...
				for (int i = 0, n = pre.SlideMaster.CustomLayouts.Count; i < n; i++) {
					var layout = pre.SlideMaster.CustomLayouts[i + 1];
					var layoutShapes = new List<Tuple<float, float>>();
					for (int j = 0, m = layout.Shapes.Count; j < m; j++) {
						var shape = layout.Shapes[j + 1];
						layoutShapes.Add(new Tuple<float, float>(shape.Left, shape.Width));
					}
					oldMasterLayoutShapes.Add(layoutShapes);
				}
				// 保存母版完成
				RatioSum += ReadMasterRatio;
				ProgressBarUpdate(RatioSum);

				// 宽度变为2倍
				float oldWidth = pre.PageSetup.SlideWidth;
				pre.PageSetup.SlideWidth *= 2.0f;
				// 扩大完成
				RatioSum += ResizeRatio;
				ProgressBarUpdate(RatioSum);

				// 开始还原母版
				StatusStripUpdate(3, 0, 0);
				// 还原幻灯片母版的图形的旧位置参数
				{
					var master = pre.SlideMaster;
					for (int j = 0, m = master.Shapes.Count; j < m; j++) {
						var shape = master.Shapes[j + 1];
						var oldShape = oldMasterShapes[j];
						shape.Left = oldShape.Item1;
						shape.Width = oldShape.Item2;

						// 复制一份，同时清空文本
						var range = shape.Duplicate();
						range.Left = shape.Left + oldWidth;
						range.Top = shape.Top;
						if (range.TextFrame.HasText == MsoTriState.msoTrue)
							range.TextFrame.TextRange.Text = "";
						if (range.TextFrame2.HasText == MsoTriState.msoTrue)
							range.TextFrame2.TextRange.Text = "";
					}
				}
				// 还原母版子项...
				for (int i = 0, n = pre.SlideMaster.CustomLayouts.Count; i < n; i++) {
					var layout = pre.SlideMaster.CustomLayouts[i + 1];
					var layoutShapes = oldMasterLayoutShapes[i];
					for (int j = 0, m = layout.Shapes.Count; j < m; j++) {
						var shape = layout.Shapes[j + 1];
						var oldShape = layoutShapes[j];
						shape.Left = oldShape.Item1;
						shape.Width = oldShape.Item2;

						// 复制一份，同时清空文本
						var range = shape.Duplicate();
						range.Left = shape.Left + oldWidth;
						range.Top = shape.Top;
						if (range.TextFrame.HasText == MsoTriState.msoTrue)
							range.TextFrame.TextRange.Text = "";
						if (range.TextFrame2.HasText == MsoTriState.msoTrue)
							range.TextFrame2.TextRange.Text = "";
					}
				}

				// 还原母版完成
				RatioSum += WriteMasterRatio;
				ProgressBarUpdate(RatioSum);

				// 还原所有幻灯片...
				for (int i = 0, n = pre.Slides.Count; i < n; i++) {
					// 开始修改一页
					StatusStripUpdate(1, i + 1, n);
					var slide = pre.Slides[i + 1];
					var slideShapes = oldShapes[i];
					for (int j = 0, m = slide.Shapes.Count; j < m; j++) {
						var shape = slide.Shapes[j + 1];
						var oldShape = slideShapes[j];
						shape.Left = oldShape.Item1;
						shape.Width = oldShape.Item2;

						// 复制一份
						var range = shape.Duplicate();
						range.Left = shape.Left + oldWidth;
						range.Top = shape.Top;
					}
					// 写完一页
					RatioSum += WriteRatio;
					ProgressBarUpdate(RatioSum);
				}
				// 修改完成
				RatioSum =
					OpenRatio +
					ReadOldShapeInfoTotalRatio +
					ReadMasterRatio +
					ResizeRatio +
					WriteMasterRatio +
					WriteShapeTotalRatio;
				ProgressBarUpdate(RatioSum);

				// 计算新名字
				string newName =
					Path.Combine(
						Path.GetDirectoryName(filePath),
						Path.GetFileNameWithoutExtension(filePath) +
						" [pair]"
					);
				string exName = Path.GetExtension(filePath);

				// 另存为
				if (File.Exists(newName + exName)) {
					ulong cnt = 0;
					while (File.Exists(newName + cnt + exName)) {
						cnt++;
					}
					newName += cnt;
				}
				pre.SaveAs(newName + exName);
				// 保存完成
				RatioSum = 100.0f;
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
