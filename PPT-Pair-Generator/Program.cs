
using System;
using System.IO;
using PairGenLibrary;

namespace PPT_Pair_Generator {
	internal class Program {
		static void Main(string[] args) {
			var cd = Directory.GetCurrentDirectory();
			Console.WriteLine($"Current Directory: \"{cd}\"");

			string filePath = $@"{cd}\..\..\..\1.pptx";

			if (GenCore.DoGen(filePath, out string errStr)) {
				Console.WriteLine("Done.");
			}
			else {
				Console.WriteLine("An error occurred: ");
				Console.WriteLine(errStr);
			}

			Console.WriteLine("Press Any Key to Continue.");
			Console.ReadKey();
			return;
		}
	}
}
