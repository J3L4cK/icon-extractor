using Shell32;
using System.Linq;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace IconExtractor {
	public class IconExtractorCLI {

		private RootCommand RootCommand { get; init; }

		public IconExtractorCLI() {
			RootCommand = new RootCommand {
				new Option<string>(
					alias: "--file",
					arity: new ArgumentArity(1, 1),
					description: "path to extract iamge from file \r\nRequired") {
					IsRequired = true
				},
				new Option<string>(
					alias: "--output",
					arity: new ArgumentArity(1, 1),
					description: "path to save the extracted image \r\nRequired") {
					IsRequired = true
				},
				new Option<string>(
					alias: "--filename",
					description: "alternative name for the extracted image \r\nOptional \r\nDefault: filename from extracted file"),
				new Option<int>(
					alias: "--size",
					getDefaultValue: () => 16,
					description: "try to get the size from .ico"),
				new Option<int>(
					alias: "--resize",
					getDefaultValue: () => -1,
					description: "resize the extracted image with loosing quality \r\nOptional \r\nDefault: original size"),
				new Option<bool>(
					alias: "--search-icon",
					description: "search for .ico file in the .exe directory"),
			};


			RootCommand.Description = "CLI to extract icon from file";
		}


		private static bool IsEmpty(string value) => string.IsNullOrEmpty(value);

		private static bool IsOS(OSPlatform platform) => RuntimeInformation.IsOSPlatform(platform);

		public int Execute(string[] args) {
			RootCommand.Handler = CommandHandler.Create<string, string, string, int, int, bool>(CmdHandler);

			return RootCommand.InvokeAsync(args).Result;
		}

		private static void CmdHandler(string file, string output, string filename, int size, int resize, bool searchIcon) {
			var origFilename = Path.GetFileNameWithoutExtension(file);
			var outputName = IsEmpty(filename) ? origFilename : filename;
			var outputPath = Path.Combine(output, outputName);
			outputPath += ".png";

			var extension = Path.GetExtension(file);
			var isIcoFile = extension == ".ico";
			if (extension == ".lnk") {
				file = GetExePath(file);
			}

			if(searchIcon) {
				var directory = Path.GetDirectoryName(file);
				var files = Directory.GetFiles(directory, "*.ico");

				var ico = files.FirstOrDefault();
				if(!IsEmpty(ico)) {
					isIcoFile = true;
					file = ico;
				}
			}

			Icon icon;
			if(isIcoFile) {
				using var stream = File.OpenRead(file);
				icon = new Icon(stream, new Size(size, size));
			} else {
				icon = Icon.ExtractAssociatedIcon(file);
			}

			using var bitmap = icon.ToBitmap();
			icon.Dispose();

			if (resize < 0) {
				bitmap.MakeTransparent();
				bitmap.Save(outputPath, ImageFormat.Png);
			} else {
				using var image = ResizeImage(bitmap, resize);
				image.Save(outputPath, ImageFormat.Png);
			}
		}

		[SuppressMessage("Interoperability", "CA1416", Justification = "<Pending>")]
		private static string GetExePath(string filePath) {
			var exePath = string.Empty;
			Thread thread = new Thread(() => {
				var path = Path.GetDirectoryName(filePath);
				var filename = Path.GetFileName(filePath);

				Shell shell = new Shell();
				Folder folder = shell.NameSpace(path);
				FolderItem folderItem = folder.ParseName(filename);
				if (folderItem != null) {
					ShellLinkObject link = (ShellLinkObject)folderItem.GetLink;
					exePath = link.Path;
				}
			});
			if (IsOS(OSPlatform.Windows)) {
				thread.SetApartmentState(ApartmentState.STA);
			}

			thread.Start();
			thread.Join();

			return exePath;
		}

		private static Bitmap ResizeImage(Image image, int size) {
			var destRect = new Rectangle(0, 0, size, size);
			var destImage = new Bitmap(size, size);

			destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

			using var graphics = Graphics.FromImage(destImage);
			graphics.CompositingMode = CompositingMode.SourceCopy;
			graphics.CompositingQuality = CompositingQuality.HighQuality;
			graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
			graphics.SmoothingMode = SmoothingMode.HighQuality;
			graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

			using var wrapMode = new ImageAttributes();
			wrapMode.SetWrapMode(WrapMode.TileFlipXY);
			graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);


			return destImage;
		}
	}
}
