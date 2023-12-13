
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace PaperTools
{
    public partial class RibbonPaperTools
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

            UpdatePandocVersion();
        }

        private void UpdatePandocVersion()
        {
            string pandocVersion = GetPandocVersion();
            if (pandocVersion.Equals("none"))
            {
                buttonPandocVersion.Label = "配置Pandoc";
                buttonExportLatex.Enabled = false;
            }
            else
            {
                buttonPandocVersion.Label = "Pandoc " + pandocVersion;
                buttonExportLatex.Enabled = true;
            }
        }

        private string GetPandocVersion()
        {

            try
            {
                // 执行 pandoc --version 命令
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "pandoc",
                    Arguments = "--version",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    if (process != null)
                    {
                        process.WaitForExit();

                        // 读取 Pandoc 版本信息
                        string pandocOutput = process.StandardOutput.ReadToEnd();

                        // 提取版本号信息
                        string[] outputLines = pandocOutput.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string line in outputLines)
                        {
                            if (line.StartsWith("pandoc "))
                            {
                                return line.Substring("pandoc ".Length);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 处理异常
                Console.WriteLine("Exception: " + ex.Message);
            }

            return "none";
        }

        private void citationColor_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Range currentRange = doc.Content;

            List<Bookmark> zoteroBookmarks = new List<Bookmark>();

            // 寻找包含“ZOTERO”的书签
            foreach (Bookmark bookmark in doc.Bookmarks)
            {
                if (bookmark.Name.Contains("ZOTERO"))
                {
                    zoteroBookmarks.Add(bookmark);
                }
            }
            if (zoteroBookmarks.Count > 0)
            {
                WdColor newColor = WdColor.wdColorAutomatic;
                if (zoteroBookmarks[0].Range.Font.Color == newColor)
                {
                    newColor = WdColor.wdColorBlue;  // 蓝色
                }

                foreach (Bookmark bookmark in zoteroBookmarks)
                {
                    bookmark.Range.Font.Color = newColor;  // 设置颜色
                }
            }
            // doc.Save();  // 保存文档


        }

        private void buttonExportLatex_Click(object sender, RibbonControlEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "LaTeX files (*.tex)|*.tex|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 1; // 默认选择第一个过滤器
            saveFileDialog.RestoreDirectory = true; // 恢复之前的目录

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            saveFileDialog.FileName = Path.GetFileNameWithoutExtension(doc.FullName) + ".tex"; ;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string texFilePath = saveFileDialog.FileName;
                exportLatexFile(texFilePath);
                ExtractImages(Path.GetDirectoryName(texFilePath));
            }

        }

        private void exportLatexFile(string outputFilePath)
        {
            // 获取当前活动的 Word 应用程序和文档
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            try
            {
                if (!string.IsNullOrEmpty(outputFilePath))
                {
                    // 执行 pandoc 命令，将 Word 文档转换为 LaTeX
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "pandoc",
                        Arguments = $"-s \"{doc.FullName}\" -o \"{outputFilePath}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    using (Process process = Process.Start(psi))
                    {
                        if (process != null)
                        {
                            process.WaitForExit();

                            // 输出 Pandoc 命令执行结果到控制台
                            Console.WriteLine(process.StandardOutput.ReadToEnd());
                            Console.WriteLine(process.StandardError.ReadToEnd());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 处理异常
                Console.WriteLine("Exception: " + ex.Message);
            }
        }

        public void ExtractImages(string outputFolder)
        {
            string mediaFolderPath = Path.Combine(outputFolder, "media");
            Directory.CreateDirectory(mediaFolderPath);
            // 获取当前活动的 Word 应用程序和文档
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            try
            {
                int imageIndex = 0;
                foreach (Microsoft.Office.Interop.Word.InlineShape ils in doc.InlineShapes)
                {
                    if (ils != null)
                    {
                        if (ils.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture)
                        {
                            ils.Select();
                            Globals.ThisAddIn.Application.Selection.CopyAsPicture();
                            IDataObject ido = Clipboard.GetDataObject();
                            if (ido != null)
                            {
                                if (ido.GetDataPresent(DataFormats.Bitmap))
                                {
                                    Bitmap bmp = (Bitmap)ido.GetData(DataFormats.Bitmap);
                                    string imagePath = Path.Combine(mediaFolderPath, $"Image{++imageIndex}.png");
                                    bmp.Save(imagePath, ImageFormat.Png);
                                }
                            }
                        }
                    }
                }
                Console.WriteLine("Images extracted successfully.");
            }
            catch (Exception ex)
            {
                // 处理异常
                Console.WriteLine("Exception: " + ex.Message);
            }
        }


    }
}
