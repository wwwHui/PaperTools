
//using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading.Tasks;
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
        private void RibbonPaperTools_Load(object sender, RibbonUIEventArgs e)
        {
            zoteroColorButton.Tag = DateTime.MinValue;
            zoteroCitationButton.Tag = DateTime.MinValue;
            UpdatePandocVersion();
            UpdateZoteroInfo();
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

        private async void UpdateZoteroInfo()
        {
            Word.Application wordApp = Globals.ThisAddIn.Application;
            if (wordApp.Documents.Count > 0)
            {

                Word.Document doc = wordApp.ActiveDocument;

                // 初始化计数器
                int zoteroItemCount = 0;

                // 遍历文档中的所有书签，统计包含 "ZOTERO_BIBL" 的书签数量
                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    if (bookmark.Name.Contains("ZOTERO"))  // ZOTERO_BIBL
                    {
                        zoteroItemCount++;
                    }
                }

                // 遍历文档中的所有字段，统计包含 "ADDIN ZOTERO_ITEM" 的字段数量
                foreach (Word.Field field in doc.Fields)
                {
                    if (field.Code.Text.Contains("ZOTERO"))
                    {
                        zoteroItemCount++;
                    }
                }
                zoteroInfoButton.Label = $"Zotero[{zoteroItemCount}]";
                string text = $"Zotero items count: {zoteroItemCount}";
                await DisplayStatusBarTextAsync(text);  // 异步更新状态栏
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


        private async Task DisplayStatusBarTextAsync(string message)
        {
            await Task.Run(() =>
            {
                string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string text = $"[PaperTools {currentTime}]: {message}";
                // 更新 StatusBar 内容
                Globals.ThisAddIn.Application.StatusBar = text;
            });
        }

        private void zoteroInfoButton_Click(object sender, RibbonControlEventArgs e)
        {
            UpdateZoteroInfo();
        }

        private async void zoteroColorButton_Click(object sender, RibbonControlEventArgs e)
        {
            
            TimeSpan timeDifference = DateTime.Now - (DateTime)zoteroColorButton.Tag;
            Word.WdColor newColor = Word.WdColor.wdColorBlue; // 蓝色
            if (timeDifference.TotalMilliseconds < 300)
            {
                newColor = Word.WdColor.wdColorAutomatic;  
            }

            // 寻找包含“ZOTERO”的书签  即 Zotero Bibliography
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            int count = 0;
            foreach (Word.Bookmark bookmark in doc.Bookmarks)
            {
                if (bookmark.Name.Contains("ZOTERO"))
                {
                    bookmark.Range.Font.Color = newColor;  // 设置颜色
                    count += 1;
                }
            }
            // 遍历文档中的所有字段，更新 包含 "ADDIN ZOTERO _ITEM" 的字段颜色
            foreach (Word.Field field in doc.Fields)
            {
                if (field.Code.Text.Contains("ZOTERO"))
                {
                    field.Result.Font.Color = newColor;  // 设置颜色;
                    count++;
                }
            }


            zoteroColorButton.Tag = DateTime.Now;
            // doc.Save();  // 保存文档
            // Console.WriteLine(format: "共{0}个书签，刷新其中{1}处与Zotero相关的颜色。", doc.Bookmarks.Count,count);
            string text = $"Zotero Items: {count}";
            //Globals.ThisAddIn.Application.StatusBar = text;
            await DisplayStatusBarTextAsync(text);

        }


        private async void zoteroCitationButton_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            foreach (Word.Field field in doc.Fields)
            {
                // await DisplayStatusBarTextAsync($"Field Text: {field.Code.Text}");

                if (field.Code.Text.Contains("ADDIN ZOTERO_ITEM"))
                {
                    // 输出 Zotero 引用字段的代码
                    Console.WriteLine($"Found Zotero Citation Field: {field.Code.Text}");
                }
            }


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
                ExportLatexFile(texFilePath);
                ExtractImages(Path.GetDirectoryName(texFilePath));
            }

        }

        private void ExportLatexFile(string outputFilePath)
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
                foreach (Word.InlineShape ils in doc.InlineShapes)
                {
                    if (ils == null)  continue;
                    
                    // ils.Type == WdInlineShapeType.wdInlineShapePicture 判断类型
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
                Clipboard.Clear();
                Console.WriteLine("Images extracted successfully.");
            }
            catch (Exception ex)
            {
                // 处理异常
                Console.WriteLine("Exception: " + ex.Message);
            }
        }

        private void buttonReomve_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            if (selection != null && !selection.Text.Equals(string.Empty))
            {
                string selectedText = selection.Text;
                selectedText = selectedText.Replace(" ", ""); // 去除空格
                selectedText = selectedText.Replace("\r", ""); // 去除回车符
                selectedText = selectedText.Replace("\n", ""); // 去除换行符

                selection.Text = selectedText; // 将处理后的文本赋回选中内容
            }
            

        }

        private void buttonENReplace_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            if (selection != null && !selection.Text.Equals(string.Empty))
            {
                string selectedText = selection.Text;
                selectedText = selectedText.Replace("，", ","); // 替换中文逗号为英文逗号
                selectedText = selectedText.Replace("。", "."); // 替换中文句号为英文句号
                selectedText = selectedText.Replace("；", ";"); // 替换中文分号为英文分号
                selectedText = selectedText.Replace("：", ":"); // 替换中文冒号为英文冒号
                                                               // 可以继续添加其他标点符号的替换规则...

                selection.Text = selectedText; // 将处理后的文本赋回选中内容
            }
        }

        private void buttonCNReplace_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;

            if (selection != null && !selection.Text.Equals(string.Empty))
            {
                string selectedText = selection.Text;
                selectedText = selectedText.Replace(",", "，"); // 替换英文逗号为中文逗号
                selectedText = selectedText.Replace(".", "。"); // 替换英文句号为中文句号
                selectedText = selectedText.Replace(";", "；"); // 替换英文分号为中文分号
                selectedText = selectedText.Replace(":", "："); // 替换英文冒号为中文冒号
                                                               // 可以继续添加其他标点符号的替换规则...

                selection.Text = selectedText; // 将处理后的文本赋回选中内容
            }
        }

        private async void wordCitationColor_Click(object sender, RibbonControlEventArgs e)
        {
            TimeSpan timeDifference = DateTime.Now - (DateTime)zoteroCitationButton.Tag;
            Word.WdColor newColor = Word.WdColor.wdColorGreen; // 色
            if (timeDifference.TotalMilliseconds < 300)
            {
                newColor = Word.WdColor.wdColorAutomatic;
            }

            // 获取当前文档
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            // 遍历文档中的每个交叉引用
            int count = 0;
            foreach (Word.Field field in doc.Fields)
            {
                if (field.Type == Word.WdFieldType.wdFieldRef)
                {
                    field.Result.Font.Color = newColor;  // 设置颜色;
                    ++count;
                }
            }

            zoteroCitationButton.Tag = DateTime.Now;
            string text = $"Fields: {count}/{doc.Fields.Count}";
            // Globals.ThisAddIn.Application.StatusBar = text;
            await DisplayStatusBarTextAsync(text);
        }


    }
}
