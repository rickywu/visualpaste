using System;
using Microsoft.Office.Tools.Ribbon;
using System.Text.RegularExpressions;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace VisualPaste
{
    public partial class CangjieRibbon
    {
        private void CangjieImportRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // 为按钮注册点击事件
            btnTableTitle.Click += new RibbonControlEventHandler(CheckDocument);
            btnCopyImp.Click += new RibbonControlEventHandler(CheckDocument);
        }

        // 导出文件方法
        private void CheckDocument(object sender, RibbonControlEventArgs e)
        {
            Document Doc = Globals.ThisAddIn.Application.ActiveDocument;
            switch (e.Control.Id)
            {
                // 判断点击的按钮ID
                case "btnTableTitle":
                    foreach (Table tbl in Doc.Tables)
                    {
                        tbl.ApplyStyleHeadingRows = true;   
                    }
                    break;
                case "btnCopyImp":
                    string docName = Doc.FullName;
                    //select all and copy
                    //Doc.Application.ActiveDocument.Content.Select();
                    //Doc.Application.Selection.Copy();
                    string fileName = Path.Combine(Path.GetTempPath(), "CangJieSaveAs.html");
                    object missing = System.Reflection.Missing.Value;
                    object FileName = fileName;
                    object FileFormat = WdSaveFormat.wdFormatHTML;
                    Doc.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
                    Doc.Application.ActiveDocument.SaveAs2(ref FileName, ref FileFormat);
                    Doc.Application.ActiveDocument.Close();
                    Globals.ThisAddIn.Application.Documents.Open(docName);

                    //load to process
                    var saveDoc = new HtmlAgilityPack.HtmlDocument();
                    saveDoc.OptionDefaultStreamEncoding = Encoding.UTF8;
                    saveDoc.DetectEncodingAndLoad(fileName, true);
                    //remove comment node
                    var commentNodes = saveDoc.DocumentNode.SelectNodes("//comment()");
                    if (commentNodes != null)
                    {
                        foreach (var comment in commentNodes)
                        {
                            if (!comment.InnerText.StartsWith("<!DOCTYPE"))
                            {
                                comment.Remove();
                            }
                        }
                    }
                    //parse img source
                    var imgNodes = saveDoc.DocumentNode.SelectNodes("//img/@src");
                    if (imgNodes != null)
                    {
                        foreach (var img in imgNodes)
                        {
                            string imageName = img.GetAttributeValue("src", "");
                            string localPath = Path.Combine(Path.GetTempPath(), imageName);
                            string imgBase64String = GetBase64StringForImage(localPath);
                            string imageType = imageName.Substring(imageName.Length - 3);
                            //use correct MIME type for jpeg image
                            imageType = (imageType == "jpg") ? imageType.Replace("jpg","jpeg") : imageType;
                            img.SetAttributeValue("src", "data:image/" + imageType + ";base64," + imgBase64String);
                        }
                    }
                    //clear html
                    String bodyHtml = saveDoc.DocumentNode.SelectSingleNode("//body").InnerHtml;
                    bodyHtml = Regex.Replace(bodyHtml, @"( |\t|\r?\n)\1+", "$1");
                    bodyHtml = bodyHtml.Trim(Environment.NewLine.ToCharArray());

                    //copy to clipboard and delete temp files
                    Utilities.CopyHtmlToClipBoard(bodyHtml);
                    File.Delete(fileName);
                    Directory.Delete(fileName.Replace("html","files"), true);
                    break;
                default:
                    return;
            }
        }
        protected static string GetBase64StringForImage(string imgPath)
        {
            byte[] imageBytes = System.IO.File.ReadAllBytes(imgPath);
            string base64String = Convert.ToBase64String(imageBytes);
            return base64String;
        }
    }
}
