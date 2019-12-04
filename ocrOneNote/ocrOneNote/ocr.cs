using System;
using System.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Xml;
using System.Xml.Linq;

/* 添加对COM组件的引用
Microsoft Graph 16.0 Object Library
Microsoft OneNote 15.0 Type Library
*/

namespace ocrOneNote
{
    public class ocr
    {
        // OCR延迟等待时间
        private int waitTime;

        // 构造函数
        public ocr(int wt)
        {
            waitTime = wt;
        }

        // 获取图片的Base64编码
        private Tuple<string, int, int> GetBase64(FileInfo fi)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                Bitmap bp = new Bitmap(fi.FullName);
                switch (fi.Extension.ToLower())
                {
                    case ".jpg":
                        bp.Save(ms, ImageFormat.Jpeg);
                        break;
                    case ".gif":
                        bp.Save(ms, ImageFormat.Gif);
                        break;
                    case ".bmp":
                        bp.Save(ms, ImageFormat.Bmp);
                        break;
                    case ".tif":
                        bp.Save(ms, ImageFormat.Tiff);
                        break;
                    case ".png":
                        bp.Save(ms, ImageFormat.Png);
                        break;
                    case ".emf":
                        bp.Save(ms, ImageFormat.Emf);
                        break;
                    default:
                        return new Tuple<string, int, int>("Unsupported picture formats", 0, 0);
                }
                byte[] buffer = ms.GetBuffer();
                return new Tuple<string, int, int>(Convert.ToBase64String(buffer), bp.Width, bp.Height);
            }
        }

        private Tuple<string, int, int> GetBase64(string strImgPath)
        {
            return GetBase64(new FileInfo(strImgPath));
        }

        // ocr图像识别
        public string OcrImg(FileInfo fi)
        {
            try
            {
                // 新建一个OneNote对象
                var onenoteApp = new OneNote.Application();

                // 新建一个One文件
                string sectionID;
                onenoteApp.OpenHierarchy(AppDomain.CurrentDomain.BaseDirectory + "newfile.one", null, out sectionID, OneNote.CreateFileType.cftSection);

                // 新建一个页面
                string pageID;
                onenoteApp.CreateNewPage(sectionID, out pageID, OneNote.NewPageStyle.npsBlankPageNoTitle);

                // 获取页面ID
                string notebookXml;
                onenoteApp.GetHierarchy(null, OneNote.HierarchyScope.hsPages, out notebookXml);
                var doc = XDocument.Parse(notebookXml);
                var ns = doc.Root.Name.Namespace;
                var pageNode = doc.Descendants(ns + "Page").LastOrDefault();
                var existingPageId = pageNode.Attribute("ID").Value;

                // 构建符合OneNote的XML页面格式
                Tuple<string, int, int> imgInfo = this.GetBase64(fi);
                var page = new XDocument(new XElement(ns + "Page",
                                                new XElement(ns + "Outline",
                                                    new XElement(ns + "OEChildren",
                                                        new XElement(ns + "OE",
                                                            new XElement(ns + "Image",
                                                                new XAttribute("format", fi.Extension.Remove(0, 1)),
                                                                    new XAttribute("originalPageNumber", "0"),
                                                                        new XElement(ns + "Position",
                                                                            new XAttribute("x", "0"),
                                                                                new XAttribute("y", "0"),
                                                                                    new XAttribute("z", "0")),
                                                                                        new XElement(ns + "Size",
                                                                                            new XAttribute("width", imgInfo.Item2),
                                                                                                new XAttribute("height", imgInfo.Item3)),
                                                                                                    new XElement(ns + "Data", imgInfo.Item1)))))));
                page.Root.SetAttributeValue("ID", existingPageId);

                // 更新OneNote页面内容
                onenoteApp.UpdatePageContent(page.ToString(), DateTime.MinValue);

                // 线程休眠时间，单位毫秒，若图片很大，可延长休眠时间
                int fileSize = Convert.ToInt32(fi.Length / 1024 / 1024);
                System.Threading.Thread.Sleep(waitTime * (fileSize > 1 ? fileSize : 1));

                // 获取识别结果
                string pageXml;
                onenoteApp.GetPageContent(existingPageId, out pageXml, OneNote.PageInfo.piBinaryData);

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(pageXml);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("one", ns.ToString());

                XmlNode xmlNode = xmlDoc.SelectSingleNode("//one:Image//one:OCRText", nsmgr);

                string strRet;
                if (xmlNode != null)
                {
                    strRet = xmlNode.InnerText;
                }
                else
                {
                    strRet = "Unrecognized image files";
                }

                // 销毁页面
                onenoteApp.DeleteHierarchy(existingPageId, DateTime.MinValue, true);

                return strRet;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }
    }
}
