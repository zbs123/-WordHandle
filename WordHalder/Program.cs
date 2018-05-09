using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Xml;
using System.Xml.Xsl;
using WordHalder;

namespace MyWord
{
    class Program
    {
        static void Main(string[] args)
        {

            ExcelHanlder eh = new ExcelHanlder();
            eh.ExcelToDataTable("nihao", true, @"C:\Users\admin\Desktop\nihao.xlsx");


            //MyXml1(@"C:\Users\admin\Desktop\测试\2121 - 副本.docx", "");
            // ThumbnailMaker.MakeThumbnail(@"C:\Users\admin\Desktop\测试\aa\word\media\image1.png", @"C:\Users\admin\Desktop" +  "\\aaaaa.png", 500, 50, ThumbnailMode.Cut);
            ////CreateMinImage(@"C:\Users\admin\Desktop\测试\aa\word\media\image1.png", @"C:\Users\admin\Desktop" + "\\aaaaa.jpg", 500, 500);
            //Console.WriteLine("OK");

            //Console.ReadKey();

        }
        
        public static void CopyDirectory(string srcPath, string destPath)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //获取目录下（不包含子目录）的文件和子目录
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)     //判断是否文件夹
                    {
                        if (!Directory.Exists(destPath + "\\" + i.Name))
                        {
                            Directory.CreateDirectory(destPath + "\\" + i.Name);   //目标目录下不存在此文件夹即创建子文件夹
                        }
                        CopyDirectory(i.FullName, destPath + "\\" + i.Name);    //递归调用复制子文件夹
                    }
                    else
                    {
                        File.Copy(i.FullName, destPath + "\\" + i.Name, true);      //不是文件夹即复制文件，true表示可以覆盖同名文件
                    }
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }
        /// <summary>
        /// 处理段落
        /// </summary>
        /// <param name="xml">w:p节点</param>
        /// <param name="m">命名空间</param>
        /// <param name="myPath">源文件路径</param>
        /// <returns></returns>
        public static string ParagraphsHandler(XmlNode xml, System.Xml.XmlNamespaceManager m, string myPath)
        {
            string myHtmlStr = "<p>";

            foreach (XmlNode r in xml.SelectNodes("./w:r", m))
            {
                XmlNodeList rt = r.SelectNodes("./w:t", m);//取当前节点下的w：t
                XmlNodeList robject = r.SelectNodes("./w:object", m);//取当前节点下的w：object
                XmlNodeList rdrawing = r.SelectNodes("./w:drawing", m);
                XmlNodeList rmc = r.SelectNodes("./mc:AlternateContent", m);

                if (rt.Count > 0)
                {
                    myHtmlStr += rt[0].InnerText.Replace(" ", "&nbsp;");
                }
                //v:imagedata
                if (robject.Count > 0)
                {
                    foreach (XmlNode o in robject)
                    {
                        XmlNode img = o.SelectSingleNode(".//v:imagedata", m);//获取imagedata节点
                        var rid = img.Attributes["r:id"].Value;//图片id
                        XmlDocument xmldocRel = new XmlDocument();
                        xmldocRel.Load(myPath + "\\aa\\word\\_rels\\document.xml.rels");//加载关系文档
                        System.Xml.XmlNamespaceManager m1 = new System.Xml.XmlNamespaceManager(xmldocRel.NameTable);
                        m1.AddNamespace("ns", "http://schemas.openxmlformats.org/package/2006/relationships");
                        XmlNode imgxml = xmldocRel.SelectSingleNode("/ns:Relationships/ns:Relationship[@Id='" + rid + "']", m1);
                        string imgPath = imgxml.Attributes["Target"].Value;//相对路径
                        string path = myPath + "\\aa\\word\\" + imgPath.Replace("/", "\\");//图片绝对路径
                        string newPath = "";
                        if (imgPath.IndexOf(".wmf") != -1)
                        {
                            var im = Bitmap.FromFile(path);
                            newPath = path.Split('.')[0] + ".png";
                            im.Save(newPath, ImageFormat.Png);
                        }
                        else
                        {
                            newPath = path;
                        }

                        string strBase64 = GetBase64FromImage(newPath);
                        myHtmlStr += "<img src=\"data:image/png;base64," + strBase64 + "\">";
                    }
                }
                //a:blip
                if (rdrawing.Count > 0)
                {
                    foreach (XmlNode o in rdrawing)
                    {
                        XmlNode img = o.SelectSingleNode(".//a:blip", m);//获取imagedata节点
                        var rid = img.Attributes["r:embed"].Value;//图片id
                        XmlDocument xmldocRel = new XmlDocument();
                        xmldocRel.Load(myPath + "\\aa\\word\\_rels\\document.xml.rels");//加载关系文档
                        System.Xml.XmlNamespaceManager m1 = new System.Xml.XmlNamespaceManager(xmldocRel.NameTable);
                        m1.AddNamespace("ns", "http://schemas.openxmlformats.org/package/2006/relationships");
                        XmlNode imgxml = xmldocRel.SelectSingleNode("/ns:Relationships/ns:Relationship[@Id='" + rid + "']", m1);
                        string imgPath = imgxml.Attributes["Target"].Value;//相对路径
                        string path = myPath + "\\aa\\word\\" + imgPath.Replace("/", "\\");//图片绝对路径
                        string newPath = "";
                        if (imgPath.IndexOf(".wmf") != -1)
                        {
                            var im = Bitmap.FromFile(path);
                            newPath = path.Split('.')[0] + ".png";
                            im.Save(newPath, ImageFormat.Png);
                        }
                        else
                        {
                            newPath = path;
                        }

                        string strBase64 = GetBase64FromImage(newPath);
                        myHtmlStr += "<img src=\"data:image/png;base64," + strBase64 + "\">";
                    }
                }
                if (rmc.Count > 0)
                {
                    foreach (XmlNode o in rmc)
                    {
                        XmlNodeList imgs = o.SelectNodes(".//v:imagedata[@r:id]", m);
                        List<string> temp = new List<string>();
                        foreach (XmlNode img in imgs)
                        {
                            //XmlNode img = o.SelectSingleNode(".//v:imagedata[@r:id]", m);//获取imagedata节点
                            bool flag = false;

                            var rid = img.Attributes["r:id"].Value;//图片id
                            if (temp.Count > 0)
                            {
                                foreach (var t in temp)
                                {
                                    if (t == rid)
                                    {
                                        flag = true;
                                    }
                                }
                            }
                            if (flag) continue;
                            temp.Add(rid);
                            XmlDocument xmldocRel = new XmlDocument();
                            xmldocRel.Load(myPath + "\\aa\\word\\_rels\\document.xml.rels");//加载关系文档
                            System.Xml.XmlNamespaceManager m1 = new System.Xml.XmlNamespaceManager(xmldocRel.NameTable);
                            m1.AddNamespace("ns", "http://schemas.openxmlformats.org/package/2006/relationships");
                            XmlNode imgxml = xmldocRel.SelectSingleNode("/ns:Relationships/ns:Relationship[@Id='" + rid + "']", m1);
                            string imgPath = imgxml.Attributes["Target"].Value;//相对路径
                            string path = myPath + "\\aa\\word\\" + imgPath.Replace("/", "\\");//图片绝对路径
                            string newPath = "";
                            if (imgPath.IndexOf(".wmf") != -1)
                            {
                                var im = Bitmap.FromFile(path);
                                newPath = path.Split('.')[0] + ".png";
                                im.Save(newPath, ImageFormat.Png);
                            }
                            else
                            {
                                newPath = path;
                            }

                            string strBase64 = GetBase64FromImage(newPath);
                            myHtmlStr += "<img src=\"data:image/png;base64," + strBase64 + "\">";
                        }



                    }
                }
            }

            myHtmlStr += "</p>";
            return myHtmlStr;
        }
        public static void MyXml1(string fileName, string ansStr)
        {
            string[] myPathArr = fileName.Split('\\');
            string myPath = string.Join("\\", myPathArr, 0, myPathArr.Length - 1);
            //将docx转为zip
            string dfileName = System.IO.Path.ChangeExtension(fileName, ".zip");
            if (File.Exists(dfileName))
            {
                File.Delete(dfileName);
            }
            File.Move(fileName, dfileName);
            //解压
            string err = "error";
            UPZipFile(dfileName, myPath + "\\aa", ref err);

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(myPath + "\\aa\\word\\document.xml");

            System.Xml.XmlNamespaceManager m = new System.Xml.XmlNamespaceManager(xmldoc.NameTable);
            XmlNode rootDoc = xmldoc.ChildNodes[1];//获取根节点
            //添加命名空间
            for (int i = 0; i < rootDoc.Attributes.Count; i++)
            {
                m.AddNamespace(rootDoc.Attributes[i].LocalName, rootDoc.Attributes[i].Value);
            }

            string myXmlStr = "";
            string myHtmlStr = "";
            IList<string> list = new List<string>();
            List<string> list1 = new List<string>();

            foreach (XmlNode xml in xmldoc.SelectNodes("/w:document/w:body/w:p", m))
            {
                string text = xml.InnerText;

                myHtmlStr += ParagraphsHandler(xml, m, myPath);


                XmlNode xmlSibling = xml.NextSibling;

                myXmlStr += "<w:p>" + xml.InnerXml + "</w:p>";
                if (xmlSibling != null && xmlSibling.Name == "w:tbl")
                {
                    //处理表格
                    string tabStr = "<table  border=\"1\" cellspacing=\"0\" cellpadding=\"0\" style=\"text-align:center\"> ";
                    foreach (XmlNode tr in xmlSibling.SelectNodes("./w:tr", m))
                    {
                        tabStr += "<tr>";
                        foreach (XmlNode td in tr.SelectNodes("./w:tc", m))
                        {
                            tabStr += "<td>";
                            foreach (XmlNode tdStr in td.SelectNodes("./w:p", m))
                            {
                                tabStr += ParagraphsHandler(tdStr, m, myPath);
                            }
                            tabStr += "</td>";
                        }
                        tabStr += "</tr>";
                    }
                    tabStr += "</table>";
                    myHtmlStr += tabStr;
                    myXmlStr += "<w:tbl>" + xmlSibling.InnerXml + "</w:tbl>";
                }
                if (text.IndexOf("【类型】") != -1)
                {

                    list.Add(myXmlStr);
                    list1.Add(myHtmlStr);
                    myXmlStr = "";
                    myHtmlStr = "";
                }
            }
            string[] aaaa = list1.ToArray();
            string bbb = String.Join("", aaaa);
            //Directory.Delete(myPath + "\\aa", true);
            int count = list.Count;
        }
        public static void MyXml(string fileName, string ansStr)
        {
            string[] myPathArr = fileName.Split('\\');
            string myPath = string.Join("\\", myPathArr, 0, myPathArr.Length - 1);
            //将docx转为zip
            string dfileName = System.IO.Path.ChangeExtension(fileName, ".zip");
            if (File.Exists(dfileName))
            {
                File.Delete(dfileName);
            }
            File.Move(fileName, dfileName);
            //解压
            string err = "error";
            UPZipFile(dfileName, myPath + "\\aa", ref err);

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(myPath + "\\aa\\word\\document.xml");

            System.Xml.XmlNamespaceManager m = new System.Xml.XmlNamespaceManager(xmldoc.NameTable);
            m.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            m.AddNamespace("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            m.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            m.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
            m.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            m.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            m.AddNamespace("v", "urn:schemas-microsoft-com:vml");
            m.AddNamespace("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            m.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            m.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            m.AddNamespace("w10", "urn:schemas-microsoft-com:office:word");
            m.AddNamespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            m.AddNamespace("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            m.AddNamespace("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            m.AddNamespace("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            m.AddNamespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            m.AddNamespace("wpsCustomData", "http://www.wps.cn/officeDocument/2013/wpsCustomData");

            string myXmlStr = "";
            IList<string> list = new List<string>();
            //XmlNode xn = xmldoc.SelectSingleNode("/w:document", m);
            //string attribute = xn.Attributes["xmlns"].Value;
            string TH = "";
            int count = 0;
            foreach (XmlNode xml in xmldoc.SelectNodes("/w:document/w:body/w:p", m))
            {
                string text = xml.InnerText;
                XmlNode xmlSibling = xml.NextSibling;
                if (text.IndexOf("【题号】") != -1)
                {
                    string th = text.Substring(4).Trim();
                    TH = th;
                    if (th.IndexOf("-") != -1)
                    {
                        string[] ths = th.Split('-');
                        count = int.Parse(ths[1]);
                        if (int.Parse(ths[1]) <= ansStr.Length)
                        {
                            myXmlStr += "<w:p><w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:t>【答案】</w:t></w:r></w:p>";
                            for (var i = int.Parse(ths[0]); i <= int.Parse(ths[1]); i++)
                            {
                                myXmlStr += "<w:p><w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:t>（" + i + "）</w:t></w:r><w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:rPr><w:rFonts w:hint=\"eastAsia\" /><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\" /></w:rPr><w:t>" + ansStr[i - 1] + "</w:t></w:r></w:p>";
                            }
                        }

                    }
                    else
                    {
                        count = int.Parse(th);
                        if (int.Parse(th) <= ansStr.Length)
                        {
                            myXmlStr += "<w:p><w:pPr></w:pPr><w:r><w:t xml:space=\"preserve\">【答案】" + ansStr[int.Parse(th) - 1] + "</w:t></w:r></w:p>";

                        }
                    }
                    continue;
                }
                myXmlStr += "<w:p>" + xml.InnerXml + "</w:p>";
                if (xmlSibling.Name == "w:tbl")
                {
                    myXmlStr += "<w:tbl>" + xmlSibling.InnerXml + "</w:tbl>";
                }
                if (text.IndexOf("【类型】") != -1)
                {
                    if (TH.Length == 0)
                    {
                        count++;
                        TH = count + "";
                    }
                    myXmlStr = "<?xml version=\"1.0\" encoding=\"utf-8\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:wpsCustomData=\"http://www.wps.cn/officeDocument/2013/wpsCustomData\" mc:Ignorable=\"w14 w15 wp14\"><w:body> " + myXmlStr + " </w:body></w:document> ";

                    //拷贝一份解压后的源文件
                    string docxPath = myPath + "\\" + TH;
                    CopyDirectory(myPath + "\\aa", docxPath);
                    //File.Copy(myPath + "\\aa", docxPath, false);
                    //将xml字符串写进xml文档里
                    StringReader Reader = new StringReader(myXmlStr);
                    XmlDocument xmlDoc1 = new XmlDocument();
                    xmlDoc1.Load(Reader);
                    xmlDoc1.Save(docxPath + "\\word\\document.xml");
                    //压缩文件
                    string docxZip = myPath + "\\" + TH + ".zip";
                    Zip(docxPath, docxZip, ref err);
                    Directory.Delete(docxPath, true);
                    //zip --docx
                    string docxName = System.IO.Path.ChangeExtension(docxZip, ".docx");
                    if (File.Exists(docxName))
                    {
                        File.Delete(docxName);
                    }
                    File.Move(docxZip, docxName);
                    //File.Delete(dfileName);



                    list.Add(myXmlStr);
                    myXmlStr = "";
                    TH = "";
                }
            }
            Directory.Delete(myPath + "\\aa", true);

        }

        public static bool Zip(string fileToZip, string zipedFile, string password, ref string errorOut)
        {
            bool result = false;
            try
            {
                if (Directory.Exists(fileToZip))
                    result = ZipDirectory(fileToZip, zipedFile, password);
                else if (File.Exists(fileToZip))
                    result = ZipFile(fileToZip, zipedFile, password);
            }
            catch (Exception ex)
            {
                errorOut = ex.Message;
            }
            return result;
        }

        /// <summary>     
        /// 压缩文件或文件夹 ----无密码   
        /// </summary>     
        /// <param name="fileToZip">要压缩的路径-文件夹或者文件</param>     
        /// <param name="zipedFile">压缩后的文件名</param>  
        /// <param name="errorOut">如果失败返回失败信息</param>  
        /// <returns>压缩结果</returns>     
        public static bool Zip(string fileToZip, string zipedFile, ref string errorOut)
        {
            bool result = false;
            try
            {
                if (Directory.Exists(fileToZip))
                    result = ZipDirectory(fileToZip, zipedFile, null);
                else if (File.Exists(fileToZip))
                    result = ZipFile(fileToZip, zipedFile, null);
            }
            catch (Exception ex)
            {
                errorOut = ex.Message;
            }
            return result;
        }


        #region 内部处理方法  
        /// <summary>     
        /// 压缩文件     
        /// </summary>     
        /// <param name="fileToZip">要压缩的文件全名</param>     
        /// <param name="zipedFile">压缩后的文件名</param>     
        /// <param name="password">密码</param>     
        /// <returns>压缩结果</returns>     
        private static bool ZipFile(string fileToZip, string zipedFile, string password)
        {
            bool result = true;
            ZipOutputStream zipStream = null;
            FileStream fs = null;
            ZipEntry ent = null;

            if (!File.Exists(fileToZip))
                return false;

            try
            {
                fs = File.OpenRead(fileToZip);
                byte[] buffer = new byte[fs.Length];
                fs.Read(buffer, 0, buffer.Length);
                fs.Close();

                fs = File.Create(zipedFile);
                zipStream = new ZipOutputStream(fs);
                if (!string.IsNullOrEmpty(password)) zipStream.Password = password;
                ent = new ZipEntry(Path.GetFileName(fileToZip));
                zipStream.PutNextEntry(ent);
                zipStream.SetLevel(6);

                zipStream.Write(buffer, 0, buffer.Length);

            }
            catch (Exception ex)
            {
                result = false;
                throw ex;
            }
            finally
            {
                if (zipStream != null)
                {
                    zipStream.Finish();
                    zipStream.Close();
                }
                if (ent != null)
                {
                    ent = null;
                }
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }
            }
            GC.Collect();
            GC.Collect(1);

            return result;
        }

        /// <summary>  
        /// 压缩文件夹  
        /// </summary>  
        /// <param name="strFile">带压缩的文件夹目录</param>  
        /// <param name="strZip">压缩后的文件名</param>  
        /// <param name="password">压缩密码</param>  
        /// <returns>是否压缩成功</returns>  
        private static bool ZipDirectory(string strFile, string strZip, string password)
        {
            bool result = false;
            if (!Directory.Exists(strFile)) return false;
            if (strFile[strFile.Length - 1] != Path.DirectorySeparatorChar)
                strFile += Path.DirectorySeparatorChar;
            ZipOutputStream s = new ZipOutputStream(File.Create(strZip));
            s.SetLevel(6); // 0 - store only to 9 - means best compression  
            if (!string.IsNullOrEmpty(password)) s.Password = password;
            try
            {
                result = zip(strFile, s, strFile);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                s.Finish();
                s.Close();
            }
            return result;
        }

        /// <summary>  
        /// 压缩文件夹内部方法  
        /// </summary>  
        /// <param name="strFile"></param>  
        /// <param name="s"></param>  
        /// <param name="staticFile"></param>  
        /// <returns></returns>  
        private static bool zip(string strFile, ZipOutputStream s, string staticFile)
        {
            bool result = true;
            if (strFile[strFile.Length - 1] != Path.DirectorySeparatorChar) strFile += Path.DirectorySeparatorChar;
            Crc32 crc = new Crc32();
            try
            {
                string[] filenames = Directory.GetFileSystemEntries(strFile);
                foreach (string file in filenames)
                {

                    if (Directory.Exists(file))
                    {
                        zip(file, s, staticFile);
                    }

                    else // 否则直接压缩文件  
                    {
                        //打开压缩文件  
                        FileStream fs = File.OpenRead(file);

                        byte[] buffer = new byte[fs.Length];
                        fs.Read(buffer, 0, buffer.Length);
                        string tempfile = file.Substring(staticFile.LastIndexOf("\\") + 1);
                        ZipEntry entry = new ZipEntry(tempfile);

                        entry.DateTime = DateTime.Now;
                        entry.Size = fs.Length;
                        fs.Close();
                        crc.Reset();
                        crc.Update(buffer);
                        entry.Crc = crc.Value;
                        s.PutNextEntry(entry);

                        s.Write(buffer, 0, buffer.Length);
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
                throw ex;
            }
            return result;

        }
        #endregion





        /// <summary>     
        /// 解压功能(解压压缩文件到指定目录)---->不需要密码  
        /// </summary>     
        /// <param name="fileToUnZip">待解压的文件</param>     
        /// <param name="zipedFolder">指定解压目标目录</param>    
        /// <param name="errorOut">如果失败返回失败信息</param>   
        /// <returns>解压结果</returns>     
        public static bool UPZipFile(string fileToUnZip, string zipedFolder, ref string errorOut)
        {
            bool result = false;
            try
            {
                result = UPZipFileByPassword(fileToUnZip, zipedFolder, null);
            }
            catch (Exception ex)
            {
                errorOut = ex.Message;
            }
            return result;
        }
        /// <summary>     
        /// 解压功能(解压压缩文件到指定目录)---->需要密码  
        /// </summary>     
        /// <param name="fileToUnZip">待解压的文件</param>     
        /// <param name="zipedFolder">指定解压目标目录</param>  
        /// <param name="password">密码</param>   
        /// <param name="errorOut">如果失败返回失败信息</param>   
        /// <returns>解压结果</returns>  
        public static bool UPZipFile(string fileToUnZip, string zipedFolder, string password, ref string errorOut)
        {
            bool result = false;
            try
            {
                result = UPZipFileByPassword(fileToUnZip, zipedFolder, password);
            }
            catch (Exception ex)
            {
                errorOut = ex.Message;
            }

            return result;
        }

        /// <summary>  
        /// 解压功能 内部处理方法  
        /// </summary>  
        /// <param name="TargetFile">待解压的文件</param>  
        /// <param name="fileDir">指定解压目标目录</param>  
        /// <param name="password">密码</param>  
        /// <returns>成功返回true</returns>  
        private static bool UPZipFileByPassword(string TargetFile, string fileDir, string password)
        {
            bool rootFile = true;
            try
            {
                //读取压缩文件(zip文件)，准备解压缩  
                ZipInputStream zipStream = new ZipInputStream(File.OpenRead(TargetFile.Trim()));
                ZipEntry theEntry;
                string path = fileDir;

                string rootDir = " ";
                if (!string.IsNullOrEmpty(password)) zipStream.Password = password;
                List<string> ddddd = new List<string>();
                while ((theEntry = zipStream.GetNextEntry()) != null)
                {
                    ddddd.Add(theEntry.Name);
                    rootDir = Path.GetDirectoryName(theEntry.Name);
                    if (rootDir.IndexOf("\\") >= 0)
                    {
                        rootDir = rootDir.Substring(0, rootDir.IndexOf("\\") + 1);
                    }
                    string dir = Path.GetDirectoryName(theEntry.Name);
                    string fileName = Path.GetFileName(theEntry.Name);
                    if (dir != " ")
                    {
                        path = fileDir + "\\" + dir;
                        if (!Directory.Exists(fileDir + "\\" + dir))
                        {
                            Directory.CreateDirectory(path);
                        }
                    }
                    else if (dir == " " && fileName != "")
                    {
                        path = fileDir;
                    }
                    else if (dir != " " && fileName != "")
                    {
                        if (dir.IndexOf("\\") > 0)
                        {
                            path = fileDir + "\\" + dir;
                        }
                    }

                    if (dir == rootDir)
                    {
                        path = fileDir + "\\" + rootDir;
                    }

                    //以下为解压缩zip文件的基本步骤  
                    //基本思路就是遍历压缩文件里的所有文件，创建一个相同的文件。  
                    if (fileName != String.Empty)
                    {
                        FileStream streamWriter = File.Create(path + "\\" + fileName);

                        int size = 2048;
                        byte[] data = new byte[2048];
                        while (true)
                        {
                            size = zipStream.Read(data, 0, data.Length);
                            if (size > 0)
                            {
                                streamWriter.Write(data, 0, size);
                            }
                            else
                            {
                                break;
                            }
                        }

                        streamWriter.Close();
                    }
                }
                int cou = ddddd.Count;
                if (theEntry != null)
                {
                    theEntry = null;
                }
                if (zipStream != null)
                {
                    zipStream.Close();
                }
            }
            catch (Exception ex)
            {
                rootFile = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
                GC.Collect(1);
            }
            return rootFile;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="base64string"></param>
        /// <returns></returns>
        public Bitmap GetImageFromBase64(string base64string)
        {
            byte[] b = Convert.FromBase64String(base64string);
            MemoryStream ms = new MemoryStream(b);
            Bitmap bitmap = new Bitmap(ms);
            return bitmap;
        }
        public static string GetBase64FromImage(string imagefile)
        {
            string strbaser64 = "";
            try
            {
                using (Bitmap bmp = new Bitmap(imagefile))
                {
                    Graphics graphics = Graphics.FromImage(bmp);
                    MemoryStream ms = new MemoryStream();
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    byte[] arr = new byte[ms.Length];
                    ms.Position = 0;
                    ms.Read(arr, 0, (int)ms.Length);
                    ms.Close();
                    strbaser64 = Convert.ToBase64String(arr);
                }

            }
            catch (Exception)
            {
                throw new Exception("Something wrong during convert!");
            }
            return strbaser64;
        }
    }
    /// <summary>
    /// 文件夹操作类
    /// </summary>
    public static class FolderHelper
    {
        /// <summary>
        /// 复制文件夹
        /// </summary>
        /// <param name="sourceFolderName">源文件夹目录</param>
        /// <param name="destFolderName">目标文件夹目录</param>
        public static void Copy(string sourceFolderName, string destFolderName)
        {
            Copy(sourceFolderName, destFolderName, false);
        }

        /// <summary>
        /// 复制文件夹
        /// </summary>
        /// <param name="sourceFolderName">源文件夹目录</param>
        /// <param name="destFolderName">目标文件夹目录</param>
        /// <param name="overwrite">允许覆盖文件</param>
        public static void Copy(string sourceFolderName, string destFolderName, bool overwrite)
        {
            var sourceFilesPath = Directory.GetFileSystemEntries(sourceFolderName);

            for (int i = 0; i < sourceFilesPath.Length; i++)
            {
                var sourceFilePath = sourceFilesPath[i];
                var directoryName = Path.GetDirectoryName(sourceFilePath);
                var forlders = directoryName.Split('\\');
                var lastDirectory = forlders[forlders.Length - 1];
                var dest = Path.Combine(destFolderName, lastDirectory);
                //var dest = destFolderName;
                if (File.Exists(sourceFilePath))
                {
                    var sourceFileName = Path.GetFileName(sourceFilePath);
                    if (!Directory.Exists(dest))
                    {
                        Directory.CreateDirectory(dest);
                    }
                    File.Copy(sourceFilePath, Path.Combine(dest, sourceFileName), overwrite);
                }
                else
                {
                    Copy(sourceFilePath, dest, overwrite);
                }
            }
        }
    }
    /// <summary>
    /// 缩略图
    /// </summary>
    public class ThumbnailMaker
    {
        //private static ThumbnailMaker instance;
        //public static ThumbnailMaker Instance
        //{
        //    get
        //    {
        //        lock (typeof(ThumbnailMaker))
        //        {
        //            if (instance == null)
        //                instance = new ThumbnailMaker();
        //            return instance;
        //        }
        //    }
        //}


        public static void CreateDirectory(string path)
        {
            if (path == "") return;
            string head = path.Substring(0, path.IndexOf("\\"));  //d:
            string weibu = path.Substring(head.Length + 1);  //       \1\2
            string hpath = head;
            while (weibu.IndexOf("\\") != -1)
            {
                string p = hpath + "\\" + weibu.Substring(0, weibu.IndexOf("\\"));
                hpath = p;
                if (!Directory.Exists(p))
                    Directory.CreateDirectory(p);
                int ix = weibu.IndexOf("\\") + 1;
                weibu = weibu.Substring(ix);
            }

        }

        /// <summary>
        /// 制作图片的缩略图
        /// </summary>
        /// <param name="originalImage">原图</param>
        /// <param name="width">缩略图的宽（像素）</param>
        /// <param name="height">缩略图的高（像素）</param>
        /// <param name="mode">缩略方式</param>
        /// <returns>缩略图</returns>
        /// <remarks>
        ///        <paramref name="mode"/>：
        ///            <para>HW：指定的高宽缩放（可能变形）</para>
        ///            <para>HWO：指定高宽缩放（可能变形）（过小则不变）</para>
        ///            <para>W：指定宽，高按比例</para>
        ///            <para>WO：指定宽（过小则不变），高按比例</para>
        ///            <para>H：指定高，宽按比例</para>
        ///            <para>HO：指定高（过小则不变），宽按比例</para>
        ///            <para>CUT：指定高宽裁减（不变形）</para>
        /// </remarks>
        public static Image MakeThumbnail(Image originalImage, int width, int height, ThumbnailMode mode)
        {
            int towidth = width;
            int toheight = height;

            int x = 0;
            int y = 0;
            int ow = originalImage.Width;
            int oh = originalImage.Height;


            switch (mode)
            {
                case ThumbnailMode.UsrHeightWidth: //指定高宽缩放（可能变形）
                    break;
                case ThumbnailMode.UsrHeightWidthBound: //指定高宽缩放（可能变形）（过小则不变）
                    if (originalImage.Width <= width && originalImage.Height <= height)
                    {
                        return originalImage;
                    }
                    if (originalImage.Width < width)
                    {
                        towidth = originalImage.Width;
                    }
                    if (originalImage.Height < height)
                    {
                        toheight = originalImage.Height;
                    }
                    break;
                case ThumbnailMode.UsrWidth: //指定宽，高按比例
                    toheight = originalImage.Height * width / originalImage.Width;
                    break;
                case ThumbnailMode.UsrWidthBound: //指定宽（过小则不变），高按比例
                    if (originalImage.Width <= width)
                    {
                        return originalImage;
                    }
                    else
                    {
                        toheight = originalImage.Height * width / originalImage.Width;
                    }
                    break;
                case ThumbnailMode.UsrHeight: //指定高，宽按比例
                    towidth = originalImage.Width * height / originalImage.Height;
                    break;
                case ThumbnailMode.UsrHeightBound: //指定高（过小则不变），宽按比例
                    if (originalImage.Height <= height)
                    {
                        return originalImage;
                    }
                    else
                    {
                        towidth = originalImage.Width * height / originalImage.Height;
                    }
                    break;
                case ThumbnailMode.Cut: //指定高宽裁减（不变形）
                    if ((double)originalImage.Width / (double)originalImage.Height > (double)towidth / (double)toheight)
                    {
                        oh = originalImage.Height;
                        ow = originalImage.Height * towidth / toheight;
                        y = 0;
                        x = (originalImage.Width - ow) / 2;
                    }
                    else
                    {
                        ow = originalImage.Width;
                        oh = originalImage.Width * height / towidth;
                        x = 0;
                        y = (originalImage.Height - oh) / 2;
                    }
                    break;
                default:
                    break;
            }

            //新建一个bmp图片
            Image bitmap = new Bitmap(towidth, toheight);

            //新建一个画板
            Graphics g = Graphics.FromImage(bitmap);
            Pen pen = new Pen(Color.Black, 1);

            //设置高质量插值法
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            //设置高质量,低速度呈现平滑程度
            g.SmoothingMode = SmoothingMode.AntiAlias;

            //清空画布并以透明背景色填充
            g.Clear(Color.Transparent);

            //在指定位置并且按指定大小绘制原图片的指定部分
            g.DrawImage(originalImage, new Rectangle(0, 0, towidth, toheight),
                        new Rectangle(x, y, ow, oh),
                        GraphicsUnit.Pixel);
            g.Dispose();
            return bitmap;
        }

        /// <summary>
        /// 制作图片的缩略图
        /// </summary>
        /// <param name="originalStream">原图</param>
        /// <param name="thumbnailPath">保存缩略图的路径</param>
        /// <param name="width">缩略图的宽（像素）</param>
        /// <param name="height">缩略图的高（像素）</param>
        /// <param name="mode">缩略方式，参见<seealso cref="MakeThumbnail(Image, int, int, string)"/></param>
        public static void MakeThumbnail(Stream originalStream, string thumbnailPath, int width, int height, ThumbnailMode mode)
        {
            Image originalImage = Image.FromStream(originalStream);
            try
            {
                MakeThumbnail(originalImage, thumbnailPath, width, height, mode);
            }
            finally
            {
                originalImage.Dispose();
            }
        }

        /// <summary>
        /// 制作图片的缩略图
        /// </summary>
        /// <param name="originalImage">原图</param>
        /// <param name="thumbnailPath">保存缩略图的路径</param>
        /// <param name="width">缩略图的宽（像素）</param>
        /// <param name="height">缩略图的高（像素）</param>
        /// <param name="mode">缩略方式，参见<seealso cref="MakeThumbnail(Image, int, int, string)"/></param>
        public static void MakeThumbnail(Image originalImage, string thumbnailPath, int width, int height, ThumbnailMode mode)
        {
            Image bitmap = MakeThumbnail(originalImage, width, height, mode);
            try
            {
                //以jpg格式保存缩略图
                bitmap.Save(thumbnailPath, ImageFormat.Png);
                
            }
            finally
            {
                bitmap.Dispose();
                
            }
        }

        /// <summary>
        /// 制作图片的缩略图
        /// </summary>
        /// <param name="originalImagePath">原图的路径</param>
        /// <param name="thumbnailPath">保存缩略图的路径</param>
        /// <param name="width">缩略图的宽（像素）</param>
        /// <param name="height">缩略图的高（像素）</param>
        /// <param name="mode">缩略方式，参见<seealso cref="MakeThumbnail(Image, int, int, string)"/></param>
        public static void MakeThumbnail(string originalImagePath, string thumbnailPath, int width, int height, ThumbnailMode mode)
        {
            Image originalImage = Image.FromFile(originalImagePath);
            try
            {
                MakeThumbnail(originalImage, thumbnailPath, width, height, mode);
            }
            finally
            {
                originalImage.Dispose();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="img"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="mode"></param>
        public static byte[] MakeThumbnail(byte[] img, int width, int height, ThumbnailMode mode)
        {
            Stream byteStream = new MemoryStream(img);

            Image originalImage = Image.FromStream(byteStream);

            try
            {
                Image _img = MakeThumbnail(originalImage, width, height, mode);

                MemoryStream imgStream = new MemoryStream();

                _img.Save(imgStream, ImageFormat.Jpeg);

                return imgStream.ToArray();
            }
            finally
            {
                originalImage.Dispose();
            }
        }
        /// <summary>
        /// 产生高清缩略图 固定大小
        /// </summary>
        /// <param name="original_image_file">源文件</param>
        /// <param name="object_width">缩略图宽度</param>
        /// <param name="object_height">缩略图高度</param>
        public static void MakeHighQualityThumbnail(string original_image_file, string output, int object_width, int object_height)
        {


            int actual_width = 0;
            int actual_heigh = 0;
            string outputfilename = output; //original_image_file + ".jpg";

            System.Drawing.Bitmap original_image = new Bitmap(original_image_file);//读取源文件           
            actual_width = original_image.Width;
            actual_heigh = original_image.Height;

            Bitmap img = new Bitmap(object_width, object_height);
            img.SetResolution(108f, 108f);
            Graphics gdiobj = Graphics.FromImage(img);
            gdiobj.CompositingQuality = CompositingQuality.HighQuality;
            gdiobj.SmoothingMode = SmoothingMode.HighQuality;
            gdiobj.InterpolationMode = InterpolationMode.HighQualityBicubic;
            gdiobj.PixelOffsetMode = PixelOffsetMode.HighQuality;

            gdiobj.FillRectangle(new SolidBrush(Color.White), 0, 0, object_width, object_height);
            Rectangle destrect = new Rectangle(0, 0, object_width, object_height);

            gdiobj.DrawImage(original_image, destrect, 0, 0, actual_width, actual_heigh, GraphicsUnit.Pixel);

            System.Drawing.Imaging.EncoderParameters ep = new System.Drawing.Imaging.EncoderParameters(1);
            ep.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)100);

            System.Drawing.Imaging.ImageCodecInfo ici = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()[0];

            if (ici != null)
            {
                if (File.Exists(outputfilename))
                    File.Delete(outputfilename);
                img.Save(outputfilename, ici, ep);

            }
            else
            {
                img.Save(outputfilename, System.Drawing.Imaging.ImageFormat.Jpeg);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="img"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="mode"></param>
        public static byte[] MakeThumbnail(byte[] img, int width, int height)
        {
            return MakeThumbnail(img, width, height, ThumbnailMode.UsrHeightBound);
        }
    }



    //<para>HW：指定的高宽缩放（可能变形）</para>
    //<para>HWO：指定高宽缩放（可能变形）（过小则不变）</para>
    //<para>W：指定宽，高按比例</para>
    //<para>WO：指定宽（过小则不变），高按比例</para>
    //<para>H：指定高，宽按比例</para>
    //<para>HO：指定高（过小则不变），宽按比例</para>
    //<para>CUT：指定高宽裁减（不变形）</para>
    public enum ThumbnailMode
    {
        UsrHeightWidth,
        UsrHeightWidthBound,
        UsrWidth,
        UsrWidthBound,
        UsrHeight,
        UsrHeightBound,
        Cut,
        NONE,
    }
}