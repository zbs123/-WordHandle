
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Dml.ChartDrawing;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;

namespace eyuan
{
    public static class NOPIHandler
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static List<List<List<string>>> ReadExcel(string fileName)
        {
            //打开Excel工作簿
            XSSFWorkbook hssfworkbook = null;
            try
            {
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new XSSFWorkbook(file);
                }
            }
            catch (Exception e)
            {
                //LogHandler.LogWrite(string.Format("文件{0}打开失败，错误：{1}", new string[] { fileName, e.ToString() }));
            }
            //循环Sheet页
            int sheetsCount = hssfworkbook.NumberOfSheets;
            List<List<List<string>>> workBookContent = new List<List<List<string>>>();
            for (int i = 0; i < sheetsCount; i++)
            {
                //Sheet索引从0开始
                ISheet sheet = hssfworkbook.GetSheetAt(i);
                //循环行
                List<List<string>> sheetContent = new List<List<string>>();
                int rowCount = sheet.PhysicalNumberOfRows;
                for (int j = 0; j < rowCount; j++)
                {
                    //Row（逻辑行）的索引从0开始
                    IRow row = sheet.GetRow(j);
                    //循环列（各行的列数可能不同）
                    List<string> rowContent = new List<string>();
                    int cellCount = row.PhysicalNumberOfCells;
                    for (int k = 0; k < cellCount; k++)
                    {
                        //ICell cell = row.GetCell(k);
                        NPOI.SS.UserModel.ICell cell = row.Cells[k];
                        if (cell == null)
                        {
                            rowContent.Add("NIL");
                        }
                        else
                        {
                            rowContent.Add(cell.ToString());
                            //rowContent.Add(cell.StringCellValue);
                        }
                    }
                    //添加行到集合中
                    sheetContent.Add(rowContent);
                }
                //添加Sheet到集合中
                workBookContent.Add(sheetContent);
            }

            return workBookContent;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string ReadExcelText(string fileName)
        {
            string ExcelCellSeparator = ConfigurationManager.AppSettings["ExcelCellSeparator"];
            string ExcelRowSeparator = ConfigurationManager.AppSettings["ExcelRowSeparator"];
            string ExcelSheetSeparator = ConfigurationManager.AppSettings["ExcelSheetSeparator"];
            //
            List<List<List<string>>> excelContent = ReadExcel(fileName);
            string fileText = string.Empty;
            StringBuilder sbFileText = new StringBuilder();
            //循环处理WorkBook中的各Sheet页
            List<List<List<string>>>.Enumerator enumeratorWorkBook = excelContent.GetEnumerator();
            while (enumeratorWorkBook.MoveNext())
            {

                //循环处理当期Sheet页中的各行
                List<List<string>>.Enumerator enumeratorSheet = enumeratorWorkBook.Current.GetEnumerator();
                while (enumeratorSheet.MoveNext())
                {

                    string[] rowContent = enumeratorSheet.Current.ToArray();
                    sbFileText.Append(string.Join(ExcelCellSeparator, rowContent));
                    sbFileText.Append(ExcelRowSeparator);
                }
                sbFileText.Append(ExcelSheetSeparator);
            }
            //
            fileText = sbFileText.ToString();
            return fileText;
        }

        /// <summary>
        /// 读取Word内容
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string ReadWordText(string fileName)
        {
            string[] myPathArr = fileName.Split('\\');
            string myPash = string.Join("\\", myPathArr, 0, myPathArr.Length - 1);
            string WordTableCellSeparator = ConfigurationManager.AppSettings["WordTableCellSeparator"];
            string WordTableRowSeparator = ConfigurationManager.AppSettings["WordTableRowSeparator"];
            string WordTableSeparator = ConfigurationManager.AppSettings["WordTableSeparator"];
            //
            string CaptureWordHeader = ConfigurationManager.AppSettings["CaptureWordHeader"];
            string CaptureWordFooter = ConfigurationManager.AppSettings["CaptureWordFooter"];
            string CaptureWordTable = ConfigurationManager.AppSettings["CaptureWordTable"];
            string CaptureWordImage = ConfigurationManager.AppSettings["CaptureWordImage"];
            //
            string CaptureWordImageFileName = ConfigurationManager.AppSettings["CaptureWordImageFileName"];
            //
            string fileText = string.Empty;
            StringBuilder sbFileText = new StringBuilder();

            #region 打开文档
            XWPFDocument document = null;
            try
            {
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    document = new XWPFDocument(file);
                    
                    IList<XWPFPictureData> picList = document.AllPictures;
                    foreach (XWPFPictureData pic in picList)
                    {
                        Console.WriteLine(pic.GetPictureType() + pic.SuggestFileExtension()
                           + pic.FileName + pic.Checksum);
                        using (FileStream sw2 = new FileStream(myPash + "\\" + pic.FileName, FileMode.Create))
                        {
                            byte[] picFileContent = pic.Data;
                            sw2.Write(picFileContent, 0, picFileContent.Length);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                //LogHandler.LogWrite(string.Format("文件{0}打开失败，错误：{1}", new string[] { fileName, e.ToString() }));
            }
            #endregion

            #region 页眉、页脚
            //页眉
            if (CaptureWordHeader == "true")
            {
                sbFileText.AppendLine("Capture Header Begin");
                foreach (XWPFHeader xwpfHeader in document.HeaderList)
                {
                    sbFileText.AppendLine(string.Format("{0}", new string[] { xwpfHeader.Text }));
                }
                sbFileText.AppendLine("Capture Header End");
            }
            //页脚
            if (CaptureWordFooter == "true")
            {
                sbFileText.AppendLine("Capture Footer Begin");
                foreach (XWPFFooter xwpfFooter in document.FooterList)
                {
                    sbFileText.AppendLine(string.Format("{0}", new string[] { xwpfFooter.Text }));
                }
                sbFileText.AppendLine("Capture Footer End");
            }
            #endregion

            #region 表格
            if (CaptureWordTable == "true")
            {
                sbFileText.AppendLine("Capture Table Begin");
                foreach (XWPFTable table in document.Tables)
                {
                    //循环表格行
                    foreach (XWPFTableRow row in table.Rows)
                    {
                        foreach (XWPFTableCell cell in row.GetTableCells())
                        {
                            sbFileText.Append(cell.GetText());
                            //
                            sbFileText.Append(WordTableCellSeparator);
                        }

                        sbFileText.Append(WordTableRowSeparator);
                    }
                    sbFileText.Append(WordTableSeparator);
                }
                sbFileText.AppendLine("Capture Table End");
            }
            #endregion

            #region 图片
            if (CaptureWordImage == "true")
            {
                sbFileText.AppendLine("Capture Image Begin");
                foreach (XWPFPictureData pictureData in document.AllPictures)
                {
                    string picExtName = pictureData.SuggestFileExtension();
                    string picFileName = pictureData.FileName;
                    byte[] picFileContent = pictureData.Data;
                    //
                    string picTempName = string.Format(CaptureWordImageFileName, new string[] { Guid.NewGuid().ToString() + "_" + picFileName + "." + picExtName });
                    //
                    using (FileStream fs = new FileStream(picTempName, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(picFileContent, 0, picFileContent.Length);
                        fs.Close();
                    }
                    //
                    sbFileText.AppendLine(picTempName);
                }
                sbFileText.AppendLine("Capture Image End");
            }
            #endregion

            //正文段落
            #region

            sbFileText.AppendLine("Capture Paragraph Begin");
            bool flag = true;
            XWPFDocument doc = null;

            FileStream sw = null;
            int count = 0;
            foreach (XWPFParagraph paragraph in document.Paragraphs)
            {
                if (flag)
                {
                    doc = new XWPFDocument();
                    sw = new FileStream(myPash + "\\" + count + ".docx", FileMode.Create);
                    //count= paragraph.ParagraphText.Split('.')[0].ToString(); 
                    flag = false;
                }
                XWPFParagraph para = doc.CreateParagraph();

                XWPFRun run = para.CreateRun();
                run.SetText(paragraph.ParagraphText);
                IList<XWPFRun> runs = paragraph.Runs;
                foreach (XWPFRun r in runs)
                {
                    //string aaa = r.ToString();
                    int i = r.GetEmbeddedPictures().Count;
                    if (i > 0)
                    {
                        string ajdfso = r.GetEmbeddedPictures()[0].GetPictureData().FileName;

                        using (FileStream fs = new FileStream(myPash + "\\" + ajdfso, FileMode.Open, FileAccess.Read))
                        {
                            //System.Drawing.Image image = System.Drawing.Image.FromStream(fs);
                            run.AddPicture(fs, (int)NPOI.XWPF.UserModel.PictureType.PNG, "asd",(int)(400.0 * 9525), (int)(300.0 * 9525));
                            
                        }


                    }
                }
                if (paragraph.ParagraphText.IndexOf("【类型】") != -1)
                {
                    doc.Write(sw);
                    sw.Close();
                    count++;
                    flag = true;
                }

                //XWPFPicture pict = r[0].GetEmbeddedPictures()[0];
                sbFileText.AppendLine(paragraph.ParagraphText);

            }
            #endregion
            #region
            //sbFileText.AppendLine("Capture Paragraph Begin");
            //bool flag = true;
            //XWPFDocument doc = null;
            //FileStream sw = null;
            //int count = 0;
            //string TH = "";
            //foreach (XWPFParagraph paragraph in document.Paragraphs)
            //{ 
            //    if (flag)
            //    {
            //        doc = new XWPFDocument();

            //        //count= paragraph.ParagraphText.Split('.')[0].ToString(); 
            //        flag = false;
            //    }
            //    if (paragraph.ParagraphText.IndexOf("【题号】") != -1)
            //    {


            //        //XWPFRun run = para.CreateRun();

            //        //读答案
            //        string th = paragraph.ParagraphText.Substring(4);
            //        TH = th;
            //        if (th.IndexOf("-") != -1)
            //        {
            //            XWPFParagraph para = doc.CreateParagraph();
            //            string[] ths = th.Split('-');
            //            XWPFRun run = para.CreateRun();
            //            run.SetText("【答案】");
            //            for (var i = int.Parse(ths[0]); i <= int.Parse(ths[1]); i++)
            //            {
            //                Dictionary<string, string> dic = ReadAnswer(@"C:\Users\admin\Desktop\英语\2018届高三第五学期10月月考英语答案.docx", i);
            //                IList<string> list = new List<string>();
            //                if (dic["answer"].Length > 0)
            //                {
            //                    list.Add("(" + i+ ")" + dic["answer"]);
            //                }
            //                for (var j = 0; j < list.Count; j++)
            //                {
            //                    XWPFParagraph para1 = doc.CreateParagraph();
            //                    XWPFRun run1 = para1.CreateRun();
            //                    run1.SetText(list[j]);
            //                }
            //            }
            //            IList<string> list1 = new List<string>();
            //            for (var i = int.Parse(ths[0]); i <= int.Parse(ths[1]); i++)
            //            {
            //                Dictionary<string, string> dic = ReadAnswer(@"C:\Users\admin\Desktop\英语\2018届高三第五学期10月月考英语答案.docx", i);

            //                if (dic["parsing"].Length > 0)
            //                {
            //                    if (list1.Count == 0)
            //                    {
            //                        list1.Add("【解析】");
            //                    }
            //                    list1.Add("(" + i + ")" + dic["parsing"]);
            //                }
            //                for (var j = 0; j < list1.Count; j++)
            //                {
            //                    XWPFParagraph para2 = doc.CreateParagraph();
            //                    XWPFRun run2 = para2.CreateRun();
            //                    run2.SetText(list1[j]);
            //                }
            //            }
            //        }
            //        else
            //        {
            //            Dictionary<string, string> dic = ReadAnswer(@"C:\Users\admin\Desktop\英语\2018届高三第五学期10月月考英语答案.docx", int.Parse(th));
            //            IList<string> list = new List<string>();
            //            list.Add("【答案】");
            //            if (dic["answer"].Length > 0)
            //            {
            //                list.Add(dic["answer"]);
            //            }
            //            if (dic["parsing"].Length > 0)
            //            {
            //                list.Add("【解析】");
            //                list.Add(dic["parsing"]);
            //            }
            //            for (var i = 0; i < list.Count; i++)
            //            {
            //                XWPFParagraph para = doc.CreateParagraph();
            //                XWPFRun run = para.CreateRun();
            //                run.SetText(list[i]);
            //            }

            //        }

            //    }
            //    else
            //    {
            //        XWPFParagraph para = doc.CreateParagraph();

            //        XWPFRun run = para.CreateRun();
            //        run.SetText(paragraph.ParagraphText);
            //        IList<XWPFRun> runs = paragraph.Runs;
            //        foreach (XWPFRun r in runs)
            //        {
            //            //string aaa = r.ToString();
            //            int i = r.GetEmbeddedPictures().Count;
            //            if (i > 0)
            //            {
            //                string ajdfso = r.GetEmbeddedPictures()[0].GetPictureData().FileName;
            //                using (FileStream fs = new FileStream(myPash + "\\" + ajdfso, FileMode.Open, FileAccess.Read))
            //                {
            //                    run.AddPicture(fs, (int)NPOI.XWPF.UserModel.PictureType.PNG, "asd", (int)(400.0 * 9525), (int)(300.0 * 9525));
            //                }
            //            }
            //        }
            //    }

            //    if (paragraph.ParagraphText.IndexOf("【类型】") != -1)
            //    {
            //        if (TH.Length == 0)
            //        {
            //            sw = new FileStream(myPash + "\\" + count + ".docx", FileMode.Create);

            //        }
            //        else
            //        {
            //            sw = new FileStream(myPash + "\\" + TH + ".docx", FileMode.Create);

            //        }
            //        doc.Write(sw);
            //        sw.Close();
            //        TH = "";
            //        count++;
            //        flag = true;
            //    }

            //    //XWPFPicture pict = r[0].GetEmbeddedPictures()[0];
            //    sbFileText.AppendLine(paragraph.ParagraphText);

            //}
#endregion
            fileText = sbFileText.ToString();
            return fileText;
        }

        public static void testSimpleWrite()
        {
            //新建一个文档  
            XWPFDocument doc = new XWPFDocument();
            //创建一个段落  
            XWPFParagraph para = doc.CreateParagraph();

            //一个XWPFRun代表具有相同属性的一个区域。  
            XWPFRun run = para.CreateRun();

            run.SetText("");
            using (FileStream sw = new FileStream("", FileMode.Create))
            {
                doc.Write(sw);
            }
        }
        public static Dictionary<string, string> ReadAnswer(string fileName, int n)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            XWPFDocument answerDoc = null;

            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                answerDoc = new XWPFDocument(file);
            }
            string answer = "";
            string parsing = "";
            bool answerFlag = false;
            bool parsingFlag = false;
            foreach (XWPFParagraph paragraph in answerDoc.Paragraphs)
            {
                if (paragraph.ParagraphText.IndexOf(n + ".【答案】") != -1)
                {
                    answerFlag = true;
                    answer += paragraph.ParagraphText.Substring(6);
                    continue;
                }
                if (answerFlag)
                {
                    if (paragraph.ParagraphText.IndexOf("【答案】") != -1)
                    {
                        answerFlag = false;
                        parsingFlag = false;
                        break;
                    }
                    if (paragraph.ParagraphText.IndexOf("【解析】") != -1)
                    {
                        parsingFlag = true;
                        answerFlag = false;
                    }
                    else
                    {
                        answer += paragraph.ParagraphText;

                    }

                }
                if (parsingFlag)
                {
                    parsing += paragraph.ParagraphText;
                }

                //XWPFPicture pict = r[0].GetEmbeddedPictures()[0];

            }
            dic.Add("answer", answer);
            dic.Add("parsing", parsing);
            return dic;
        }
        public static void WriteWord()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph para = doc.CreateParagraph();

            XWPFRun run = para.CreateRun();
            run.SetText("aaaa");
            FileStream sw = new FileStream(@"D:\工作夹\刘涛\新建文件夹\1111.docx", FileMode.Create);
            doc.Write(sw);
            sw.Close();
        }

    }
}
