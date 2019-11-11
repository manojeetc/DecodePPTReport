using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections;
using System.Drawing.Drawing2D;

namespace DecodingPowerPointReports
{
    public partial class DecodingPPT : Form
    {
        public const string sPowerPointFolder = @"D:\FBR_powerpoint\";
        public const string sPowerPointFolderImgProcessed = @"D:\FBR_powerpoint_processed\";
        public const string sPowerPointFolderImg = @"D:\FBR_powerpoint_img\";
        public const string sPowerPointFolderImgMergedName = @"D:\FBR_powerpoint_img\merged.png";
        public const string sPowerPointPictureExt = @"D:\fbrpics\";
        public const string filePath = @"D:\FBR_powerpoint\powerpoint2.pptx";
        public const string csvfilePath = @"D:\FBR_powerpoint_img\FBRDataSummary.csv";

        public Dictionary<string, string> dicFBRKeyValue = new Dictionary<string, string>();
        public DataTable dtFBR = new DataTable();
        ArrayList listofImages = new ArrayList();

        public DecodingPPT()
        {
            InitializeComponent();
        }
        public class imageInfo
        {
            public int top;
            public int left;
            public int width;
            public int height;
            public int position;
            public string filename;
        }
        public class imageRedraw
        {
            public int x;
            public int y;
            public int width;
            public int height;
            public int position;
            public string filename;
        }
        public Dictionary<int, imageInfo> dicImageFileInfo = new Dictionary<int, imageInfo>();

        private void button1_Click(object sender, EventArgs e)
        {
            initalizeRBRDT();

            string[] filePaths = Directory.GetFiles(sPowerPointFolder);
            foreach (string filePath in filePaths)
                processFBRReport(filePath);
            saveTCSV();
        }
        public void saveTCSV()
        {
            StringBuilder sb = new StringBuilder();

            string[] columnNames = dtFBR.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName).
                                              ToArray();
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dtFBR.Rows)
            {
                string[] fields = row.ItemArray.Select(field => field.ToString()).
                                                ToArray();
                sb.AppendLine(string.Join(",", fields));
            }

            File.WriteAllText(csvfilePath, sb.ToString());
        }
        public void initalizeRBRDT()
        {
            dtFBR.Clear();
            dtFBR.Columns.Add("FILE_NAME", typeof(string));
            dtFBR.Columns.Add("VIN", typeof(string));
            dtFBR.Columns.Add("RANK", typeof(string));
            dtFBR.Columns.Add("TRACKING", typeof(string));
            dtFBR.Columns.Add("TITLE", typeof(string));
            dtFBR.Columns.Add("AF Off Date", typeof(string));
            dtFBR.Columns.Add("Days to Fail", typeof(string));
            dtFBR.Columns.Add("Part Name", typeof(string));
            dtFBR.Columns.Add("Part #", typeof(string));
            dtFBR.Columns.Add("Issuer:", typeof(string));
            dtFBR.Columns.Add("Model:", typeof(string));
            dtFBR.Columns.Add("Year:", typeof(string));
            dtFBR.Columns.Add("Plant:", typeof(string));
            dtFBR.Columns.Add("Dept:", typeof(string));
            dtFBR.Columns.Add("Zone:", typeof(string));
            dtFBR.Columns.Add("Process#", typeof(string));
            dtFBR.Columns.Add("Issued Date:", typeof(string));
            dtFBR.Columns.Add("Remove Date:", typeof(string));
            dtFBR.Columns.Add("Customer Concern:", typeof(string));
            dtFBR.Columns.Add("Dealer Repair:", typeof(string));
            dtFBR.Columns.Add("Additional Details:", typeof(string));
            dtFBR.Columns.Add("CLAIM COST:", typeof(string));
            dtFBR.Columns.Add("ACTUAL CUSTOMER COMPLAINT", typeof(string));
        }
        public void processFBRReport(string fbrFileLocation)
        {
            DataRow dr = dtFBR.NewRow();
            dr["FILE_NAME"] = Path.GetFileName(fbrFileLocation);

            var stringBuilder = new StringBuilder();

            Microsoft.Office.Interop.PowerPoint.Application pptApp =
                                new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentations pptPresentations =
                                                                    pptApp.Presentations;
            Microsoft.Office.Interop.PowerPoint.Presentation pptPresentation =
                                                pptPresentations.Open(fbrFileLocation,
                                                Microsoft.Office.Core.MsoTriState.msoTrue,
                                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoFalse);

            Microsoft.Office.Interop.PowerPoint.Slides pptSlides = pptPresentation.Slides;

            Graphics gr = this.CreateGraphics();

            var slidesCount = pptSlides.Count;
            int imgCrt = 0;

            for (int slideIndex = 1; slideIndex <= slidesCount; slideIndex++)
            {

                var slide = pptSlides[slideIndex];

                foreach (Microsoft.Office.Interop.PowerPoint.Shape textShape in slide.Shapes)
                {

                    if (textShape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                    {

                        listofImages.Add(textShape);

                        imgCrt++;
                        textShape.Export(sPowerPointFolderImg + textShape.ZOrderPosition.ToString()
                                        + "-" + ((textShape.Left * gr.DpiX) / 72).ToString("0.00")
                                         + "-" + ((textShape.Top * gr.DpiX) / 72).ToString("0.00")
                                         + "-" + ((textShape.Width * gr.DpiX) / 72).ToString("0.00")
                                         + "-" + ((textShape.Height * gr.DpiX) / 72).ToString("0.00")
                                         + ".png",
                                         Microsoft.Office.Interop.PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                                         0,
                                         0,
                                         Microsoft.Office.Interop.PowerPoint.PpExportMode.ppScaleXY);

                    }

                    if (textShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue &&
                                   textShape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        Microsoft.Office.Interop.PowerPoint.TextRange pptTextRange = textShape.TextFrame.TextRange;

                        if (pptTextRange != null && pptTextRange.Length > 0)
                        {
                            stringBuilder.Append(" " + pptTextRange.Text);

                            if (pptTextRange.Text.StartsWith("ACTUAL CUSTOMER COMPLAINT") == true)
                            {
                                string tmpString = pptTextRange.Text;
                                tmpString = pptTextRange.Text.Replace("ACTUAL CUSTOMER COMPLAINT", string.Empty);

                                dr["ACTUAL CUSTOMER COMPLAINT"] = tmpString.Replace(",", "").Replace("\r", String.Empty);
                            }
                            else if (pptTextRange.Text.Trim().ToUpper().Equals("MARKET FEED BACK") == true)
                            {

                            }
                            else if (pptTextRange.Text.Trim().ToUpper().StartsWith("RANK:") == true)
                            {
                                string tmpString = pptTextRange.Text.ToUpper();
                                tmpString = pptTextRange.Text.Replace("RANK:", string.Empty);
                                var regex = new Regex(Regex.Escape("\r"));
                                tmpString = regex.Replace(tmpString, "", 1);

                                dr["RANK"] = tmpString;
                            }
                            else if (pptTextRange.Text.Trim().ToUpper().StartsWith("TRACKING #") == true)
                            {
                                string tmpString = pptTextRange.Text.ToUpper();
                                tmpString = tmpString.ToUpper().Replace("TRACKING #:", string.Empty);
                                dr["TRACKING"] = tmpString.Replace(",", "").Replace("\r", String.Empty);
                            }
                            else if (pptTextRange.Text.Trim().StartsWith("This sheet is intended for quick feed back to increase associate") == true)
                            {
                            }
                            else if (pptTextRange.Text.Trim().Equals("For Reference Only") == true)
                            {
                            }
                            else
                            {
                                if (textShape.Name.ToString().Equals("Title 1") == true)
                                {
                                    dr["TITLE"] = pptTextRange.Text.Replace(",", "").Replace("\r", String.Empty);
                                }
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(pptTextRange);
                        }
                    }

                    if (textShape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {

                        if (textShape.Table.Rows.Count > 1)
                        {
                            int iNumRows = textShape.Table.Rows.Count;
                            int iNumCols = textShape.Table.Rows[1].Cells.Count;

                            string sKey = textShape.Table.Rows[1].Cells[1].Shape.TextFrame.TextRange.Text;
                            if ((sKey.Trim().ToUpper().Equals("VIN") == true) ||
                                 (sKey.Trim().ToUpper().Equals("PART NAME") == true) ||
                                 (sKey.Trim().ToUpper().Equals("CUSTOMER CONCERN:") == true) ||
                                 (sKey.Trim().ToUpper().Equals("DEALER REPAIR:") == true) ||
                                 (sKey.Trim().ToUpper().Equals("ADDITIONAL DETAILS:") == true) ||
                                 (sKey.Trim().ToUpper().Equals("CLAIM COST:") == true)
                                )
                            {
                                //Process VIN Block
                                for (int iCol = 1; iCol <= iNumCols; iCol++)
                                {
                                    dr[textShape.Table.Rows[1].Cells[iCol].Shape.TextFrame.TextRange.Text] =
                                        textShape.Table.Rows[2].Cells[iCol].Shape.TextFrame.TextRange.Text.Replace(",", String.Empty).Replace("\r", String.Empty);

                                }
                            }
                            if ((sKey.Trim().ToUpper().Equals("MODEL:") == true) ||
                                (sKey.Trim().ToUpper().Equals("DEPT:") == true) ||
                                (sKey.Trim().ToUpper().Equals("ISSUED DATE:") == true) ||
                                (sKey.Trim().ToUpper().Equals("ISSUER:") == true)
                                )
                            {
                                //Process VIN Block
                                for (int iRow = 1; iRow <= iNumRows; iRow++)
                                {
                                    dr[textShape.Table.Rows[iRow].Cells[1].Shape.TextFrame.TextRange.Text] =
                                        textShape.Table.Rows[iRow].Cells[2].Shape.TextFrame.TextRange.Text.Replace(",", "");
                                }
                            }
                        }
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(textShape);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(slide);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(pptSlides);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pptPresentation);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pptPresentations);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApp);

            dtFBR.Rows.Add(dr);
            string[] filePaths = Directory.GetFiles(sPowerPointFolderImg);

            int min_left = 10000000;
            int min_top = 10000000;
            int max_bottom = 0;
            int max_right = 0;

            foreach (string filePath in filePaths)
            {
                string imagefilename = Path.GetFileNameWithoutExtension(filePath);
                string[] filesplit = imagefilename.Split('-');

                imageInfo imgInfoObj = new imageInfo();
                imgInfoObj.position = Convert.ToInt32(filesplit[0]);
                imgInfoObj.left = Convert.ToInt32(filesplit[1].Substring(0, filesplit[1].IndexOf('.')));
                imgInfoObj.top = Convert.ToInt32(filesplit[2].Substring(0, filesplit[2].IndexOf('.')));
                imgInfoObj.width = Convert.ToInt32(filesplit[3].Substring(0, filesplit[3].IndexOf('.')));
                imgInfoObj.height = Convert.ToInt32(filesplit[4].Substring(0, filesplit[4].IndexOf('.')));
                imgInfoObj.filename = filePath;
                dicImageFileInfo.Add(imgInfoObj.position, imgInfoObj);

                if (imgInfoObj.left < min_left) { min_left = imgInfoObj.left; }
                if (imgInfoObj.top < min_top) { min_top = imgInfoObj.top; }
                if (max_bottom < (imgInfoObj.height + imgInfoObj.top)) { max_bottom = (imgInfoObj.height + imgInfoObj.top); }
                if (max_right < (imgInfoObj.left + imgInfoObj.width)) { max_right = imgInfoObj.left + imgInfoObj.width; }

            }

            List<int> list = dicImageFileInfo.Keys.ToList();
            list.Sort();

            Bitmap target = new Bitmap(max_right - min_left + 50, max_bottom - min_top + 50);

            using (Graphics g = Graphics.FromImage(target))
            {

                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.InterpolationMode = InterpolationMode.NearestNeighbor;

                foreach (var key in list)
                {
                    imageInfo temImgInfo = (imageInfo)dicImageFileInfo[key];
                    Bitmap src = Image.FromFile(temImgInfo.filename) as Bitmap;

                    g.DrawImage(src,
                     (temImgInfo.left - min_left),
                     ((temImgInfo.top - min_top)));
                    src.Dispose();

                }

                g.Dispose();
            }
            string sImageNameCropTemp = sPowerPointFolderImgProcessed + Path.GetFileName(fbrFileLocation) + "_MergedImages.png";
            target.Save(sImageNameCropTemp);
            dicImageFileInfo.Clear();
            target.Dispose();
            gr.Dispose();
            Array.ForEach(Directory.GetFiles(sPowerPointFolderImg), File.Delete);
        }

    }
}
