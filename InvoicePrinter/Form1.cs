using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
namespace InvoicePrinter
{
    public partial class Form1 : Form
    {
        Dictionary<string, Image> invoiceList = new Dictionary<string, Image>();
        public Form1()
        {
            InitializeComponent();
        }

        private void loadInvoice_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "选择 PDF 文件";
                openFileDialog.Filter = "发票 文件|*.pdf;*.jpg;*.jpeg;*.png;*.bmp|所有文件|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = true; // 是否允许多选

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //return openFileDialog.FileName;
                    foreach (var name in openFileDialog.FileNames)
                    {
                        listView1.Items.Add(name);
                    }

                }
            }
        }
        void file2image()
        {
            invoiceList = new Dictionary<string, Image>();
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                string path = listView1.Items[i].Text;
                string extension = Path.GetExtension(path).ToLower();
                switch (extension)
                {
                    case ".jpg":
                    case ".jpeg":
                    case ".png":
                    case ".bmp":
                        Stream stream = File.OpenRead(path);
                        Image image = Image.FromStream(stream);
                        invoiceList[path] = image;
                        break;
                    case ".pdf":
                        invoiceList[path] = pdf2image(path);
                        break;
                    default:
                        MessageBox.Show(extension);
                        break;
                }

            }
        }
        Image pdf2image(string path)
        {
            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile(path);
            Image image = doc.SaveAsImage(0, PdfImageType.Bitmap,300,300);
            //Image image = Image.FromStream(stream);
            doc.Dispose();
            return image;
        }
        XWPFDocument doc;
        private void exportInvoice_Click(object sender, EventArgs e)
        {
            file2image();
            CreatDocx();
            bool n=false;
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (n)
                {
                    NextPage();
                    n=false;
                }
                string path=listView1.Items[i].Text;
                for (int j = 0; j < (int)numericUpDown1.Value; j++)
                {
                    AddImage(invoiceList[path]);
                }
                if ((int)numericUpDown1.Value % 2 == 1)
                {
                    n=true;
                    
                }
            }
            // 创建 SaveFileDialog 实例
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            // 设置对话框属性
            saveFileDialog.Filter = "Word 文档|*.docx";
            saveFileDialog.FilterIndex = 1; // 默认选择第一个过滤器
            saveFileDialog.Title = "保存文件";
            saveFileDialog.FileName = "发票.docx"; // 默认文件名
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); // 初始目录
            saveFileDialog.RestoreDirectory = true; // 恢复当前目录

            // 显示对话框并处理结果
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取用户选择的文件路径
                string filePath = saveFileDialog.FileName;

                try
                {
                    // 保存文件
                    Save(filePath);
                    MessageBox.Show("文件保存成功！");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"保存文件时出错: {ex.Message}");
                }
            }
        }
        void CreatDocx()
        {
            doc = new XWPFDocument();
            var body = doc.Document.body;
            if (body.sectPr == null)
            {
                body.sectPr = new CT_SectPr();
            }

            // 设置页面尺寸
            if (body.sectPr.pgSz == null)
            {
                body.sectPr.pgSz = new CT_PageSz();
            }
            body.sectPr.pgSz.w = 11906;  // A4宽度 210mm (210 * 56.7)
            body.sectPr.pgSz.h = 16838;  // A4高度 297mm (297 * 56.7)
                                         // 设置页边距
            if (body.sectPr.pgMar == null)
            {
                body.sectPr.pgMar = new CT_PageMar();
            }
            body.sectPr.pgMar.top = 56;    // 上边距 10mm (10 * 56.7)
            body.sectPr.pgMar.bottom = 56; // 下边距 10mm
            body.sectPr.pgMar.left = 56;   // 左边距 10mm
            body.sectPr.pgMar.right = 56;  // 右边距 10mm
        }
        public  void AddImage(Image img)
        {
            XWPFParagraph paragraph1 = doc.CreateParagraph();
            XWPFRun run1 = paragraph1.CreateRun();
            //run1.AddBreak();
            // 创建段落并居中
            var paragraph = doc.CreateParagraph();
            paragraph.Alignment = ParagraphAlignment.CENTER;
            if (img != null)
            {
                // 计算图片尺寸（保持宽高比，宽度填满页面）
                int availableWidth = 210 - 2; // A4宽度减去左右边距（毫米）
                double scaleFactor = (double)availableWidth / img.Width;
                int displayWidth = availableWidth;
                int displayHeight = (int)(img.Height * scaleFactor);

                // 添加图片
                var run = paragraph.CreateRun();
                AddImageToRun(run, img, displayWidth, displayHeight);

            }
        }
        public void NextPage()
        {
            XWPFParagraph breakParagraph = doc.CreateParagraph();
            XWPFRun breakRun = breakParagraph.CreateRun();
            breakRun.AddBreak(BreakType.PAGE);
        }
        public  void Save(string outputPath)
        {
            // 保存文档
            using (var fileStream = new FileStream(outputPath, FileMode.Create))
            {
                doc.Write(fileStream);
            }
        }
        private static void AddImageToRun(XWPFRun run, Image image, int widthMM, int heightMM)
        {
            using (var ms = new MemoryStream())
            {
                // 确保图片保存为支持的格式
                if (image.RawFormat.Equals(System.Drawing.Imaging.ImageFormat.MemoryBmp))
                {
                    image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                }
                else
                {
                    image.Save(ms, image.RawFormat);
                }

                byte[] imageData = ms.ToArray();

                // 添加图片（使用正确的单位转换）
                int widthEMU = MillimetersToEMU(widthMM);
                int heightEMU = MillimetersToEMU(heightMM);

                run.AddPicture(new MemoryStream(imageData), (int)GetPictureType(image.RawFormat),
                             "image", widthEMU, heightEMU);
            }
        }
        private static int MillimetersToEMU(int millimeters)
        {
            // 1毫米 = 36000 EMU
            return millimeters * 36000;
        }

        /// <summary>
        /// 根据图片格式获取NPOI图片类型
        /// </summary>
        private static PictureType GetPictureType(System.Drawing.Imaging.ImageFormat format)
        {
            if (format == System.Drawing.Imaging.ImageFormat.Jpeg)
                return PictureType.JPEG;
            if (format == System.Drawing.Imaging.ImageFormat.Png)
                return PictureType.PNG;
            if (format == System.Drawing.Imaging.ImageFormat.Bmp)
                return PictureType.BMP;
            if (format == System.Drawing.Imaging.ImageFormat.Gif)
                return PictureType.GIF;

            return PictureType.PNG; // 默认使用PNG
        }
    }
}
