using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
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
            if (listView1.Items.Count == 0)
            {
                MessageBox.Show("请先载入发票文件！", "提示");
                return;
            }

            file2image();

            // 构建平面图片列表（按重复次数展开）
            List<Image> flatImages = new List<Image>();
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                string path = listView1.Items[i].Text;
                int repeat = (int)numericUpDown1.Value;
                for (int j = 0; j < repeat; j++)
                {
                    flatImages.Add(invoiceList[path]);
                }
            }

            PrintDocument pd = new PrintDocument();
            // 从打印机支持的纸张中选取 A4
            var a4Paper = pd.PrinterSettings.PaperSizes
                .OfType<PaperSize>()
                .FirstOrDefault(p => p.Kind == PaperKind.A4);
            if (a4Paper != null)
                pd.DefaultPageSettings.PaperSize = a4Paper;
            pd.DefaultPageSettings.Margins = new Margins(39, 39, 39, 39); // ~10mm

            int index = 0;
            pd.PrintPage += (s, ev) =>
            {
                float y = ev.MarginBounds.Top;

                while (index < flatImages.Count)
                {
                    Image img = flatImages[index];

                    // 保持宽高比，宽度适应页面可用宽度
                    float scale = (float)ev.MarginBounds.Width / img.Width;
                    float imgWidth = ev.MarginBounds.Width;
                    float imgHeight = img.Height * scale;

                    // 如果图片超出页面剩余高度，换页
                    if (y + imgHeight > ev.MarginBounds.Bottom)
                    {
                        ev.HasMorePages = true;
                        return;
                    }

                    ev.Graphics.DrawImage(img, ev.MarginBounds.Left, y, imgWidth, imgHeight);
                    y += imgHeight;
                    index++;
                }

                ev.HasMorePages = false;
            };

            // 自定义预览窗体（PrintPreviewDialog 的打印按钮行为不可靠，自己做）
            Form previewForm = new Form();
            previewForm.Text = "打印预览";
            previewForm.WindowState = FormWindowState.Maximized;
            previewForm.Icon = this.Icon;
            previewForm.StartPosition = FormStartPosition.CenterParent;

            // 预览控件（先创建，后面缩放控制和滚轮事件需要引用它）
            PrintPreviewControl ppc = new PrintPreviewControl();
            ppc.Document = pd;
            ppc.Dock = DockStyle.Fill;
            ppc.Zoom = 0.3;

            // 工具栏
            Panel toolbar = new Panel();
            toolbar.Height = 40;
            toolbar.Dock = DockStyle.Top;
            toolbar.BackColor = SystemColors.Control;

            Button printBtn = new Button();
            printBtn.Text = "打印";
            printBtn.Location = new Point(12, 8);
            printBtn.Size = new Size(75, 26);
            printBtn.Click += (s, ev) =>
            {
                using (PrintDialog printDlg = new PrintDialog())
                {
                    printDlg.Document = pd;
                    if (printDlg.ShowDialog(previewForm) == DialogResult.OK)
                    {
                        index = 0; // 重置索引，重新渲染页面
                        pd.Print();
                    }
                }
            };
            toolbar.Controls.Add(printBtn);

            Button closeBtn = new Button();
            closeBtn.Text = "关闭";
            closeBtn.Location = new Point(93, 8);
            closeBtn.Size = new Size(75, 26);
            closeBtn.Click += (s, ev) => previewForm.Close();
            toolbar.Controls.Add(closeBtn);

            Button saveWordBtn = new Button();
            saveWordBtn.Text = "保存为 Word";
            saveWordBtn.Location = new Point(175, 8);
            saveWordBtn.Size = new Size(100, 26);
            saveWordBtn.Click += (s, ev) =>
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Word 文档|*.docx";
                    sfd.FileName = "发票.docx";
                    if (sfd.ShowDialog(previewForm) == DialogResult.OK)
                    {
                        file2image();
                        CreatDocx();
                        for (int i = 0; i < listView1.Items.Count; i++)
                        {
                            if (i > 0) NextPage();
                            string path = listView1.Items[i].Text;
                            int repeat = (int)numericUpDown1.Value;
                            for (int j = 0; j < repeat; j++)
                                AddImage(invoiceList[path]);
                        }
                        Save(sfd.FileName);
                        MessageBox.Show("Word 文件已保存！");
                    }
                }
            };
            toolbar.Controls.Add(saveWordBtn);

            // 操作提示（居中显示）
            Label hintLabel = new Label();
            hintLabel.Text = "Ctrl+滚轮缩放 · 滚轮翻页 · 拖动平移";
            hintLabel.AutoSize = true;
            hintLabel.TextAlign = ContentAlignment.MiddleCenter;
            previewForm.Shown += (s, ev) =>
                hintLabel.Left = (toolbar.Width - hintLabel.Width) / 2;
            hintLabel.Top = 12;
            toolbar.Controls.Add(hintLabel);

            // 翻页控制
            int totalPages = 0;
            float pageH = pd.DefaultPageSettings.Bounds.Height
                - pd.DefaultPageSettings.Margins.Top
                - pd.DefaultPageSettings.Margins.Bottom;
            float contentW = pd.DefaultPageSettings.Bounds.Width
                - pd.DefaultPageSettings.Margins.Left
                - pd.DefaultPageSettings.Margins.Right;

            float yPos = 0;
            foreach (var img in flatImages)
            {
                float scale = contentW / img.Width;
                float imgH = img.Height * scale;
                if (yPos + imgH > pageH && yPos > 0)
                {
                    totalPages++;
                    yPos = 0;
                }
                yPos += imgH;
            }
            if (yPos > 0 || totalPages == 0) totalPages++;

            // 先创建翻页控件（必须赋值给变量后才能被事件捕获）
            Button prevBtn = new Button();
            Label pageLabel = new Label();
            Button nextBtn = new Button();

            // 上一页
            prevBtn.Text = "◀ 上一页";
            prevBtn.Location = new Point(330, 8);
            prevBtn.Size = new Size(75, 26);
            prevBtn.Enabled = false;
            prevBtn.Click += (s, ev) =>
            {
                if (ppc.StartPage > 0)
                {
                    ppc.StartPage--;
                    prevBtn.Enabled = ppc.StartPage > 0;
                    nextBtn.Enabled = true;
                    pageLabel.Text = $"第 {ppc.StartPage + 1} 页 / 共 {totalPages} 页";
                }
            };
            toolbar.Controls.Add(prevBtn);

            // 页码
            pageLabel.Text = $"第 1 页 / 共 {totalPages} 页";
            pageLabel.Location = new Point(410, 12);
            pageLabel.Size = new Size(100, 20);
            pageLabel.TextAlign = ContentAlignment.MiddleCenter;
            toolbar.Controls.Add(pageLabel);

            // 下一页
            nextBtn.Text = "下一页 ▶";
            nextBtn.Location = new Point(510, 8);
            nextBtn.Size = new Size(75, 26);
            nextBtn.Enabled = totalPages > 1;
            nextBtn.Click += (s, ev) =>
            {
                if (ppc.StartPage < totalPages - 1)
                {
                    ppc.StartPage++;
                    prevBtn.Enabled = true;
                    nextBtn.Enabled = ppc.StartPage < totalPages - 1;
                    pageLabel.Text = $"第 {ppc.StartPage + 1} 页 / 共 {totalPages} 页";
                }
            };
            toolbar.Controls.Add(nextBtn);

            // 滚轮翻页 + Ctrl+滚轮缩放
            ppc.MouseWheel += (s, ev) =>
            {
                if (ModifierKeys.HasFlag(Keys.Control))
                {
                    // Ctrl+滚轮 → 缩放
                    ppc.Zoom = Math.Max(0.1, Math.Min(3.0, ppc.Zoom + (ev.Delta > 0 ? 0.1 : -0.1)));
                }
                else
                {
                    // 纯滚轮 → 翻页
                    int oldPage = ppc.StartPage;
                    if (ev.Delta > 0 && ppc.StartPage > 0)
                        ppc.StartPage--;
                    else if (ev.Delta < 0 && ppc.StartPage < totalPages - 1)
                        ppc.StartPage++;
                    if (ppc.StartPage != oldPage)
                    {
                        prevBtn.Enabled = ppc.StartPage > 0;
                        nextBtn.Enabled = ppc.StartPage < totalPages - 1;
                        pageLabel.Text = $"第 {ppc.StartPage + 1} 页 / 共 {totalPages} 页";
                    }
                }
            };

            // 鼠标拖动平移（缩放后内容超出预览区时可用）
            Point dragStart = Point.Empty;
            int scrollAccY = 0, scrollAccX = 0;
            ppc.MouseDown += (s, ev) =>
            {
                if (ev.Button == MouseButtons.Left || ev.Button == MouseButtons.Right)
                {
                    dragStart = ev.Location;
                    scrollAccY = 0;
                    scrollAccX = 0;
                    ppc.Cursor = Cursors.SizeAll;
                }
            };
            ppc.MouseUp += (s, ev) => { dragStart = Point.Empty; ppc.Cursor = Cursors.Default; };
            ppc.MouseMove += (s, ev) =>
            {
                if (dragStart == Point.Empty) return;
                int dy = ev.Y - dragStart.Y;
                int dx = ev.X - dragStart.X;

                // 垂直拖动：累加鼠标移动距离，达到阈值触发一次滚动
                scrollAccY += dy;
                while (scrollAccY >= 4)
                {
                    Win32SendMessage(ppc.Handle, WM_VSCROLL, (IntPtr)SB_LINEUP, IntPtr.Zero);
                    scrollAccY -= 4;
                }
                while (scrollAccY <= -4)
                {
                    Win32SendMessage(ppc.Handle, WM_VSCROLL, (IntPtr)SB_LINEDOWN, IntPtr.Zero);
                    scrollAccY += 4;
                }

                // 水平拖动
                scrollAccX += dx;
                while (scrollAccX >= 4)
                {
                    Win32SendMessage(ppc.Handle, WM_HSCROLL, (IntPtr)SB_LINEUP, IntPtr.Zero);
                    scrollAccX -= 4;
                }
                while (scrollAccX <= -4)
                {
                    Win32SendMessage(ppc.Handle, WM_HSCROLL, (IntPtr)SB_LINEDOWN, IntPtr.Zero);
                    scrollAccX += 4;
                }

                dragStart = ev.Location;
            };

            previewForm.Controls.Add(ppc);
            previewForm.Controls.Add(toolbar);
            previewForm.ShowDialog(this);
            previewForm.Dispose();
        }

        private void saveWord_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count == 0)
            {
                MessageBox.Show("请先载入发票文件！");
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Word 文档|*.docx";
                sfd.FileName = "发票.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    file2image();
                    CreatDocx();
                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        if (i > 0) NextPage();
                        string path = listView1.Items[i].Text;
                        int repeat = (int)numericUpDown1.Value;
                        for (int j = 0; j < repeat; j++)
                            AddImage(invoiceList[path]);
                    }
                    Save(sfd.FileName);
                    MessageBox.Show("Word 文件已保存！", "提示");
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

        #region P/Invoke: 用于鼠标拖动平移

        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = CharSet.Auto)]
        private static extern IntPtr Win32SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool GetScrollInfo(IntPtr hWnd, int nBar, ref SCROLLINFO lpsi);

        [DllImport("user32.dll")]
        private static extern int SetScrollInfo(IntPtr hWnd, int nBar, ref SCROLLINFO lpsi, bool redraw);

        [StructLayout(LayoutKind.Sequential)]
        private struct SCROLLINFO
        {
            public int cbSize;
            public int fMask;
            public int nMin;
            public int nMax;
            public int nPage;
            public int nPos;
            public int nTrackPos;
        }

        private const int WM_VSCROLL = 0x115;
        private const int WM_HSCROLL = 0x114;
        private const int SB_THUMBTRACK = 5;
        private const int SIF_ALL = 0x0017;
        private const int SB_LINEUP = 0;
        private const int SB_LINEDOWN = 1;

        #endregion
    }
}
