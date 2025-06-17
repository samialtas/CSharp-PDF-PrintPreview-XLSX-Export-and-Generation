using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
namespace PDF_PrintPreview_XLSX_Export
{
    public partial class MainForm : Form
    {
        private static readonly GraphicsRecorder graphicsRecorder = new GraphicsRecorder();
        private static readonly List<string> pdfContent = new List<string>();
        private static int pdfPageWidth = 612;
        private static int pdfPageHeight = 792;
        private readonly List<DataGridView> dataGridViews = new List<DataGridView>();
        private readonly Margins initialMargins = new Margins(200, 200, 200, 200);
        private int currentDGVIndex = 0;
        private int currentPageNumber = 1;
        private int currentRow = 0;
        private bool printingSettingsSet = false;
        private int totalPageCount = 1;
        private readonly List<string> intermediateCommands = new List<string>();
        private byte[] DevModeArray;
        private bool isExportingToExcel = false;
        public MainForm()
        {
            InitializeComponent();
        }
        #region GraphicsRecorder Class
        public class GraphicsRecorder
        {
            public List<string> Commands { get; private set; } = new List<string>();
            private int currentPage = 1;
            public void SetPage(int page)
            {
                currentPage = page;
            }
            public void ClearCommands()
            {
                Commands.Clear();
            }
            public void DrawRectangle(Pen pen, float x, float y, float width, float height)
            {
                Commands.Add(string.Format(CultureInfo.InvariantCulture, "DrawRectangle|{0}|{1}|{2}|{3}|{4}", x, y, width, height, currentPage));
            }
            public void DrawString(string text, Font font, Brush brush, float x, float y)
            {
                Commands.Add(string.Format(CultureInfo.InvariantCulture, "DrawString|{0}|{1}|{2}|{3}|{4}|{5}|{6}", text, font.Name, font.Size, font.Style, x, y + font.Size, currentPage));
            }
        }
        #endregion
        #region Windows API Declarations
        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);
        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteObject(IntPtr hObject);
        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetFontData(IntPtr hdc, uint dwTable, uint dwOffset, IntPtr lpvBuffer, uint cbData);
        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern uint GetOutlineTextMetrics(IntPtr hdc, uint cbData, IntPtr lpOTM);
        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool GetTextExtentPoint32(IntPtr hdc, string lpString, int cbString, out SIZE lpSize);
        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool GetCharWidth32(IntPtr hdc, uint uFirstChar, uint uLastChar, out int lpBuffer);
        [StructLayout(LayoutKind.Sequential)]
        public struct SIZE
        {
            public int cx;
            public int cy;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct TEXTMETRIC
        {
            public int tmHeight;
            public int tmAscent;
            public int tmDescent;
            public int tmInternalLeading;
            public int tmExternalLeading;
            public int tmAveCharWidth;
            public int tmMaxCharWidth;
            public int tmWeight;
            public int tmOverhang;
            public int tmDigitizedAspectX;
            public int tmDigitizedAspectY;
            public char tmFirstChar;
            public char tmLastChar;
            public char tmDefaultChar;
            public char tmBreakChar;
            public byte tmItalic;
            public byte tmUnderlined;
            public byte tmStruckOut;
            public byte tmPitchAndFamily;
            public byte tmCharSet;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct PANOSE
        {
            public byte bFamilyType;
            public byte bSerifStyle;
            public byte bWeight;
            public byte bProportion;
            public byte bContrast;
            public byte bStrokeVariation;
            public byte bArmStyle;
            public byte bLetterform;
            public byte bMidline;
            public byte bXHeight;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct OUTLINETEXTMETRIC
        {
            public uint otmSize;
            public TEXTMETRIC otmTextMetrics;
            public byte otmFiller;
            public PANOSE otmPanoseNumber;
            public uint otmfsSelection;
            public uint otmfsType;
            public int otmsCharSlopeRise;
            public int otmsCharSlopeRun;
            public int otmItalicAngle;
            public uint otmEMSquare;
            public int otmAscent;
            public int otmDescent;
            public uint otmLineGap;
            public uint otmsCapEmHeight;
            public uint otmsXHeight;
            public RECT otmrcFontBox;
            public int otmMacAscent;
            public int otmMacDescent;
            public uint otmMacLineGap;
            public uint otmusMinimumPPEM;
            public POINT otmptSubscriptSize;
            public POINT otmptSubscriptOffset;
            public POINT otmptSuperscriptSize;
            public POINT otmptSuperscriptOffset;
            public uint otmsStrikeoutSize;
            public int otmsStrikeoutPosition;
            public int otmsUnderscoreSize;
            public int otmsUnderscorePosition;
        }
        #endregion
        #region PDF Generation Helpers
        private void CaptureDrawingCommands(List<string> pageCommands, bool isLandscape, int pageNumber, Dictionary<(string fontName, string style), string> fontNameToPdfName)
        {
            pdfContent.Clear();
            pdfContent.Add("q");
            foreach (string command in pageCommands)
            {
                string[] parts = command.Split('|');
                switch (parts[0])
                {
                    case "DrawString":
                        {
                            string text = parts[1];
                            string fontName = parts[2];
                            float fontSize = float.Parse(parts[3], CultureInfo.InvariantCulture);
                            string fontStyle = parts[4];
                            float x = float.Parse(parts[5], CultureInfo.InvariantCulture);
                            float y = float.Parse(parts[6], CultureInfo.InvariantCulture);
                            float pdfX = x;
                            float pdfY = pdfPageHeight - y;
                            if (!fontNameToPdfName.TryGetValue((fontName, fontStyle), out string pdfFont))
                            {
                                pdfFont = "/F1";
                            }
                            pdfContent.Add("  BT");
                            pdfContent.Add($"     {pdfFont} {fontSize.ToString(CultureInfo.InvariantCulture)} Tf");
                            pdfContent.Add($"     {pdfX.ToString(CultureInfo.InvariantCulture)} {pdfY.ToString(CultureInfo.InvariantCulture)} Td");
                            pdfContent.Add($"     ({EscapeString(text)}) Tj");
                            pdfContent.Add("  ET");
                            break;
                        }
                    case "DrawRectangle":
                        {
                            float rx = float.Parse(parts[1], CultureInfo.InvariantCulture);
                            float ry = float.Parse(parts[2], CultureInfo.InvariantCulture);
                            float width = float.Parse(parts[3], CultureInfo.InvariantCulture);
                            float height = float.Parse(parts[4], CultureInfo.InvariantCulture);
                            float pdfRX = rx;
                            float pdfRY = pdfPageHeight - ry - height;
                            pdfContent.Add($"  {pdfRX.ToString(CultureInfo.InvariantCulture)} {pdfRY.ToString(CultureInfo.InvariantCulture)} {width.ToString(CultureInfo.InvariantCulture)} {height.ToString(CultureInfo.InvariantCulture)} re");
                            pdfContent.Add("S");
                            break;
                        }
                }
            }
            pdfContent.Add("Q");
        }
        private static string EscapeString(string text)
        {
            return text.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
        }
        private (byte[] fontData, OUTLINETEXTMETRIC otm, float capHeight, uint emSquare) GetFontInfo(Font font)
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                IntPtr hdc = g.GetHdc();
                IntPtr hFont = font.ToHfont();
                IntPtr oldFont = SelectObject(hdc, hFont);
                uint fontDataSize = GetFontData(hdc, 0, 0, IntPtr.Zero, 0);
                if (fontDataSize == 0xFFFFFFFF)
                {
                    SelectObject(hdc, oldFont);
                    DeleteObject(hFont);
                    g.ReleaseHdc(hdc);
                    throw new Exception("Failed to get font data size.");
                }
                IntPtr fontDataPtr = Marshal.AllocHGlobal((int)fontDataSize);
                uint result = GetFontData(hdc, 0, 0, fontDataPtr, fontDataSize);
                if (result == 0xFFFFFFFF)
                {
                    Marshal.FreeHGlobal(fontDataPtr);
                    SelectObject(hdc, oldFont);
                    DeleteObject(hFont);
                    g.ReleaseHdc(hdc);
                    throw new Exception("Failed to retrieve font data.");
                }
                byte[] fontData = new byte[fontDataSize];
                Marshal.Copy(fontDataPtr, fontData, 0, (int)fontDataSize);
                Marshal.FreeHGlobal(fontDataPtr);
                uint otmSize = GetOutlineTextMetrics(hdc, 0, IntPtr.Zero);
                if (otmSize == 0)
                {
                    SelectObject(hdc, oldFont);
                    DeleteObject(hFont);
                    g.ReleaseHdc(hdc);
                    throw new Exception("Failed to get outline text metrics size.");
                }
                IntPtr otmPtr = Marshal.AllocHGlobal((int)otmSize);
                GetOutlineTextMetrics(hdc, otmSize, otmPtr);
                OUTLINETEXTMETRIC otm = (OUTLINETEXTMETRIC)Marshal.PtrToStructure(otmPtr, typeof(OUTLINETEXTMETRIC));
                Marshal.FreeHGlobal(otmPtr);
                GetTextExtentPoint32(hdc, "H", 1, out SIZE sizeH);
                float capHeight = sizeH.cy;
                uint emSquare = otm.otmEMSquare;
                SelectObject(hdc, oldFont);
                DeleteObject(hFont);
                g.ReleaseHdc(hdc);
                return (fontData, otm, capHeight, emSquare);
            }
        }
        private int[] GetWidths(Font font, uint emSquare)
        {
            int[] widths = new int[224];
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                IntPtr hdc = g.GetHdc();
                Font fontSized = new Font(font.FontFamily, emSquare, font.Style, GraphicsUnit.Pixel);
                IntPtr hFont = fontSized.ToHfont();
                IntPtr oldFont = SelectObject(hdc, hFont);
                for (int i = 32; i <= 255; i++)
                {
                    if (!GetCharWidth32(hdc, (uint)i, (uint)i, out int width))
                    {
                        width = 0;
                    }
                    int pdfWidth = (int)Math.Round((double)width * 1000 / emSquare);
                    widths[i - 32] = pdfWidth;
                }
                SelectObject(hdc, oldFont);
                DeleteObject(hFont);
                g.ReleaseHdc(hdc);
                fontSized.Dispose();
            }
            return widths;
        }
        #endregion
        #region Form Setup and Data
        private void AddSampleData()
        {
            Random random = new Random();
            for (int i = 0; i < 200; i++)
            {
                Junctions.Rows.Add($"J{i + 1}", random.NextDouble() * 100, random.NextDouble() * 100, random.NextDouble() * 100);
                Pipes.Rows.Add($"P{i + 1}", $"J{random.Next(1, 101)}", $"J{random.Next(1, 101)}");
                Vertices.Rows.Add($"V{i + 1}", $"P{random.Next(1, 101)}", random.NextDouble() * 100, random.NextDouble() * 100);
            }
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            SetPrintingSettings();
            dataGridViews.Add(Junctions);
            dataGridViews.Add(Pipes);
            dataGridViews.Add(Vertices);
            Junctions.Columns.Add("JunctionID", "Junction ID");
            Junctions.Columns.Add("X", "X Coordinate");
            Junctions.Columns.Add("Y", "Y Coordinate");
            Junctions.Columns.Add("Z", "Z Coordinate");
            Pipes.Columns.Add("PipeID", "Pipe ID");
            Pipes.Columns.Add("StartJunction", "Start Junction");
            Pipes.Columns.Add("EndJunction", "End Junction");
            Vertices.Columns.Add("VertexID", "Vertex ID");
            Vertices.Columns.Add("PipeID", "Pipe ID");
            Vertices.Columns.Add("X", "X Coordinate");
            Vertices.Columns.Add("Y", "Y Coordinate");
            AddSampleData();
        }
        #endregion
        #region Printing and Preview
        private void CaptureIntermediateLanguage()
        {
            intermediateCommands.Clear();
            graphicsRecorder.ClearCommands();
            bool isLandscape = PrintDocument1.DefaultPageSettings.Landscape;
            int page = 1;
            foreach (DataGridView dgv in dataGridViews)
            {
                int tempRow = 0;
                while (tempRow < dgv.Rows.Count)
                {
                    using (Bitmap tempBitmap = new Bitmap(1, 1))
                    using (Graphics tempGraphics = Graphics.FromImage(tempBitmap))
                    {
                        graphicsRecorder.SetPage(page);
                        tempRow = DrawContent(dgv, tempGraphics, true, isLandscape, tempRow, page, totalPageCount);
                        intermediateCommands.AddRange(graphicsRecorder.Commands);
                        graphicsRecorder.ClearCommands();
                        page++;
                    }
                }
            }
            totalPageCount = intermediateCommands.Select(cmd => int.Parse(cmd.Split('|').Last())).Distinct().Count();
        }
        private int CalculateTotalPages(Graphics graphics, bool isLandscape)
        {
            int pages = 0;
            foreach (DataGridView dgv in dataGridViews)
            {
                int tempRow = 0;
                while (tempRow < dgv.Rows.Count)
                {
                    int nextRow = DrawContent(dgv, graphics, false, isLandscape, tempRow, pages + 1, pages + 1);
                    if (nextRow == tempRow)
                    {
                        break;
                    }
                    pages++;
                    tempRow = nextRow;
                }
            }
            return pages;
        }
        private int DrawContent(DataGridView dgv, Graphics graphics, bool recordCommands, bool isLandscape, int currentRow, int pageNumber, int totalPageCount)
        {
            const float HundredthsOfInchToPoints = 72f / 100f;
            float marginLeft = PageSetupDialog1.PageSettings.Margins.Left * HundredthsOfInchToPoints;
            float marginTop = PageSetupDialog1.PageSettings.Margins.Top * HundredthsOfInchToPoints;
            float marginRight = PageSetupDialog1.PageSettings.Margins.Right * HundredthsOfInchToPoints;
            float marginBottom = PageSetupDialog1.PageSettings.Margins.Bottom * HundredthsOfInchToPoints;
            if (graphics != null)
            {
                graphics.PageUnit = GraphicsUnit.Point;
            }
            Font pageNumberFont = new Font("Arial", 9, FontStyle.Regular);
            string pageNumberText = $"Page {pageNumber} / {totalPageCount}";
            float computedPageWidth = (float)Math.Round(PageSetupDialog1.PageSettings.PaperSize.Width / 100.0f * 72.0f);
            float computedPageHeight = (float)Math.Round(PageSetupDialog1.PageSettings.PaperSize.Height / 100.0f * 72.0f);
            if (isLandscape)
            {
                float temp = computedPageWidth;
                computedPageWidth = computedPageHeight;
                computedPageHeight = temp;
            }
            SizeF pageNumberSize = graphics?.MeasureString(pageNumberText, pageNumberFont) ?? new SizeF(0, 0);
            float pageNumberX = computedPageWidth - marginRight - pageNumberSize.Width;
            float pageNumberY = marginTop;
            if (recordCommands)
            {
                graphicsRecorder.DrawString(pageNumberText, pageNumberFont, Brushes.Black, pageNumberX, pageNumberY + 3);
            }
            graphics?.DrawString(pageNumberText, pageNumberFont, Brushes.Black, pageNumberX, pageNumberY + 3);
            Font titleFont = new Font("Arial", 12, FontStyle.Bold);
            Font headerFont = new Font("Arial", 9, FontStyle.Bold);
            Font cellFont = new Font("Arial", 8);
            float rowHeight = 15;
            float cellPadding = 3;
            float printableWidth = computedPageWidth - marginLeft - marginRight;
            float maxY = computedPageHeight - marginBottom;
            float currentY = marginTop;
            string title = dgv.Name + " Data";
            if (recordCommands)
            {
                graphicsRecorder.DrawString(title, titleFont, Brushes.Black, marginLeft, currentY);
            }
            graphics?.DrawString(title, titleFont, Brushes.Black, marginLeft, currentY);
            currentY += titleFont.GetHeight() + 10;
            int columnCount = dgv.Columns.Count;
            float columnWidth = printableWidth / columnCount;
            float currentX = marginLeft;
            for (int i = 0; i < columnCount; i++)
            {
                string headerText = dgv.Columns[i].HeaderText;
                if (recordCommands)
                {
                    graphicsRecorder.DrawRectangle(Pens.Black, currentX, currentY, columnWidth, rowHeight);
                    graphicsRecorder.DrawString(headerText, headerFont, Brushes.Black, currentX + cellPadding, currentY + cellPadding);
                }
                graphics?.DrawRectangle(Pens.Black, currentX, currentY, columnWidth, rowHeight);
                graphics?.DrawString(headerText, headerFont, Brushes.Black, currentX + cellPadding, currentY + cellPadding);
                currentX += columnWidth;
            }
            currentY += rowHeight;
            for (int i = currentRow; i < dgv.Rows.Count; i++)
            {
                DataGridViewRow row = dgv.Rows[i];
                if (row.IsNewRow)
                {
                    continue;
                }
                if (currentY + rowHeight > maxY)
                {
                    return i;
                }
                currentX = marginLeft;
                for (int j = 0; j < columnCount; j++)
                {
                    string cellValue = row.Cells[j].Value?.ToString() ?? "";
                    if (recordCommands)
                    {
                        graphicsRecorder.DrawRectangle(Pens.Black, currentX, currentY, columnWidth, rowHeight);
                        graphicsRecorder.DrawString(cellValue, cellFont, Brushes.Black, currentX + cellPadding, currentY + cellPadding);
                    }
                    graphics?.DrawRectangle(Pens.Black, currentX, currentY, columnWidth, rowHeight);
                    graphics?.DrawString(cellValue, cellFont, Brushes.Black, currentX + cellPadding, currentY + cellPadding);
                    currentX += columnWidth;
                }
                currentY += rowHeight;
            }
            return dgv.Rows.Count;
        }
        private void PrintDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (currentDGVIndex < dataGridViews.Count)
            {
                DataGridView dgv = dataGridViews[currentDGVIndex];
                int nextRow = DrawContent(dgv, e.Graphics, false, PrintDocument1.DefaultPageSettings.Landscape, currentRow, currentPageNumber, totalPageCount);
                if (nextRow < dgv.Rows.Count)
                {
                    currentRow = nextRow;
                    e.HasMorePages = true;
                    currentPageNumber++;
                }
                else
                {
                    currentDGVIndex++;
                    currentRow = 0;
                    if (currentDGVIndex < dataGridViews.Count)
                    {
                        e.HasMorePages = true;
                        currentPageNumber++;
                    }
                    else
                    {
                        e.HasMorePages = false;
                        currentDGVIndex = 0;
                        currentRow = 0;
                        currentPageNumber = 1;
                    }
                }
            }
            else
            {
                e.HasMorePages = false;
            }
        }
        private void PrintDocument1_BeginPrint(object sender, PrintEventArgs e)
        {
            int paperWidthHundredths = PrintDocument1.DefaultPageSettings.PaperSize.Width;
            int paperHeightHundredths = PrintDocument1.DefaultPageSettings.PaperSize.Height;
            pdfPageWidth = (int)Math.Round(paperWidthHundredths / 100.0f * 72.0f);
            pdfPageHeight = (int)Math.Round(paperHeightHundredths / 100.0f * 72.0f);
            if (PrintDocument1.DefaultPageSettings.Landscape)
            {
                int temp = pdfPageWidth;
                pdfPageWidth = pdfPageHeight;
                pdfPageHeight = temp;
            }
            using (Bitmap bmp = new Bitmap(1, 1))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                totalPageCount = CalculateTotalPages(g, PrintDocument1.DefaultPageSettings.Landscape);
            }
            currentDGVIndex = 0;
            currentRow = 0;
            currentPageNumber = 1;
        }
        private void SetPrintingSettings()
        {
            PageSetupDialog1.Document = PrintDocument1;
            PageSetupDialog1.EnableMetric = true;
            PrintPreviewDialog1.Document = PrintDocument1;
            PrintDialog1.Document = PrintDocument1;
            IEnumerable<PaperSize> paperSizes = PageSetupDialog1.PrinterSettings.PaperSizes.Cast<PaperSize>();
            PaperSize sizeA4 = paperSizes.First(size => size.Kind == PaperKind.A4);
            PageSetupDialog1.PageSettings = new PageSettings()
            {
                Margins = PrinterUnitConvert.Convert(initialMargins, PrinterUnit.HundredthsOfAMillimeter, PrinterUnit.ThousandthsOfAnInch),
                PaperSize = sizeA4
            };
            PrintDocument1.PrinterSettings = PageSetupDialog1.PrinterSettings;
            PrintDocument1.DefaultPageSettings = PageSetupDialog1.PageSettings;
            printingSettingsSet = true;
        }
        #endregion
        #region Button Handlers
        private uint Adler32(byte[] data)
        {
            const uint MOD_ADLER = 65521;
            uint a = 1, b = 0;
            foreach (byte byteValue in data)
            {
                a = (a + byteValue) % MOD_ADLER;
                b = (b + a) % MOD_ADLER;
            }
            return (b << 16) | a;
        }
        private byte[] CompressWithZlib(byte[] data)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ms.WriteByte(0x78);
                ms.WriteByte(0x9C);
                using (DeflateStream deflateStream = new DeflateStream(ms, CompressionMode.Compress, true))
                {
                    deflateStream.Write(data, 0, data.Length);
                }
                uint adler32 = Adler32(data);
                byte[] adlerBytes = BitConverter.GetBytes(adler32);
                if (BitConverter.IsLittleEndian)
                {
                    Array.Reverse(adlerBytes);
                }
                ms.Write(adlerBytes, 0, 4);
                return ms.ToArray();
            }
        }
        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("© 2025\nMustafa Sami Altas ", "About...", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void ConverttoPDF_Click(object sender, EventArgs e)
        {
            SaveFileDialog1.Filter = "PDF Files (*.pdf)|*.pdf";
            SaveFileDialog1.Title = "Save PDF File";
            SaveFileDialog1.DefaultExt = "pdf";
            if (SaveFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string pdfPath = SaveFileDialog1.FileName;
            if (Path.GetExtension(pdfPath).ToLower() != ".pdf")
            {
                pdfPath += ".pdf";
            }
            string title = Path.GetFileNameWithoutExtension(pdfPath);
            pdfContent.Clear();

            using (Bitmap bmp = new Bitmap(1, 1))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                totalPageCount = CalculateTotalPages(g, PrintDocument1.DefaultPageSettings.Landscape);
            }

            CaptureIntermediateLanguage();
            bool isLandscape = PrintDocument1.DefaultPageSettings.Landscape;
            int paperWidthHundredths = PrintDocument1.DefaultPageSettings.PaperSize.Width;
            int paperHeightHundredths = PrintDocument1.DefaultPageSettings.PaperSize.Height;
            pdfPageWidth = (int)Math.Round(paperWidthHundredths / 100.0f * 72.0f);
            pdfPageHeight = (int)Math.Round(paperHeightHundredths / 100.0f * 72.0f);
            if (isLandscape)
            {
                int temp = pdfPageWidth;
                pdfPageWidth = pdfPageHeight;
                pdfPageHeight = temp;
            }
            List<PdfObject> pdfObjects = new List<PdfObject>();
            int nextObjNumber = 1;
            int catalogObj = nextObjNumber++;
            int pagesTreeObj = nextObjNumber++;
            Dictionary<(string fontName, string style), Font> uniqueFonts = new Dictionary<(string, string), Font>();
            foreach (string cmd in intermediateCommands)
            {
                if (cmd.StartsWith("DrawString"))
                {
                    string[] parts = cmd.Split('|');
                    string fontName = parts[2];
                    string styleStr = parts[4];
                    FontStyle style = FontStyle.Regular;
                    foreach (string s in styleStr.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        style |= (FontStyle)Enum.Parse(typeof(FontStyle), s);
                    }
                    (string fontName, string styleStr) key = (fontName, styleStr);
                    if (!uniqueFonts.ContainsKey(key))
                    {
                        Font font = new Font(fontName, 10, style);
                        uniqueFonts[key] = font;
                    }
                }
            }
            Dictionary<(string fontName, string style), (byte[] fontData, OUTLINETEXTMETRIC otm, float capHeight, uint emSquare)> fontInfos = new Dictionary<(string, string), (byte[], OUTLINETEXTMETRIC, float, uint)>();
            Dictionary<(string fontName, string style), string> fontNameToPdfName = new Dictionary<(string, string), string>();
            Dictionary<string, int> fontObjNumbers = new Dictionary<string, int>();
            int fontCounter = 1;
            foreach (KeyValuePair<(string fontName, string style), Font> kv in uniqueFonts)
            {
                Font font = kv.Value;
                (byte[] fontData, OUTLINETEXTMETRIC otm, float capHeight, uint emSquare) info = GetFontInfo(font);
                fontInfos[kv.Key] = info;
                string pdfFontName = "/F" + fontCounter++;
                fontNameToPdfName[kv.Key] = pdfFontName;
                int fontFileObj = nextObjNumber++;
                byte[] compressedFontData = CompressWithZlib(info.fontData);
                pdfObjects.Add(new PdfObject(fontFileObj,
                $" << /Length {compressedFontData.Length}\n" +
                $" /Filter /FlateDecode\n" +
                $" /Length1 {info.fontData.Length}\n" +
                $" >>\r\nstream\r\n", compressedFontData, "\r\nendstream"));
                string baseFont = kv.Key.fontName.Replace(" ", "") + "-" + kv.Key.style.Replace(", ", "");
                int flags = 32;
                if ((font.Style & FontStyle.Italic) != 0)
                {
                    flags |= 64;
                }

                int stemV = 80;
                int fontDescObj = nextObjNumber++;
                pdfObjects.Add(new PdfObject(fontDescObj,
                $" << /Type /FontDescriptor\n" +
                $" /FontName /{baseFont}\n" +
                $" /Flags {flags}\n" +
                $" /FontBBox [{info.otm.otmrcFontBox.left} {info.otm.otmrcFontBox.bottom} {info.otm.otmrcFontBox.right} {info.otm.otmrcFontBox.top}]\n" +
                $" /ItalicAngle {info.otm.otmItalicAngle}\n" +
                $" /Ascent {info.otm.otmAscent}\n" +
                $" /Descent {info.otm.otmDescent}\n" +
                $" /CapHeight {info.capHeight}\n" +
                $" /StemV {stemV}\n" +
                $" /FontFile2 {fontFileObj} 0 R\n" +
                " >>"));
                int fontObj = nextObjNumber++;
                int[] widths = GetWidths(font, info.emSquare);
                string widthsStr = "[" + string.Join(" ", widths) + "]";
                pdfObjects.Add(new PdfObject(fontObj,
                $" << /Type /Font\n" +
                $" /Subtype /TrueType\n" +
                $" /BaseFont /{baseFont}\n" +
                $" /Encoding /WinAnsiEncoding\n" +
                $" /FirstChar 32\n" +
                $" /LastChar 255\n" +
                $" /Widths {widthsStr}\n" +
                $" /FontDescriptor {fontDescObj} 0 R\n" +
                " >>"));
                fontObjNumbers[pdfFontName] = fontObj;
            }
            string fontResources = "<< /Font << ";
            foreach (KeyValuePair<string, int> kv in fontObjNumbers)
            {
                fontResources += $"{kv.Key} {kv.Value} 0 R ";
            }
            fontResources += ">> >>";
            int structTreeRootObj = nextObjNumber++;
            int documentStructObj = nextObjNumber++;
            int parentTreeObj = nextObjNumber++;
            List<int> allStructElems = new List<int>();
            Dictionary<int, List<(int structElemObj, string structType)>> pageStructElems = new Dictionary<int, List<(int, string)>>();
            List<int> pageObjNumbers = new List<int>();
            for (int pageNum = 1; pageNum <= totalPageCount; pageNum++)
            {
                List<string> pageCommands = intermediateCommands.Where(cmd => cmd.Split('|').Last() == pageNum.ToString()).ToList();
                List<string> structTypesForPage = new List<string>();
                CaptureDrawingCommands(pageCommands, isLandscape, pageNum, fontNameToPdfName);
                string contentStr = string.Join("\r\n", pdfContent);
                int contentLength = Encoding.ASCII.GetByteCount(contentStr);
                int contentObj = nextObjNumber++;
                pdfObjects.Add(new PdfObject(contentObj,
                $" << /Length {contentLength}\n" +
                $" >>\r\nstream\r\n{contentStr}\r\nendstream"));
                int pageObj = nextObjNumber++;
                pageObjNumbers.Add(pageObj);
                pdfObjects.Add(new PdfObject(pageObj,
                $" << /Type /Page\n" +
                $" /Parent {pagesTreeObj} 0 R\n" +
                $" /MediaBox [0 0 {pdfPageWidth} {pdfPageHeight}]\n" +
                $" /Resources {fontResources}\n" +
                $" /Contents {contentObj} 0 R\n" +
                $" /StructParents {pageNum - 1}\n" +
                " >>"));
                int pageIndex = pageNum - 1;
                List<(int structElemObj, string structType)> pageStructElemsList = new List<(int, string)>();
                for (int mcid = 0; mcid < structTypesForPage.Count; mcid++)
                {
                    string structType = structTypesForPage[mcid];
                    int structElemObj = nextObjNumber++;
                    allStructElems.Add(structElemObj);
                    pageStructElemsList.Add((structElemObj, structType));
                }
                pageStructElems[pageIndex] = pageStructElemsList;
            }
            string kidsArray = string.Join(" ", pageObjNumbers.Select(n => $"{n} 0 R"));
            pdfObjects.Add(new PdfObject(pagesTreeObj,
            $" << /Type /Pages\n" +
            $" /Count {pageObjNumbers.Count}\n" +
            $" /Kids [{kidsArray}]\n" +
            $" >>"));
            int infoObj = nextObjNumber++;
            int outlinesObj = nextObjNumber++;
            int metadataObj = nextObjNumber++;
            int outputIntentObj = nextObjNumber++;
            int iccProfileObj = nextObjNumber++;
            byte[] iccBytes = PDF_PrintPreview_XLSX_Export.Properties.Resources.sRGB_IEC61966_2_1;
            byte[] compressedIccBytes = CompressWithZlib(iccBytes);
            pdfObjects.Add(new PdfObject(iccProfileObj,
            $" << /N 3\n" +
            $" /Length {compressedIccBytes.Length}\n" +
            $" /Filter /FlateDecode\n" +
            $" >>\r\n" +
            "stream\r\n", compressedIccBytes, "\r\nendstream"));
            pdfObjects.Add(new PdfObject(outputIntentObj,
            $" << /Type /OutputIntent\r\n" +
            $" /S /GTS_PDFA1\r\n" +
            $" /OutputCondition (sRGB IEC61966-2.1)\r\n" +
            $" /OutputConditionIdentifier (sRGB IEC61966-2.1)\r\n" +
            $" /RegistryName (http://www.color.org)\r\n" +
            $" /DestOutputProfile {iccProfileObj} 0 R\r\n" +
            $" /Info (sRGB IEC61966-2.1)\r\n" +
            " >>"));
            pdfObjects.Add(new PdfObject(catalogObj,
            $" << /Type /Catalog\n" +
            $" /Pages {pagesTreeObj} 0 R\n" +
            $" /Outlines {outlinesObj} 0 R\n" +
            $" /Metadata {metadataObj} 0 R\n" +
            $" /Lang (en-US)\n" +
            $" /MarkInfo << /Marked true >>\n" +
            $" /StructTreeRoot {structTreeRootObj} 0 R\n" +
            $" /ViewerPreferences << /DisplayDocTitle true >>\n" +
            $" /OutputIntents [{outputIntentObj} 0 R]\n" +
            $" >>"));
            string numsArray = string.Join(" ", pageStructElems.OrderBy(kv => kv.Key)
            .Select(kv => $"{kv.Key} [{string.Join(" ", kv.Value.Select(t => $"{t.structElemObj} 0 R"))}]"));
            pdfObjects.Add(new PdfObject(parentTreeObj,
            $" << /Nums [{numsArray}]\n" +
            $" >>"));
            string documentK = string.Join(" ", allStructElems.Select(n => $"{n} 0 R"));
            pdfObjects.Add(new PdfObject(documentStructObj,
            $" << /Type /StructElem\n" +
            $" /S /Document\n" +
            $" /P {structTreeRootObj} 0 R\n" +
            $" /K [{documentK}]\n" +
            " >>"));
            foreach (KeyValuePair<int, List<(int structElemObj, string structType)>> kv in pageStructElems)
            {
                int pageIndex = kv.Key;
                int pageObj = pageObjNumbers[pageIndex];
                for (int i = 0; i < kv.Value.Count; i++)
                {
                    (int structElemObj, string structType) = kv.Value[i];
                    string mcr = $" << /Type /MCR\n" +
                    $" /Pg {pageObj} 0 R\n" +
                    $" /MCID {i}\n" +
                    $" >>";
                    pdfObjects.Add(new PdfObject(structElemObj,
                    $" << /Type /StructElem\n" +
                    $" /S /{structType}\n" +
                    $" /P {documentStructObj} 0 R\n" +
                    $" /K {mcr}\n" +
                    $" >>"));
                }
            }
            pdfObjects.Add(new PdfObject(structTreeRootObj,
            $" << /Type /StructTreeRoot\n" +
            $" /K {documentStructObj} 0 R\n" +
            $" /ParentTree {parentTreeObj} 0 R\n" +
            $" >>"));
            string creationDateStr = $"D:{DateTime.Now:yyyyMMddHHmmss}{DateTime.Now.ToString("zzz").Replace(":", "'")}'";
            string modDateStr = $"D:{DateTime.Now:yyyyMMddHHmmss}{DateTime.Now.ToString("zzz").Replace(":", "'")}'";
            pdfObjects.Add(new PdfObject(infoObj,
            $" << /Title ({EscapeString(title)})\n" +
            $" /Creator ({Application.ProductName})\n" +
            $" /Author (Mustafa Sami Altas)\n" +
            $" /ModDate ({modDateStr})\n" +
            $" /CreationDate ({creationDateStr})\n" +
            $" /Producer (Produced by {Application.ProductName})\n" +
            $" /Subject (Sample INP File)\n" +
            $" /Trapped /False\n" +
            $" >>"));
            pdfObjects.Add(new PdfObject(outlinesObj,
            $" << /Type /Outlines\n" +
            $" >>"));
            string metadataXml = $@"<?xpacket begin=""ï»¿"" id=""W5M0MpCehiHzreSzNTczkc9d""?>
<x:xmpmeta xmlns:x=""adobe:ns:meta/"" x:xmptk=""Adobe XMP Core 9.1-c001 79.675d0f7, 2023/06/11-19:21:16"">
<rdf:RDF xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#"">
<rdf:Description rdf:about=""""
xmlns:dc=""http://purl.org/dc/elements/1.1/""
xmlns:xmp=""http://ns.adobe.com/xap/1.0/""
xmlns:pdf=""http://ns.adobe.com/pdf/1.3/""
xmlns:xmpMM=""http://ns.adobe.com/xap/1.0/mm/""
xmlns:stEvt=""http://ns.adobe.com/xap/1.0/sType/ResourceEvent#""
xmlns:pdfaid=""http://www.aiim.org/pdfa/ns/id/""
xmlns:pdfuaid=""http://www.aiim.org/pdfua/ns/id/""
xmlns:pdfaExtension=""http://www.aiim.org/pdfa/ns/extension/""
xmlns:pdfaSchema=""http://www.aiim.org/pdfa/ns/schema#""
xmlns:pdfaProperty=""http://www.aiim.org/pdfa/ns/property#"">
<dc:format>application/pdf</dc:format>
<dc:title>
<rdf:Alt>
<rdf:li xml:lang=""x-default"">{EscapeString(title)}</rdf:li>
</rdf:Alt>
</dc:title>
<dc:creator>
<rdf:Seq>
<rdf:li>Mustafa Sami Altas</rdf:li>
</rdf:Seq>
</dc:creator>
<dc:description>
<rdf:Alt>
<rdf:li xml:lang=""x-default"">Sample INP File</rdf:li>
</rdf:Alt>
</dc:description>
<xmp:CreatorTool>PDF_PrintPreview_XLSX_Export</xmp:CreatorTool>
<xmp:ModifyDate>{DateTime.Now:yyyy-MM-ddTHH:mm:sszzz}</xmp:ModifyDate>
<xmp:MetadataDate>{DateTime.Now:yyyy-MM-ddTHH:mm:sszzz}</xmp:MetadataDate>
<xmp:CreateDate>{DateTime.Now:yyyy-MM-ddTHH:mm:sszzz}</xmp:CreateDate>
<pdf:Producer>Produced by PDF_PrintPreview_XLSX_Export</pdf:Producer>
<pdf:Trapped>False</pdf:Trapped>
<xmpMM:DocumentID>uuid:b108e64c-9389-48a2-85cd-a5cfbf9ac78f</xmpMM:DocumentID>
<xmpMM:InstanceID>uuid:e27a3385-4ffd-42da-8f8a-9271065266b1</xmpMM:InstanceID>
<xmpMM:RenditionClass>default</xmpMM:RenditionClass>
<xmpMM:VersionID>1</xmpMM:VersionID>
<xmpMM:History>
<rdf:Seq>
<rdf:li rdf:parseType=""Resource"">
<stEvt:action>converted</stEvt:action>
<stEvt:instanceID>uuid:1191185c-7008-4ab8-8970-7663addacd69</stEvt:instanceID>
<stEvt:parameters>converted to PDF/A-2b</stEvt:parameters>
<stEvt:softwareAgent>Preflight</stEvt:softwareAgent>
<stEvt:when>{DateTime.Now:yyyy-MM-ddTHH:mm:ssZ}</stEvt:when>
</rdf:li>
</rdf:Seq>
</xmpMM:History>
<pdfaid:part>1</pdfaid:part>
<pdfaid:conformance>A</pdfaid:conformance>
<pdfuaid:part>1</pdfuaid:part>
<pdfaExtension:schemas>
<rdf:Bag>
<rdf:li rdf:parseType=""Resource"">
<pdfaSchema:namespaceURI>http://ns.adobe.com/pdf/1.3/</pdfaSchema:namespaceURI>
<pdfaSchema:prefix>pdf</pdfaSchema:prefix>
<pdfaSchema:schema>Adobe PDF Schema</pdfaSchema:schema>
<pdfaSchema:property>
<rdf:Seq>
<rdf:li rdf:parseType=""Resource"">
<pdfaProperty:category>internal</pdfaProperty:category>
<pdfaProperty:description>A name object indicating whether the document has been modified to include trapping information</pdfaProperty:description>
<pdfaProperty:name>Trapped</pdfaProperty:name>
<pdfaProperty:valueType>Text</pdfaProperty:valueType>
</rdf:li>
</rdf:Seq>
</pdfaSchema:property>
</rdf:li>
<rdf:li rdf:parseType=""Resource"">
<pdfaSchema:namespaceURI>http://ns.adobe.com/xap/1.0/mm/</pdfaSchema:namespaceURI>
<pdfaSchema:prefix>xmpMM</pdfaSchema:prefix>
<pdfaSchema:schema>XMP Media Management Schema</pdfaSchema:schema>
<pdfaSchema:property>
<rdf:Seq>
<rdf:li rdf:parseType=""Resource"">
<pdfaProperty:category>internal</pdfaProperty:category>
<pdfaProperty:description>UUID based identifier for specific incarnation of a document</pdfaProperty:description>
<pdfaProperty:name>InstanceID</pdfaProperty:name>
<pdfaProperty:valueType>URI</pdfaProperty:valueType>
</rdf:li>
<rdf:li rdf:parseType=""Resource"">
<pdfaProperty:category>internal</pdfaProperty:category>
<pdfaProperty:description>The common identifier for all versions and renditions of a document.</pdfaProperty:description>
<pdfaProperty:name>OriginalDocumentID</pdfaProperty:name>
<pdfaProperty:valueType>URI</pdfaProperty:valueType>
</rdf:li>
</rdf:Seq>
</pdfaSchema:property>
</rdf:li>
<rdf:li rdf:parseType=""Resource"">
<pdfaSchema:namespaceURI>http://www.aiim.org/pdfa/ns/id/</pdfaSchema:namespaceURI>
<pdfaSchema:prefix>pdfaid</pdfaSchema:prefix>
<pdfaSchema:schema>PDF/A ID Schema</pdfaSchema:schema>
<pdfaSchema:property>
<rdf:Seq>
<rdf:li rdf:parseType=""Resource"">
<pdfaProperty:category>internal</pdfaProperty:category>
<pdfaProperty:description>Part of PDF/A standard</pdfaProperty:description>
<pdfaProperty:name>part</pdfaProperty:name>
<pdfaProperty:valueType>Integer</pdfaProperty:valueType>
</rdf:li>
<rdf:li rdf:parseType=""Resource"">
<pdfaProperty:category>internal</pdfaProperty:category>
<pdfaProperty:description>Amendment of PDF/A standard</pdfaProperty:description>
<pdfaProperty:name>amd</pdfaProperty:name>
<pdfaProperty:valueType>Text</pdfaProperty:valueType>
</rdf:li>
<rdf:li rdf:parseType=""Resource"">
<pdfaProperty:category>internal</pdfaProperty:category>
<pdfaProperty:description>Conformance level of PDF/A standard</pdfaProperty:description>
<pdfaProperty:name>conformance</pdfaProperty:name>
<pdfaProperty:valueType>Text</pdfaProperty:valueType>
</rdf:li>
</rdf:Seq>
</pdfaSchema:property>
</rdf:li>
<rdf:li rdf:parseType=""Resource"">
<pdfaSchema:namespaceURI>http://www.aiim.org/pdfua/ns/id/</pdfaSchema:namespaceURI>
<pdfaSchema:prefix>pdfuaid</pdfaSchema:prefix>
<pdfaSchema:schema>PDF/UA Universal Accessibility Schema</pdfaSchema:schema>
<pdfaSchema:property>
<rdf:Seq>
<rdf:li rdf:parseType=""Resource"">
<pdfaProperty:category>internal</pdfaProperty:category>
<pdfaProperty:description>Indicates, which part of ISO 14289 standard is followed</pdfaProperty:description>
<pdfaProperty:name>part</pdfaProperty:name>
<pdfaProperty:valueType>Integer</pdfaProperty:valueType>
</rdf:li>
</rdf:Seq>
</pdfaSchema:property>
</rdf:li>
</rdf:Bag>
</pdfaExtension:schemas>
</rdf:Description>
</rdf:RDF>
</x:xmpmeta>
<?xpacket end=""r""?>";
            int metadataLength = Encoding.ASCII.GetByteCount(metadataXml);
            pdfObjects.Add(new PdfObject(metadataObj,
            $" << /Type\r\n" +
            $" /Metadata\r\n" +
            $" /Length {metadataLength}\r\n" +
            $" /Subtype /XML\r\n" +
            $" >>\r\n" +
            $"stream\r\n" +
            $"{metadataXml}\r\n" +
            $"endstream"));
            using (FileStream fs = new FileStream(pdfPath, FileMode.Create))
            {
                byte[] header = Encoding.ASCII.GetBytes("%PDF-1.7\n");
                header = header.Concat(new byte[] { (byte)'%', (byte)'\xE2', (byte)'\xE3', (byte)'\xCF', (byte)'\xD3', (byte)'\n' }).ToArray();
                fs.Write(header, 0, header.Length);
                List<long> offsets = new List<long>();
                foreach (PdfObject obj in pdfObjects.OrderBy(o => o.Number))
                {
                    offsets.Add(fs.Position);
                    byte[] objHeader = Encoding.ASCII.GetBytes($"{obj.Number} 0 obj\r\n");
                    fs.Write(objHeader, 0, objHeader.Length);
                    byte[] contentBytes = Encoding.ASCII.GetBytes(obj.Content);
                    fs.Write(contentBytes, 0, contentBytes.Length);
                    if (obj.StreamData != null)
                    {
                        fs.Write(obj.StreamData, 0, obj.StreamData.Length);
                        byte[] streamFooterBytes = Encoding.ASCII.GetBytes(obj.StreamFooter);
                        fs.Write(streamFooterBytes, 0, streamFooterBytes.Length);
                    }
                    byte[] objFooter = Encoding.ASCII.GetBytes("\r\nendobj\r\n");
                    fs.Write(objFooter, 0, objFooter.Length);
                }
                long xrefPosition = fs.Position;
                byte[] xrefHeader = Encoding.ASCII.GetBytes("xref\r\n");
                fs.Write(xrefHeader, 0, xrefHeader.Length);
                byte[] xrefLine1 = Encoding.ASCII.GetBytes($"0 {nextObjNumber}\r\n");
                fs.Write(xrefLine1, 0, xrefLine1.Length);
                byte[] freeObj = Encoding.ASCII.GetBytes("0000000000 65535 f\r\n");
                fs.Write(freeObj, 0, freeObj.Length);
                foreach (long offset in offsets)
                {
                    string line = $"{offset:D10} 00000 n\r\n";
                    byte[] lineBytes = Encoding.ASCII.GetBytes(line);
                    fs.Write(lineBytes, 0, lineBytes.Length);
                }
                byte[] trailerStart = Encoding.ASCII.GetBytes("trailer\r\n");
                fs.Write(trailerStart, 0, trailerStart.Length);
                string id1 = Guid.NewGuid().ToString("N").ToUpper();
                string id2 = Guid.NewGuid().ToString("N").ToUpper();
                string trailerDict =
                $" << /Root {catalogObj} 0 R\n" +
                $" /Size {nextObjNumber}\n" +
                $" /ID [<{id1}> <{id2}>]\n" +
                $" /Info {infoObj} 0 R\n" +
                $" >>\r\n";
                byte[] trailerDictBytes = Encoding.ASCII.GetBytes(trailerDict);
                fs.Write(trailerDictBytes, 0, trailerDictBytes.Length);
                byte[] startxref = Encoding.ASCII.GetBytes("startxref\r\n");
                fs.Write(startxref, 0, startxref.Length);
                byte[] xrefPosBytes = Encoding.ASCII.GetBytes($"{xrefPosition}\r\n");
                fs.Write(xrefPosBytes, 0, xrefPosBytes.Length);
                byte[] eof = Encoding.ASCII.GetBytes("%%EOF\r\n");
                fs.Write(eof, 0, eof.Length);
            }
            Process.Start(pdfPath);
            foreach (Font font in uniqueFonts.Values)
            {
                font.Dispose();
            }
        }
        private void PrintPreview_Click(object sender, EventArgs e)
        {
            currentDGVIndex = 0;
            currentRow = 0;
            try { PrintPreviewDialog1.ShowDialog(); }
            catch (System.ComponentModel.Win32Exception) { }
        }
        private void PageSetup_Click(object sender, EventArgs e)
        {
            if (!printingSettingsSet)
            {
                SetPrintingSettings();
            }
            if (PageSetupDialog1.ShowDialog() == DialogResult.OK)
            {
                PrintDocument1.PrinterSettings = PageSetupDialog1.PrinterSettings;
                PrintDocument1.DefaultPageSettings = PageSetupDialog1.PageSettings;
            }
        }
        private void ConverttoXLSX_Click(object sender, EventArgs e)
        {
            isExportingToExcel = true;
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Save Excel File"
                };
                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }
                string excelFilePath = saveFileDialog.FileName;
                if (Path.GetExtension(excelFilePath).ToLower() != ".xlsx")
                {
                    excelFilePath += ".xlsx";
                }
                CaptureIntermediateLanguage();
                if (intermediateCommands == null || !intermediateCommands.Any())
                {
                    MessageBox.Show("No intermediate commands available to generate the Excel file.");
                    return;
                }
                List<IGrouping<int, string>> pages = intermediateCommands.GroupBy(cmd => int.Parse(cmd.Split('|').Last())).OrderBy(g => g.Key).ToList();
                using (MemoryStream ms = new MemoryStream())
                {
                    using (ZipArchive archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                    {
                        string contentTypesXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Default Extension=""xml"" ContentType=""application/xml""/>
<Default Extension=""bin"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings""/>
<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>";
                        for (int i = 0; i < pages.Count; i++)
                        {
                            contentTypesXml += $@"<Override PartName=""/xl/worksheets/sheet{i + 1}.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>";
                        }
                        contentTypesXml += "</Types>";
                        AddTextFileToZip(archive, "[Content_Types].xml", contentTypesXml);
                        AddTextFileToZip(archive, "_rels/.rels", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");
                        string workbookRels = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">";
                        for (int i = 0; i < pages.Count; i++)
                        {
                            workbookRels += $@"<Relationship Id=""rId{i + 1}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet{i + 1}.xml""/>";
                        }
                        workbookRels += $@"<Relationship Id=""rId{pages.Count + 1}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>";
                        workbookRels += "</Relationships>";
                        AddTextFileToZip(archive, "xl/_rels/workbook.xml.rels", workbookRels);
                        string workbookXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets>";
                        for (int i = 0; i < pages.Count; i++)
                        {
                            string sheetName = $"Page{i + 1}";
                            workbookXml += $@"<sheet name=""{EscapeXmlAttribute(sheetName)}"" sheetId=""{i + 1}"" r:id=""rId{i + 1}""/>";
                        }
                        workbookXml += @"</sheets><definedNames>";
                        for (int i = 0; i < pages.Count; i++)
                        {
                            string sheetName = $"Page{i + 1}";
                            string escapedSheetName = "'" + sheetName.Replace("'", "''") + "'";
                            workbookXml += $@"<definedName name=""_xlnm.Print_Titles"" localSheetId=""{i}"">{escapedSheetName}!$1:$2</definedName>";
                        }
                        workbookXml += "</definedNames></workbook>";
                        AddTextFileToZip(archive, "xl/workbook.xml", workbookXml);
                        AddTextFileToZip(archive, "xl/styles.xml", @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<fonts count=""6"">
<font><sz val=""11"" /><name val=""Calibri"" /></font>
<font><b /><sz val=""11"" /><name val=""Calibri"" /></font>
<font><b /><sz val=""9"" /><name val=""Arial"" /></font>
<font><sz val=""8"" /><name val=""Arial"" /></font>
<font><b /><sz val=""12"" /><name val=""Arial"" /></font>
<font><sz val=""9"" /><name val=""Arial"" /></font>
</fonts>
<fills count=""2"">
<fill><patternFill patternType=""none""/></fill>
<fill><patternFill patternType=""gray125""/></fill>
</fills>
<borders count=""2"">
<border><left /><right /><top /><bottom /><diagonal /></border>
<border diagonalDown=""false"" diagonalUp=""false"">
<left style=""thin""><color auto=""1"" /></left>
<right style=""thin""><color auto=""1"" /></right>
<top style=""thin""><color auto=""1"" /></top>
<bottom style=""thin""><color auto=""1"" /></bottom>
<diagonal />
</border>
</borders>
<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>
<cellXfs count=""6"">
<xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""center""/></xf>
<xf numFmtId=""0"" fontId=""1"" fillId=""0"" borderId=""0"" xfId=""0"" applyFont=""1"" applyAlignment=""1""><alignment vertical=""center""/></xf>
<xf numFmtId=""0"" fontId=""2"" fillId=""0"" borderId=""1"" xfId=""0"" applyFont=""1"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""center""/></xf>
<xf numFmtId=""0"" fontId=""3"" fillId=""0"" borderId=""1"" xfId=""0"" applyFont=""1"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""center""/></xf>
<xf numFmtId=""0"" fontId=""4"" fillId=""0"" borderId=""0"" xfId=""0"" applyFont=""1"" applyAlignment=""1""><alignment vertical=""center""/></xf>
<xf numFmtId=""0"" fontId=""5"" fillId=""0"" borderId=""0"" xfId=""0"" applyFont=""1"" applyAlignment=""1""><alignment vertical=""center"" horizontal=""right""/></xf>
</cellXfs>
</styleSheet>");
                        for (int i = 0; i < pages.Count; i++)
                        {
                            ProcessCommands(pages[i].ToList(), i + 1, archive);
                            GetDevmode(PageSetupDialog1.PrinterSettings, 2, "");
                            ZipArchiveEntry psEntry = archive.CreateEntry($"xl/printerSettings/printerSettings{i + 1}.bin");
                            using (BinaryWriter writer = new BinaryWriter(psEntry.Open()))
                            {
                                writer.Write(DevModeArray);
                            }
                            string sheetRels = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId{i + 1}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"" Target=""../printerSettings/printerSettings{i + 1}.bin""/>
</Relationships>";
                            AddTextFileToZip(archive, $"xl/worksheets/_rels/sheet{i + 1}.xml.rels", sheetRels);
                        }
                    }
                    using (FileStream fileStream = new FileStream(excelFilePath, FileMode.Create))
                    {
                        ms.Seek(0, SeekOrigin.Begin);
                        ms.CopyTo(fileStream);
                    }
                }
                System.Diagnostics.Process.Start(new ProcessStartInfo { FileName = excelFilePath, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            isExportingToExcel = false;
        }
        private void ProcessCommands(List<string> commands, int pageNumber, ZipArchive archive)
        {
            List<string> drawStrings = commands.Where(cmd => cmd.StartsWith("DrawString")).ToList();
            List<string> drawRectangles = commands.Where(cmd => cmd.StartsWith("DrawRectangle")).ToList();
            if (!drawRectangles.Any())
            {
                return;
            }
            List<float> rowYs = drawRectangles.Select(cmd => float.Parse(cmd.Split('|')[2], CultureInfo.InvariantCulture))
                                             .Distinct()
                                             .OrderBy(y => y)
                                             .ToList();
            List<string> firstRowRects = drawRectangles.Where(cmd => float.Parse(cmd.Split('|')[2], CultureInfo.InvariantCulture) == rowYs.First()).ToList();
            List<float> columnXs = firstRowRects.Select(cmd => float.Parse(cmd.Split('|')[1], CultureInfo.InvariantCulture)).OrderBy(x => x).ToList();
            int colCount = columnXs.Count;
            Dictionary<(int dataRowIndex, int col), (float x, float y, float width, float height)> cellRects = new Dictionary<(int dataRowIndex, int col), (float x, float y, float width, float height)>();
            for (int r = 0; r < rowYs.Count; r++)
            {
                float rowY = rowYs[r];
                List<string> rowRects = drawRectangles.Where(cmd => float.Parse(cmd.Split('|')[2], CultureInfo.InvariantCulture) == rowY)
                                                    .OrderBy(cmd => float.Parse(cmd.Split('|')[1], CultureInfo.InvariantCulture)).ToList();
                for (int c = 0; c < colCount; c++)
                {
                    if (c < rowRects.Count)
                    {
                        string[] parts = rowRects[c].Split('|');
                        float x = float.Parse(parts[1], CultureInfo.InvariantCulture);
                        float y = float.Parse(parts[2], CultureInfo.InvariantCulture);
                        float width = float.Parse(parts[3], CultureInfo.InvariantCulture);
                        float height = float.Parse(parts[4], CultureInfo.InvariantCulture);
                        cellRects[(r, c)] = (x, y, width, height);
                    }
                }
            }
            Dictionary<(int dataRowIndex, int col), string> cellTexts = new Dictionary<(int dataRowIndex, int col), string>();
            string pageNumberText = "";
            string titleText = "";
            float minGridY = rowYs.First();
            foreach (string cmd in drawStrings)
            {
                string[] parts = cmd.Split('|');
                string text = parts[1];
                float tx = float.Parse(parts[5], CultureInfo.InvariantCulture);
                float ty = float.Parse(parts[6], CultureInfo.InvariantCulture);
                if (ty < minGridY)
                {
                    if (text.StartsWith("Page "))
                    {
                        pageNumberText = text;
                    }
                    else
                    {
                        titleText = text;
                    }
                    continue;
                }
                for (int r = 0; r < rowYs.Count; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        if (cellRects.TryGetValue((r, c), out (float x, float y, float width, float height) rect))
                        {
                            if (tx >= rect.x && tx < rect.x + rect.width && ty >= rect.y && ty < rect.y + rect.height)
                            {
                                cellTexts[(r, c)] = text;
                                break;
                            }
                        }
                    }
                }
            }
            StringBuilder worksheet = new StringBuilder();
            worksheet.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            worksheet.AppendLine(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">");
            worksheet.AppendLine("<sheetViews><sheetView workbookViewId=\"0\">");
            worksheet.AppendLine("</sheetView></sheetViews>");
            worksheet.AppendLine("<cols>");
            for (int c = 0; c < colCount; c++)
            {
                float width = float.Parse(firstRowRects[c].Split('|')[3], CultureInfo.InvariantCulture) - 1;
                double excelWidth = Math.Round((width / 72 * 96) / 7.5, 2);
                worksheet.AppendLine($"<col min=\"{c + 1}\" max=\"{c + 1}\" width=\"{excelWidth.ToString(CultureInfo.InvariantCulture)}\" customWidth=\"1\"/>");
            }
            worksheet.AppendLine("</cols>");
            worksheet.AppendLine("<sheetData>");
            int currentExcelRow = 1;
            if (!string.IsNullOrWhiteSpace(titleText) || !string.IsNullOrWhiteSpace(pageNumberText))
            {
                float maxFontSize = drawStrings
                    .Where(cmd => float.Parse(cmd.Split('|')[6], CultureInfo.InvariantCulture) < minGridY)
                    .Select(cmd => float.Parse(cmd.Split('|')[3], CultureInfo.InvariantCulture))
                    .DefaultIfEmpty(12).Max();
                float headerRowHeight = maxFontSize;
                worksheet.AppendLine($"<row r=\"{currentExcelRow}\" spans=\"1:{colCount}\" ht=\"{headerRowHeight}\" customHeight=\"1\">");
                if (!string.IsNullOrWhiteSpace(titleText))
                {
                    worksheet.AppendLine($"<c r=\"A{currentExcelRow}\" s=\"4\" t=\"inlineStr\"><is><t>{EscapeXml(titleText)}</t></is></c>");
                }
                if (!string.IsNullOrWhiteSpace(pageNumberText))
                {
                    string lastColumn = GetExcelColumnName(colCount - 1);
                    worksheet.AppendLine($"<c r=\"{lastColumn}{currentExcelRow}\" s=\"5\" t=\"inlineStr\"><is><t>{EscapeXml(pageNumberText)}</t></is></c>");
                }
                worksheet.AppendLine("</row>");
                currentExcelRow++;
            }
            for (int r = 0; r < rowYs.Count; r++)
            {
                float height = cellRects[(r, 0)].height;
                worksheet.AppendLine($"<row r=\"{currentExcelRow}\" spans=\"1:{colCount}\" ht=\"{height}\" customHeight=\"1\">");
                for (int c = 0; c < colCount; c++)
                {
                    if (cellTexts.TryGetValue((r, c), out string text))
                    {
                        string columnName = GetExcelColumnName(c);
                        string cellRef = columnName + currentExcelRow;
                        string style = r == 0 ? "2" : "3";
                        worksheet.AppendLine($"<c r=\"{cellRef}\" s=\"{style}\" t=\"inlineStr\"><is><t>{EscapeXml(text)}</t></is></c>");
                    }
                }
                worksheet.AppendLine("</row>");
                currentExcelRow++;
            }
            worksheet.AppendLine("</sheetData>");
            double leftMargin = PageSetupDialog1.PageSettings.Margins.Left / 100.0;
            double rightMargin = PageSetupDialog1.PageSettings.Margins.Right / 100.0;
            double topMargin = PageSetupDialog1.PageSettings.Margins.Top / 100.0;
            double bottomMargin = PageSetupDialog1.PageSettings.Margins.Bottom / 100.0;
            string orientation = PageSetupDialog1.PageSettings.Landscape ? "landscape" : "portrait";
            PaperKind kind = PageSetupDialog1.PageSettings.PaperSize.Kind;
            int paperSizeCode = GetExcelPaperSizeCode(kind);
            string paperSizeAttr = $"paperSize=\"{paperSizeCode}\"";
            string paperWidthAttr = "";
            string paperHeightAttr = "";
            if (kind == PaperKind.Custom)
            {
                double widthInPoints = (PageSetupDialog1.PageSettings.PaperSize.Width / 100.0) * 72;
                double heightInPoints = (PageSetupDialog1.PageSettings.PaperSize.Height / 100.0) * 72;
                paperWidthAttr = $"paperWidth=\"{widthInPoints}\"";
                paperHeightAttr = $"paperHeight=\"{heightInPoints}\"";
            }
            List<string> attrs = new List<string> { paperSizeAttr };
            if (!string.IsNullOrEmpty(paperWidthAttr))
            {
                attrs.Add(paperWidthAttr);
            }
            if (!string.IsNullOrEmpty(paperHeightAttr))
            {
                attrs.Add(paperHeightAttr);
            }
            attrs.Add($"orientation=\"{orientation}\"");
            attrs.Add($"r:id=\"rId{pageNumber}\"");
            string pageSetupAttrs = string.Join(" ", attrs);
            worksheet.AppendLine($"<pageMargins left=\"{leftMargin.ToString(CultureInfo.InvariantCulture)}\" right=\"{rightMargin.ToString(CultureInfo.InvariantCulture)}\" top=\"{topMargin.ToString(CultureInfo.InvariantCulture)}\" bottom=\"{bottomMargin.ToString(CultureInfo.InvariantCulture)}\" header=\"0\" footer=\"0\"/>");
            worksheet.AppendLine($"<pageSetup {pageSetupAttrs}/>");
            worksheet.AppendLine("</worksheet>");
            AddTextFileToZip(archive, $"xl/worksheets/sheet{pageNumber}.xml", worksheet.ToString());
        }
        private int GetExcelPaperSizeCode(PaperKind kind)
        {
            switch (kind)
            {
                case PaperKind.Letter: return 1;
                case PaperKind.LetterSmall: return 2;
                case PaperKind.Tabloid: return 3;
                case PaperKind.Ledger: return 4;
                case PaperKind.Legal: return 5;
                case PaperKind.Statement: return 6;
                case PaperKind.Executive: return 7;
                case PaperKind.A3: return 8;
                case PaperKind.A4: return 9;
                case PaperKind.A4Small: return 10;
                case PaperKind.A5: return 11;
                case PaperKind.B4: return 12;
                case PaperKind.B5: return 13;
                case PaperKind.Folio: return 14;
                case PaperKind.Quarto: return 15;
                case PaperKind.Standard10x14: return 16;
                case PaperKind.Standard11x17: return 17;
                case PaperKind.Note: return 18;
                case PaperKind.Number9Envelope: return 19;
                case PaperKind.Number10Envelope: return 20;
                case PaperKind.Number11Envelope: return 21;
                case PaperKind.Number12Envelope: return 22;
                case PaperKind.Number14Envelope: return 23;
                case PaperKind.CSheet: return 24;
                case PaperKind.DSheet: return 25;
                case PaperKind.ESheet: return 26;
                case PaperKind.DLEnvelope: return 27;
                case PaperKind.C5Envelope: return 28;
                case PaperKind.C3Envelope: return 29;
                case PaperKind.C4Envelope: return 30;
                case PaperKind.C6Envelope: return 31;
                case PaperKind.C65Envelope: return 32;
                case PaperKind.B4Envelope: return 33;
                case PaperKind.B5Envelope: return 34;
                case PaperKind.B6Envelope: return 35;
                case PaperKind.ItalyEnvelope: return 36;
                case PaperKind.MonarchEnvelope: return 37;
                case PaperKind.PersonalEnvelope: return 38;
                case PaperKind.USStandardFanfold: return 39;
                case PaperKind.GermanStandardFanfold: return 40;
                case PaperKind.GermanLegalFanfold: return 41;
                case PaperKind.Custom: return 256;
                default: return 256;
            }
        }
        private string EscapeXml(string input)
        {
            return System.Security.SecurityElement.Escape(input) ?? "";
        }
        private string EscapeXmlAttribute(string input)
        {
            return System.Security.SecurityElement.Escape(input) ?? "";
        }
        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            int dividend = columnNumber + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = (char)('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
        private void AddTextFileToZip(ZipArchive archive, string entryName, string content)
        {
            ZipArchiveEntry entry = archive.CreateEntry(entryName);
            using (StreamWriter writer = new StreamWriter(entry.Open()))
            {
                writer.Write(content);
            }
        }
        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesW", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter, [MarshalAs(UnmanagedType.LPWStr)] string pDeviceNameg, IntPtr pDevModeOutput, IntPtr pDevModeInput, int fMode);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GlobalFree(IntPtr handle);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GlobalLock(IntPtr handle);
        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GlobalUnlock(IntPtr handle);
        private void GetDevmode(PrinterSettings printerSettings, int mode, string Filename)
        {
            IntPtr hDevMode = IntPtr.Zero;
            IntPtr hwnd = Handle;
            try
            {
                hDevMode = printerSettings.GetHdevmode(printerSettings.DefaultPageSettings);
                IntPtr pDevMode = GlobalLock(hDevMode);
                int sizeNeeded = DocumentProperties(hwnd, IntPtr.Zero, printerSettings.PrinterName, IntPtr.Zero, pDevMode, 0);
                if (sizeNeeded <= 0)
                {
                    MessageBox.Show("Devmode Bummer, Can't get size of devmode structure");
                    GlobalUnlock(hDevMode);
                    GlobalFree(hDevMode);
                    return;
                }
                DevModeArray = new byte[sizeNeeded];
                if (mode == 1)
                {
                    FileStream fs = new FileStream(Filename, FileMode.Create);
                    for (int i = 0; i < sizeNeeded; ++i)
                    {
                        fs.WriteByte(Marshal.ReadByte(pDevMode, i));
                    }
                    fs.Close();
                    fs.Dispose();
                }
                if (mode == 2)
                {
                    for (int i = 0; i < sizeNeeded; ++i)
                    {
                        DevModeArray[i] = Marshal.ReadByte(pDevMode, i);
                    }
                }
                GlobalUnlock(hDevMode);
                GlobalFree(hDevMode);
            }
            catch (Exception)
            {
                if (hDevMode != IntPtr.Zero)
                {
                    GlobalUnlock(hDevMode);
                    GlobalFree(hDevMode);
                }
            }
        }
        #endregion
        #region PDF Object Class
        public class PdfObject
        {
            public PdfObject(int number, string content, byte[] streamData = null, string streamFooter = "")
            {
                Number = number;
                Content = content;
                StreamData = streamData;
                StreamFooter = streamFooter;
            }
            public int Number { get; set; }
            public string Content { get; set; }
            public byte[] StreamData { get; set; }
            public string StreamFooter { get; set; }
            public override string ToString()
            {
                if (StreamData == null)
                {
                    return $"{Number} 0 obj\r\n{Content}\r\nendobj\r\n";
                }
                return $"{Number} 0 obj\r\n{Content}{StreamFooter}\r\nendobj\r\n";
            }
        }
        #endregion
    }
}