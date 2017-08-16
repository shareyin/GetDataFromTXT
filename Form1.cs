using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;


namespace HanderTxt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text;
            ReadText(path, 0);
            
        }

        public void ReadText(string path,int a)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("站点名", typeof(String));
            dt.Columns.Add("信号机", typeof(String));
            dt.Columns.Add("信号闭塞", typeof(String));
            dt.Columns.Add("制式", typeof(String));
            dt.Columns.Add("公里标", typeof(String));
            dt.Columns.Add("距离", typeof(String));
            dt.Columns.Add("分区", typeof(String));
            dt.Columns.Add("分区限速", typeof(String));
            dt.Columns.Add("黄灯", typeof(String));
            dt.Columns.Add("黄灯限速", typeof(String));
            dt.Columns.Add("校正", typeof(String));
            dt.Columns.Add("有权", typeof(String));
            string[] str = new string[1024];
            ExcelCreate excelc = new ExcelCreate();
            try 
            {
                using (StreamReader sr = new StreamReader(path, GetEncoding(path, UnicodeEncoding.Default)))
                {
                    string line;
                    int index1 = 0, index2 = 0, index3 = 0;
                    string lineState="";
                    string lineSignal="";
                    string lineBisai = "";
                    string lineZhishi = "";
                    string lineGonglibiao = "";
                    string lineJuli = "";
                    string lineFenqu = "";
                    string lineFenxian = "";
                    string lineHuangdeng = "";
                    string lineHuangxian = "";
                    string lineJiaozheng = "";
                    string lineYouquan = "";
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.IndexOf("闭") > 0)
                        {
                            index2 = line.Length;
                            index1 = line.IndexOf("]");
                            index3 = line.IndexOf("闭");
                            lineYouquan = line.Substring(index2 - 2, 2);
                            lineJiaozheng = line.Substring(index2 - 5, 2);
                            lineHuangxian = line.Substring(index2 - 11, 3);
                            lineHuangdeng = line.Substring(index2 - 15, 3);
                            lineFenxian = line.Substring(index2 - 19, 3);
                            lineFenqu = line.Substring(index2 - 23, 3);
                            lineGonglibiao = line.Substring(index2 - 49, 8);
                            lineZhishi = line.Substring(index2 - 52, 2);
                            lineBisai = line.Substring(index3 - 1, 2);
                            lineSignal = line.Substring(index3 - 9, 7);
                            lineState = line.Substring(2, 10);
                            DataRow dr = dt.NewRow();
                            dr[0] = lineState;
                            dr[1] = lineSignal;
                            dr[2] = lineBisai;
                            dr[3] = lineZhishi;
                            dr[4] = lineGonglibiao;
                            dr[5] = lineJuli;
                            dr[6] = lineFenqu;
                            dr[7] = lineFenxian;
                            dr[8] = lineHuangdeng;
                            dr[9] = lineHuangxian;
                            dr[10] = lineJiaozheng;
                            dr[11] = lineYouquan;
                            dt.Rows.Add(dr);
                        }
                    }
                }
                string OutPath = string.Format("{0}\\{1}", Application.StartupPath, DateTime.Now.ToString("yyyyMMddHHmmss")+"OUT.xls");
                excelc.OutFileToDisk(dt, "导出数据表", OutPath);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        //protected void ExportExcel(DataTable dt)
        //{
        //    if (dt == null || dt.Rows.Count == 0) return;
        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

        //    if (xlApp == null)
        //    {
        //        return;
        //    }
        //    System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
        //    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        //    Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
        //    Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
        //    Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range range;
        //    long totalCount = dt.Rows.Count;
        //    long rowRead = 0;
        //    float percent = 0;
        //    for (int i = 0; i < dt.Columns.Count; i++)
        //    {
        //        worksheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
        //        range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, i + 1];
        //        range.Interior.ColorIndex = 15;
        //        range.Font.Bold = true;
        //    }
        //    for (int r = 0; r < dt.Rows.Count; r++)
        //    {
        //        for (int i = 0; i < dt.Columns.Count; i++)
        //        {
        //            worksheet.Cells[r + 2, i + 1] = dt.Rows[r][i].ToString();
        //        }
        //        rowRead++;
        //        percent = ((float)(100 * rowRead)) / totalCount;
        //    }
        //    xlApp.Visible = true;
        //}
        public Encoding GetEncoding(string fileName, Encoding defualencoding)
        {
            FileStream fs = new FileStream(fileName, FileMode.Open);
            Encoding tagetEncoding = GetEncoding(fs, defualencoding);
            fs.Close();
            return tagetEncoding;
        }

        public Encoding GetEncoding(FileStream stream, Encoding defualEncoding)
        {
            Encoding tagetEncoding = defualEncoding;
            if (stream != null && stream.Length >= 2)
            {
                byte byte1 = 0;
                byte byte2 = 0;
                byte byte3 = 0;
                byte byte4 = 0;
                long origPos = stream.Seek(0, SeekOrigin.Begin);
                stream.Seek(0, SeekOrigin.Begin);

                int nByte = stream.ReadByte();
                byte1 = Convert.ToByte(nByte);
                byte2 = Convert.ToByte(stream.ReadByte());
                if (stream.Length >= 3)
                {
                    byte3 = Convert.ToByte(stream.ReadByte());
                }
                if (stream.Length >= 4)
                {
                    byte4 = Convert.ToByte(stream.ReadByte());
                }
                if (byte1 == 0xFE && byte2 == 0xFF)
                {
                    tagetEncoding = Encoding.BigEndianUnicode;
                }
                if (byte1 == 0xff && byte2 == 0xfe && byte3 != 0xff)
                {
                    tagetEncoding = Encoding.Unicode;
                }
                if (byte1 == 0xef && byte2 == 0xbb && byte3 == 0xbf)
                {
                    tagetEncoding = Encoding.UTF8;
                }
                stream.Seek(origPos, SeekOrigin.Begin);
            }
            return tagetEncoding;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog filedialog = new OpenFileDialog();
            string FileName = "";
            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                FileName = filedialog.FileName;
                textBox1.Text = FileName;
            }
        }
    }
}
