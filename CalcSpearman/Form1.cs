using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using MathNet.Numerics.Statistics;
using System.Reflection;
using System.Collections;

namespace CalcSpearman
{

    public partial class Form1 : Form
    {
        int col;
        int row;
        int ProjectNum;
        FolderBrowserDialog folder = new FolderBrowserDialog();
        static int MaxRow = 2000000;
        static int MaxLimit = 2000000;
        List<List<object>> allpValue = NewListMatrixOfObject(MaxRow, 6);//输出矩阵的列数
        string OutputFolderPath = "";
        Microsoft.Office.Interop.Excel.Application app = null;
        string OutputFilePathToOpen = "";//用于记录要打开的文件路径
        

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folder.ShowDialog();
            textBox1.Text = folder.SelectedPath;
        }
        public void ReadExcel(string FilePath,out List<List<object>> InputMatrix,out int RowNum)
        {
            Microsoft.Office.Interop.Excel.Workbook workBook = null;
            app = new Microsoft.Office.Interop.Excel.Application();
            workBook = app.Workbooks.Open(FilePath);
            Worksheet worksheet = (Worksheet)workBook.Worksheets[2];//选择sheet
            col = worksheet.UsedRange.CurrentRegion.Columns.Count;
            row = worksheet.UsedRange.CurrentRegion.Rows.Count;
            InputMatrix = NewListMatrixOfObject(row, col);
            object[,] current;
            current = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row, col]].Value2;

            int k = 1;
            for (int i = 0; i <= row-1;i++)
            {
                if (current[i + 1,21]!=null)//读取的excel列数
                {
                    for (int j = 1; j <= col; j++)
                    {
                        InputMatrix[k - 1][j - 1] = current[i + 1, j];
                    }
                    k++;
                }
            }
            RowNum = k-1;
            app.Quit();
            app = null;
            workBook = null;
        }
        public static List<List<object>> NewListMatrixOfObject(int iRowCount, int iColCount)
        {
            List<List<object>> matrix = new List<List<object>>(iRowCount);
            for (int i = 0; i < iRowCount; i++)
            {
                matrix.Add(new List<object>(iColCount));
                for (int j = 0; j < iColCount; j++)
                {
                    matrix[i].Add("");
                }
            }
            return matrix;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folder.ShowDialog();
            textBox2.Text = folder.SelectedPath;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            label6.ForeColor = Color.Green;
            label6.Text = "Running";

            string RootPath = textBox1.Text;
            ProjectNum = 0;
            OutputFolderPath = textBox2.Text;
            DirectoryInfo root = new DirectoryInfo(RootPath);
            int index = -1;
            string subStr = "";
            allpValue[0][0] = "Project";
            allpValue[0][1] = "#MPLC";
            allpValue[0][2] = "MeanEntropyMPLC";
            allpValue[0][3] = "#Non-MPLC";
            allpValue[0][4] = "MeanEntropyNotMPLC";
            allpValue[0][5] = "P-Value";
            foreach (DirectoryInfo nextFolder in root.GetDirectories())
            {
                FileInfo[] files = nextFolder.GetFiles();
                int num = files.Length;
                FileInfo LastFile = files[num - 1];
                string FileName = LastFile.FullName;
                ProjectNum++;
                index = FileName.LastIndexOf(".");
                subStr = FileName.Substring(index + 1);
                if (subStr != "xlsx")
                {
                    label2.ForeColor = Color.Red;
                    label2.Text = "存在非Excel";
                    return;
                }
                index = FileName.IndexOf("~");
                if (index > 0)
                {
                    string str = FileName.Remove(index, 2);
                    FileName = str;
                }
                bool succ=CalcCorrelation(FileName,ProjectNum);
                if (succ == false)
                {
                    label6.Text = "界限输入有问题";
                    return;
                }
            }
            WriteExcel();
            label6.Text = "success";
        }

        public bool CalcCorrelation(string FileName,int ProjectNum)
        {
            List<List<object>> InputMatrix;
            int RowNum;
            ReadExcel(FileName, out InputMatrix, out RowNum);
           
            double meanEnMPLC = 0;
            double meanEnNotMPLC = 0;
            
            ArrayList arrMPLC = new ArrayList();
            ArrayList arrNotMPLC = new ArrayList();
            int iMPLC = 0;
            int iNotMPLC = 0;
            for(int i=1; i < RowNum; i++)
            {
                if (InputMatrix[i][7].ToString() == "0")
                {
                    iNotMPLC++;
                    meanEnNotMPLC += (double)InputMatrix[i][20];
                    arrNotMPLC.Add(InputMatrix[i][20]);
                }
                else
                {
                    iMPLC++;
                    meanEnMPLC += (double)InputMatrix[i][20];
                    arrMPLC.Add(InputMatrix[i][20]);
                }
            }
            meanEnNotMPLC /= iNotMPLC;
            meanEnMPLC /= iMPLC;
            var dataMPLCEn = new double[iMPLC];
            var dataNotMPLCEn = new double[iNotMPLC];
            for(int i=0; i < iNotMPLC; i++)
            {
                dataNotMPLCEn[i] = (double)arrNotMPLC[i];
            }
            for(int i=0; i < iMPLC; i++)
            {
                dataMPLCEn[i] = (double)arrMPLC[i];
            }
            alglib.mannwhitneyutest(dataNotMPLCEn, iNotMPLC, dataMPLCEn, iMPLC, out double p_value, out double ltail, out double rtail);


            allpValue[ProjectNum][1] = iMPLC;
            allpValue[ProjectNum][2] = meanEnMPLC;
            allpValue[ProjectNum][3] = iNotMPLC;
            allpValue[ProjectNum][4] = meanEnNotMPLC;
            allpValue[ProjectNum][5] = p_value;

            string strProjectName = FileName.Substring(FileName.LastIndexOf("\\") + 1, FileName.Length - FileName.LastIndexOf("\\") - 1);
            strProjectName = strProjectName.Substring(0, strProjectName.IndexOf("_"));
            allpValue[ProjectNum][0] = strProjectName;

            return true;
        }

        public bool WriteExcel()
        {
            int column = 6;//列数
            app = new Microsoft.Office.Interop.Excel.Application();

            string OutputFilePath = OutputFolderPath + "\\" + "P_Value"+"_"+"Entropy"+"_"+DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            OutputFilePathToOpen = OutputFilePath;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = app.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet currentSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
            object[,] arrBatchContent = new object[ProjectNum + 1, column];
            currentSheet.Name = "Result";
            for(int i=0; i < ProjectNum + 1; i++)
            {
                for(int j=0; j < column; j++)
                {
                    arrBatchContent[i, j] = allpValue[i][j];
                }
            }
            currentSheet.Range[currentSheet.Cells[1, 1], currentSheet.Cells[ProjectNum + 1, column]].value2 = arrBatchContent;
            currentSheet.Range[currentSheet.Cells[1, 1], currentSheet.Cells[ProjectNum + 1, column]].EntireColumn.AutoFit();
            currentSheet.Range[currentSheet.Cells[2, 6], currentSheet.Cells[ProjectNum + 1, 6]].NumberFormat = "#,###0.000";
            arrBatchContent = null;
            workbook.SaveCopyAs(OutputFilePath);
            workbook.Close(false, Missing.Value, Missing.Value);
            app.Quit();            
            return true;
        }

        private void buttonOpen_Click(object sender, EventArgs e)
        {
            if (OutputFilePathToOpen == null || OutputFilePathToOpen.Length == 0) return;
            if(OutputFilePathToOpen.EndsWith(".xlsx")) System.Diagnostics.Process.Start(OutputFilePathToOpen);
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }
    }
 
}
