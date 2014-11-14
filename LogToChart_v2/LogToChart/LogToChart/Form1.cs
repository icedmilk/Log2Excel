using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Microsoft.Office.Interop.Excel; 

namespace LogToChart
{

    public partial class Form1 : Form
    {
        List<ResultData> resultData = new List<ResultData>();
        int interval = 30;
        
        
        public Form1()
        {         
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();

            int[] a = new int[9] { 1, 2, 3, 4, 5, 10, 15, 30, 60 };
            foreach (int i in a)
                comboBoxInterval.Items.Add((object)i);
            comboBoxInterval.SelectedIndex = 6;
            comboBoxInterval.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            resultData.Clear();
            interval = Convert.ToInt32(comboBoxInterval.SelectedItem);

            List<CPUutilization> list = new List<CPUutilization>();
            List<String> convertString = new List<String>();
            String path = textBox1.Text;
            try
            {
                StreamReader sr = new StreamReader(path, Encoding.Default);
                string s;
                while ((s = sr.ReadLine()) != null)
                {
                    convertString.Add(s);
                }

                foreach (String i in convertString)
                {
                    if (i != "\r" && i != "" && i.Substring(1, 1) != "(")
                    {
                        String[] data = i.Split(',');
                        String[] dataTime = data[0].Trim('\"').Split(' ');
                        String[] hms = dataTime[1].Split(':');
                        int hour = Convert.ToInt32(hms[0]);
                        int minute = Convert.ToInt32(hms[1]);
                        double second = Convert.ToDouble(hms[2]);
                        double util = Convert.ToDouble(data[1].TrimEnd('\r').Trim('\"'));

                        list.Add(new CPUutilization(util, hour, minute, second));
                    }
                }

#region interval
                foreach (CPUutilization cpu in list)
                {

                    if (resultData.Count == 0)
                    {
                        resultData.Add(new ResultData(cpu.hour, cpu.minute / interval * interval));
                        resultData[resultData.Count - 1].InsertData(cpu.util);

                    }
                    else if (resultData[resultData.Count - 1].hour != cpu.hour || resultData[resultData.Count - 1].minute != cpu.minute / interval * interval)
                    {
                        resultData[resultData.Count - 1].Calc();
                        resultData.Add(new ResultData(cpu.hour, cpu.minute / interval * interval));
                        resultData[resultData.Count - 1].InsertData(cpu.util);

                    }
                    else
                    {
                        resultData[resultData.Count - 1].InsertData(cpu.util);
                    }
                }

                resultData[resultData.Count - 1].Calc();
#endregion
                //foreach (ResultData cpu in resultData)
                //    MessageBox.Show(cpu.hour + " " + cpu.utilAvg);
                TestExcel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sorry, something unexpected happens: \n" + ex);
            }
        }

        public void ToTXT()
        {
            String destPath = textBox2.Text + "result.txt";
            // File.Delete(bFilePath); 
            FileStream fsMyfile = new FileStream(destPath, FileMode.Append, FileAccess.Write);
            StreamWriter swMyfile = new StreamWriter(fsMyfile);
            foreach (ResultData cpu in resultData)
                swMyfile.Write(cpu.hour + " " + cpu.utilAvg + "\r\n");
            
            swMyfile.Flush();
            swMyfile.Close();
            fsMyfile.Close();

            MessageBox.Show("Done");
        }

        public void TestExcel()
        {
            _Application myExcel = null;
            _Workbook myBook = null;
            _Worksheet mySheet = null;
            Range myRange = null;

            myExcel = new Microsoft.Office.Interop.Excel.Application();
            myExcel.Workbooks.Add(true);
            myExcel.DisplayAlerts = false;
            myExcel.Visible = true;

            myBook = myExcel.Workbooks[1];
            myBook.Activate();
            mySheet = (_Worksheet)myBook.Worksheets[1];
            mySheet.Name = "Cells";
            mySheet.Activate();


            myRange = mySheet.get_Range("A1", Type.Missing);
            myRange.Value2 = "Time(h)";
            myRange = mySheet.get_Range("B1", Type.Missing);
            myRange.Value2 = "Utilization";
            int i = 2;
            int j = 2;

            foreach (ResultData rd in resultData)
            {
                myRange = mySheet.get_Range("A" + i++, Type.Missing);
                myRange.Value2 = rd.hour + ":" + rd.minute;
                myRange = mySheet.get_Range("B" + j++, Type.Missing);
                myRange.Value2 = rd.utilAvg;
            }

            string filePath = textBox2.Text + "result.xls" + (radioButtonXlsx.Checked ? "x" : "");

            myBook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //myBook.Close(false, Type.Missing, Type.Missing);
            //myExcel.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);

            GC.Collect();
            MessageBox.Show("Done");
        }


    }
}
