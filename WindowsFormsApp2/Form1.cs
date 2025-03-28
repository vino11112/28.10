using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        Excel.Worksheet xlSht;
        private List<string> dataList = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Объявляем приложение
            Excel.Application app = new Excel.Application
            {
                //Отобразить Excel
                Visible = true,
                //Количество листов в рабочей книге
                SheetsInNewWorkbook = 2
            };
            
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
          
            Excel.Workbook xlWB;
           
   
            xlWB = app.Workbooks.Open(@"C:\vino\Книга1.xlsx");
            xlSht = xlWB.ActiveSheet;
            if (dataList.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта.");
                return;
            }
            for (int i = 0; i < dataList.Count; i++)
            {
                xlSht.Cells[i + 1, 1] = dataList[i];
            }
            /* app.Workbooks.Open(@"C:\vino\Книга1.xls",
   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
   Type.Missing, Type.Missing);*/
            app.Application.ActiveWorkbook.SaveAs("Книга1.xlsx", Type.Missing,
   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string data = textBox1.Text;
            if (!string.IsNullOrWhiteSpace(data))
            {
                dataList.Add(data);
                listBox1.Items.Add(data);
                textBox1.Clear();
            }
            else
            {
                MessageBox.Show("Введите данные для добавления.");
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            xlSht.Cells[6, 5].FormulaLocal = "=СУММ(A1;A2)";
        }
    }
    
}
