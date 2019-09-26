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
using Exel = Microsoft.Office.Interop.Excel;

namespace Form_Datagridden_exele_aktar
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Exel.Application exeldosya = new Exel.Application();
            exeldosya.Visible = true;
            object Missing = Type.Missing;
            Workbook calismakitabi = exeldosya.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)calismakitabi.Sheets[1];
            int satir = 1;
            int sutun = 1;

            for(int j=0; j<dataGridView1.Columns.Count; j++)
            {
                Range myrange = (Range)sheet1.Cells[satir, sutun + j];
                myrange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            satir++;
            for(int i=0; i<dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Range myrange = (Range)sheet1.Cells[satir + i, sutun + j];
                    myrange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myrange.Select();
                }   
            }

        }
    }
}
