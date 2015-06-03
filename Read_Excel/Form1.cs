using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace Read_Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string strCon = @" Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = C:\Documents and Settings\xlyang\桌面\项目文档\COC\coc.xls ;Extended Properties=Excel 8.0";
            OleDbConnection myConn = new OleDbConnection(strCon);
            string strCom = " SELECT * FROM [A$] ";
            myConn.Open();
            DataSet ds=new DataSet();
            OleDbCommand oledb=new OleDbCommand(strCom,myConn);
            OleDbDataAdapter odda =new OleDbDataAdapter(oledb);
            odda.Fill(ds);
            OleDbDataReader oddr = oledb.ExecuteReader();
            while (oddr.Read())
            {
                if (oddr[0].ToString() == "Growth Method")

                {
                    label1.Text = oddr[2].ToString();
                }



            }
            myConn.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Dodrop dd=new Dodrop();
            dd.ShowDialog();
        }
    }
}
