using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace Read_Excel
{
    public partial class Dodrop : Form
    {
        public Dodrop()
        {
            InitializeComponent();
        }

        private int btn_x = 0;
        private int btn_y = 0;
       

        private void Dodrop_Load(object sender, EventArgs e)
        {
            
            button1.MouseMove -= new MouseEventHandler(button1_MouseMove);
            label1.Text = button1.Location.X.ToString() +","+ button1.Location.Y.ToString();

        }



        private void button1_MouseDown_1(object sender, MouseEventArgs e)
        {
            btn_x = e.X;
            btn_y = e.Y;

            button1.MouseMove += new MouseEventHandler(button1_MouseMove);
            label2.Text = e.X + "," + e.Y;
        }

        private void button1_MouseMove(object sender, MouseEventArgs e)
        {
            label2.Text = e.X + "," + e.Y;
          
           button1.Location=new Point(button1.Location.X+e.X-btn_x,button1.Location.Y+e.Y-btn_y);
        }

        private void button1_MouseUp(object sender, MouseEventArgs e)
        {
            button1.MouseMove-=new MouseEventHandler(button1_MouseMove);
        }
      

    }
}
