using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections;
using System.Windows.Forms;
using CSharpeLibrary;

namespace ATPPT
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            PPT ppt = new PPT();
            ppt.PPTAuto("d://11.ppt", 2);
            
        }

        

        

    }
}
