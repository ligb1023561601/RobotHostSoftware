using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Robot_HostSoft
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            this.BackgroundImage = Image.FromFile(@"123.jpg");
        }
    }
}
