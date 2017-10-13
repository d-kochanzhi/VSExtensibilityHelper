using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSIXProject_Editor.Controls.WinForms
{
    public partial class WinFormUserControl : UserControl
    {
        public WinFormUserControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            VSExtensibilityHelper.Core.Service.MsgBoxService.Info("Hellow World");
        }
    }
}
