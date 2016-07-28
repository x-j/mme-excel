using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn1 {
    public partial class PageNameDialog : Form {

        public string PageName { get; private set; }
        public int PostCount { get; private set; }

        public PageNameDialog() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            if(textBox1.Text.Trim().Length > 0) {
                PageName = textBox1.Text.Trim();
                PostCount = (int)numericUpDown1.Value;
                DialogResult = DialogResult.OK;
                Dispose();
            }
        }

        internal string GetPageName() {
            return PageName;
        }

        internal int GetPostCount() {
            return PostCount;
        }
    }
}
