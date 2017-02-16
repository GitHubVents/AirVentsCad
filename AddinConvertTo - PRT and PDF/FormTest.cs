using AddinConvertTo.Classes;
 
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddinConvertTo
{
    public partial class FormTest : Form
    {
        public FormTest()
        {
            InitializeComponent();
        }
        public List<FilesData.TaskParam> listForm { get; set; }

        private void FormTest_Load(object sender, EventArgs e)
        {
            try
            {
                dgv.DataSource = listForm;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
