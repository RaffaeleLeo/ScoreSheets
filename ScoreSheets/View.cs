using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScoreSheets
{
    public partial class View : Form, IView
    {

        public event Action SelectFilePressed;
        public event Action RegionalsCheckPressed;

        public View()
        {
            InitializeComponent();
        }

        private void Model_Load(object sender, EventArgs e)
        {

        }

        private void SelectFile_Click(object sender, EventArgs e)
        {
            SelectFilePressed?.Invoke();
        }

        private void RegionalCheck_MouseClick(object sender, MouseEventArgs e)
        {
            RegionalsCheckPressed?.Invoke();
        }
    }
}
