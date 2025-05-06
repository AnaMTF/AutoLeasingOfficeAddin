using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPlanFinanciar
{
    public partial class FormPlataAvans : Form
    {
        public double Avans { get; private set; } = 0;
        public FormPlataAvans()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if(double.TryParse(txtAvans.Text, out double avans) && avans > 0)
            {
                Avans = avans;
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Introduceţi o sumă validă în lei.");
                DialogResult = DialogResult.None;
            }
        }
    }
}
