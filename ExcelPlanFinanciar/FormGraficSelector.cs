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
    public partial class FormGraficSelector : Form
    {
        public enum TipGrafic
        {
            Linie,
            StructuraPct,
            Combo
        }

        public TipGrafic AlegereGrafic { get; private set; }

        public FormGraficSelector()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (radioLinie.Checked) AlegereGrafic = TipGrafic.Linie;
            else if (radioStructura.Checked) AlegereGrafic = TipGrafic.StructuraPct;
            else if (radioCombo.Checked) AlegereGrafic = TipGrafic.Combo;

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}

