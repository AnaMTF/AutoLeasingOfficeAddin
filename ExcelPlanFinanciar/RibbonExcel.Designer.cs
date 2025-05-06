namespace ExcelPlanFinanciar
{
    partial class RibbonExcel : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonExcel()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnVizualizareGrafic = this.Factory.CreateRibbonButton();
            this.btnPaymentPlan = this.Factory.CreateRibbonButton();
            this.btnPlataInAvans = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnPaymentPlan);
            this.group1.Items.Add(this.btnVizualizareGrafic);
            this.group1.Items.Add(this.btnPlataInAvans);
            this.group1.Label = "Analize";
            this.group1.Name = "group1";
            // 
            // btnVizualizareGrafic
            // 
            this.btnVizualizareGrafic.Label = "Vizualizati Graficul";
            this.btnVizualizareGrafic.Name = "btnVizualizareGrafic";
            this.btnVizualizareGrafic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVizualizareGrafic_Click);
            // 
            // btnPaymentPlan
            // 
            this.btnPaymentPlan.Label = "Vezi Planul de Plata";
            this.btnPaymentPlan.Name = "btnPaymentPlan";
            this.btnPaymentPlan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPaymentPlan_Click);
            // 
            // btnPlataInAvans
            // 
            this.btnPlataInAvans.Label = "Plata in Avans";
            this.btnPlataInAvans.Name = "btnPlataInAvans";
            this.btnPlataInAvans.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPlataInAvans_Click);
            // 
            // RibbonExcel
            // 
            this.Name = "RibbonExcel";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonExcel_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVizualizareGrafic;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPaymentPlan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlataInAvans;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonExcel RibbonExcel
        {
            get { return this.GetRibbon<RibbonExcel>(); }
        }
    }
}
