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
            this.grupAnalizeFinanciare = this.Factory.CreateRibbonGroup();
            this.btnPaymentPlan = this.Factory.CreateRibbonButton();
            this.btnVizualizareGrafic = this.Factory.CreateRibbonButton();
            this.btnPlataInAvans = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grupAnalizeFinanciare.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grupAnalizeFinanciare);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grupAnalizeFinanciare
            // 
            this.grupAnalizeFinanciare.Items.Add(this.btnPaymentPlan);
            this.grupAnalizeFinanciare.Items.Add(this.btnVizualizareGrafic);
            this.grupAnalizeFinanciare.Items.Add(this.btnPlataInAvans);
            this.grupAnalizeFinanciare.Label = "Analize Leasing";
            this.grupAnalizeFinanciare.Name = "grupAnalizeFinanciare";
            // 
            // btnPaymentPlan
            // 
            this.btnPaymentPlan.Image = global::ExcelPlanFinanciar.Properties.Resources._1694683_small;
            this.btnPaymentPlan.Label = "Vedeti Planul de Plata";
            this.btnPaymentPlan.Name = "btnPaymentPlan";
            this.btnPaymentPlan.ShowImage = true;
            this.btnPaymentPlan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPaymentPlan_Click);
            // 
            // btnVizualizareGrafic
            // 
            this.btnVizualizareGrafic.Image = global::ExcelPlanFinanciar.Properties.Resources.Graph_PNG_High_Quality_Image;
            this.btnVizualizareGrafic.Label = "Vizualizati Graficul";
            this.btnVizualizareGrafic.Name = "btnVizualizareGrafic";
            this.btnVizualizareGrafic.ShowImage = true;
            this.btnVizualizareGrafic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVizualizareGrafic_Click);
            // 
            // btnPlataInAvans
            // 
            this.btnPlataInAvans.Image = global::ExcelPlanFinanciar.Properties.Resources.Green_Dollar_Sign_PNG_Clipart;
            this.btnPlataInAvans.Label = "Plata in Avans";
            this.btnPlataInAvans.Name = "btnPlataInAvans";
            this.btnPlataInAvans.ShowImage = true;
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
            this.grupAnalizeFinanciare.ResumeLayout(false);
            this.grupAnalizeFinanciare.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grupAnalizeFinanciare;
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
