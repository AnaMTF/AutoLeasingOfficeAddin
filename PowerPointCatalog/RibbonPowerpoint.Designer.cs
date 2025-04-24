namespace PowerPointCatalog
{
    partial class RibbonPowerpoint : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonPowerpoint()
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
            this.groupOferta = this.Factory.CreateRibbonGroup();
            this.btnOferta = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupOferta.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupOferta);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupOferta
            // 
            this.groupOferta.Items.Add(this.btnOferta);
            this.groupOferta.Label = "Oferta";
            this.groupOferta.Name = "groupOferta";
            // 
            // btnOferta
            // 
            this.btnOferta.Label = "Selecteaza Oferta";
            this.btnOferta.Name = "btnOferta";
            this.btnOferta.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOferta_Click);
            // 
            // RibbonPowerpoint
            // 
            this.Name = "RibbonPowerpoint";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonPowerpoint_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupOferta.ResumeLayout(false);
            this.groupOferta.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupOferta;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOferta;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonPowerpoint RibbonPowerpoint
        {
            get { return this.GetRibbon<RibbonPowerpoint>(); }
        }
    }
}
