namespace PMADCCExcel
{
    partial class AddinRibbonComponent : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddinRibbonComponent()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinRibbonComponent));
            this.pmadccTab = this.Factory.CreateRibbonTab();
            this.excelPMADCCGroup = this.Factory.CreateRibbonGroup();
            this.menuPMADCCOperations = this.Factory.CreateRibbonMenu();
            this.importPMADCCCommand = this.Factory.CreateRibbonButton();
            this.exportPMADCCCommand = this.Factory.CreateRibbonButton();
            this.resetPMADCCCommand = this.Factory.CreateRibbonButton();
            this.menuShowDiagram = this.Factory.CreateRibbonMenu();
            this.showOriginalDiagramButton = this.Factory.CreateRibbonButton();
            this.showWorkDiagramButton = this.Factory.CreateRibbonButton();
            this.savePmadccFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.openPmadccFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.pmadccTab.SuspendLayout();
            this.excelPMADCCGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // pmadccTab
            // 
            this.pmadccTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.pmadccTab.ControlId.OfficeId = "TabFormulas";
            this.pmadccTab.Groups.Add(this.excelPMADCCGroup);
            this.pmadccTab.Label = "TabFormulas";
            this.pmadccTab.Name = "pmadccTab";
            // 
            // excelPMADCCGroup
            // 
            this.excelPMADCCGroup.Items.Add(this.menuPMADCCOperations);
            this.excelPMADCCGroup.Items.Add(this.menuShowDiagram);
            this.excelPMADCCGroup.Label = "PMADCC for Excel";
            this.excelPMADCCGroup.Name = "excelPMADCCGroup";
            // 
            // menuPMADCCOperations
            // 
            this.menuPMADCCOperations.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuPMADCCOperations.Image = ((System.Drawing.Image)(resources.GetObject("menuPMADCCOperations.Image")));
            this.menuPMADCCOperations.Items.Add(this.importPMADCCCommand);
            this.menuPMADCCOperations.Items.Add(this.exportPMADCCCommand);
            this.menuPMADCCOperations.Items.Add(this.resetPMADCCCommand);
            this.menuPMADCCOperations.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuPMADCCOperations.Label = "PMADCC operations";
            this.menuPMADCCOperations.Name = "menuPMADCCOperations";
            this.menuPMADCCOperations.ShowImage = true;
            // 
            // importPMADCCCommand
            // 
            this.importPMADCCCommand.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.importPMADCCCommand.Image = ((System.Drawing.Image)(resources.GetObject("importPMADCCCommand.Image")));
            this.importPMADCCCommand.Label = "Import from PMADCC ";
            this.importPMADCCCommand.Name = "importPMADCCCommand";
            this.importPMADCCCommand.ShowImage = true;
            this.importPMADCCCommand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importPMADCCCommand_Click);
            // 
            // exportPMADCCCommand
            // 
            this.exportPMADCCCommand.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.exportPMADCCCommand.Image = ((System.Drawing.Image)(resources.GetObject("exportPMADCCCommand.Image")));
            this.exportPMADCCCommand.Label = "Export to PMADCC";
            this.exportPMADCCCommand.Name = "exportPMADCCCommand";
            this.exportPMADCCCommand.ShowImage = true;
            this.exportPMADCCCommand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportPMADCCCommand_Click);
            // 
            // resetPMADCCCommand
            // 
            this.resetPMADCCCommand.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.resetPMADCCCommand.Image = ((System.Drawing.Image)(resources.GetObject("resetPMADCCCommand.Image")));
            this.resetPMADCCCommand.Label = "Reset PMADCC";
            this.resetPMADCCCommand.Name = "resetPMADCCCommand";
            this.resetPMADCCCommand.ShowImage = true;
            this.resetPMADCCCommand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.resetPMADCCCommand_Click);
            // 
            // menuShowDiagram
            // 
            this.menuShowDiagram.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuShowDiagram.Image = ((System.Drawing.Image)(resources.GetObject("menuShowDiagram.Image")));
            this.menuShowDiagram.Items.Add(this.showOriginalDiagramButton);
            this.menuShowDiagram.Items.Add(this.showWorkDiagramButton);
            this.menuShowDiagram.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuShowDiagram.Label = "Show PMADCC diagram";
            this.menuShowDiagram.Name = "menuShowDiagram";
            this.menuShowDiagram.ShowImage = true;
            // 
            // showOriginalDiagramButton
            // 
            this.showOriginalDiagramButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.showOriginalDiagramButton.Image = ((System.Drawing.Image)(resources.GetObject("showOriginalDiagramButton.Image")));
            this.showOriginalDiagramButton.Label = "Show original graphical diagram";
            this.showOriginalDiagramButton.Name = "showOriginalDiagramButton";
            this.showOriginalDiagramButton.ShowImage = true;
            this.showOriginalDiagramButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showDiagramButton_Click);
            // 
            // showWorkDiagramButton
            // 
            this.showWorkDiagramButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.showWorkDiagramButton.Image = ((System.Drawing.Image)(resources.GetObject("showWorkDiagramButton.Image")));
            this.showWorkDiagramButton.Label = "Show graphical diagram with changes";
            this.showWorkDiagramButton.Name = "showWorkDiagramButton";
            this.showWorkDiagramButton.ShowImage = true;
            this.showWorkDiagramButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showWorkDiagramButton_Click);
            // 
            // savePmadccFileDialog
            // 
            this.savePmadccFileDialog.DefaultExt = "pmadcc";
            this.savePmadccFileDialog.Filter = "Process Modeling and Dashboard Control Console Files (*.pmadcc)|*.pmadcc";
            // 
            // openPmadccFileDialog
            // 
            this.openPmadccFileDialog.DefaultExt = "pmadcc";
            this.openPmadccFileDialog.Filter = "Process Modeling and Dashboard Control Console Files (*.pmadcc)|*.pmadcc";
            // 
            // AddinRibbonComponent
            // 
            this.Name = "AddinRibbonComponent";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.pmadccTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AddinRibbonComponent_Load);
            this.pmadccTab.ResumeLayout(false);
            this.pmadccTab.PerformLayout();
            this.excelPMADCCGroup.ResumeLayout(false);
            this.excelPMADCCGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab pmadccTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup excelPMADCCGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportPMADCCCommand;
        private System.Windows.Forms.SaveFileDialog savePmadccFileDialog;
        private System.Windows.Forms.OpenFileDialog openPmadccFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showOriginalDiagramButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showWorkDiagramButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuShowDiagram;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuPMADCCOperations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton importPMADCCCommand;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resetPMADCCCommand;
    }

    partial class ThisRibbonCollection
    {
        internal AddinRibbonComponent AddinRibbonComponent
        {
            get { return this.GetRibbon<AddinRibbonComponent>(); }
        }
    }
}
