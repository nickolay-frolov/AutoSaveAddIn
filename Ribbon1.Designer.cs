namespace AutoSaveAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.onOffCB = this.Factory.CreateRibbonComboBox();
            this.saveIntervalEB = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.saveCurrentCB = this.Factory.CreateRibbonCheckBox();
            this.saveAllCB = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.onOffCB);
            this.group1.Items.Add(this.saveIntervalEB);
            this.group1.Label = "Автосохранение";
            this.group1.Name = "group1";
            // 
            // onOffCB
            // 
            ribbonDropDownItemImpl1.Label = "Выкл.";
            ribbonDropDownItemImpl2.Label = "Вкл.";
            this.onOffCB.Items.Add(ribbonDropDownItemImpl1);
            this.onOffCB.Items.Add(ribbonDropDownItemImpl2);
            this.onOffCB.Label = "Вкл./Выкл.";
            this.onOffCB.Name = "onOffCB";
            this.onOffCB.Text = null;
            this.onOffCB.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.onOffCB_TextChanged);
            // 
            // saveIntervalEB
            // 
            this.saveIntervalEB.Label = "Время, сек";
            this.saveIntervalEB.Name = "saveIntervalEB";
            this.saveIntervalEB.Text = null;
            this.saveIntervalEB.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveIntervalEB_TextChanged);
            // 
            // group2
            // 
            this.group2.Items.Add(this.saveCurrentCB);
            this.group2.Items.Add(this.saveAllCB);
            this.group2.Label = "При потере фокуса";
            this.group2.Name = "group2";
            // 
            // saveCurrentCB
            // 
            this.saveCurrentCB.Label = "Сохранять текущий";
            this.saveCurrentCB.Name = "saveCurrentCB";
            this.saveCurrentCB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveCurrentCB_Click);
            // 
            // saveAllCB
            // 
            this.saveAllCB.Label = "Сохранять все";
            this.saveAllCB.Name = "saveAllCB";
            this.saveAllCB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveAllCB_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox onOffCB;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox saveIntervalEB;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox saveCurrentCB;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox saveAllCB;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
