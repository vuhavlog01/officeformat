namespace WordAddIn1
{
    partial class RibbonVFormat : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonVFormat()
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
            this.btnCanLeA4 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnBoldNumbering = this.Factory.CreateRibbonButton();
            this.btnCheckSpacing = this.Factory.CreateRibbonButton();
            this.btnClearHighlights = this.Factory.CreateRibbonButton();
            this.btnCheckSpelling = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnChuanHoaNoiNhan = this.Factory.CreateRibbonButton();
            this.btnConvertFooterToTable = this.Factory.CreateRibbonButton();
            this.btnThemLoaiVanBan = this.Factory.CreateRibbonButton();
            this.btnXuLyNoiDungDeMucTheo = this.Factory.CreateRibbonButton();
            this.btnChuanHoaKinhGui = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "V-FORMAT";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnCanLeA4);
            this.group1.Name = "group1";
            // 
            // btnCanLeA4
            // 
            this.btnCanLeA4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCanLeA4.Image = global::WordAddIn1.Properties.Resources.margin_a41;
            this.btnCanLeA4.Label = "Căn Lề Trang A4";
            this.btnCanLeA4.Name = "btnCanLeA4";
            this.btnCanLeA4.ShowImage = true;
            this.btnCanLeA4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCanLeA4_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnBoldNumbering);
            this.group2.Items.Add(this.btnCheckSpacing);
            this.group2.Items.Add(this.btnClearHighlights);
            this.group2.Items.Add(this.btnCheckSpelling);
            this.group2.Name = "group2";
            // 
            // btnBoldNumbering
            // 
            this.btnBoldNumbering.Image = global::WordAddIn1.Properties.Resources.bold_number;
            this.btnBoldNumbering.Label = "Bold Number";
            this.btnBoldNumbering.Name = "btnBoldNumbering";
            this.btnBoldNumbering.ShowImage = true;
            this.btnBoldNumbering.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBoldNumbering_Click);
            // 
            // btnCheckSpacing
            // 
            this.btnCheckSpacing.Image = global::WordAddIn1.Properties.Resources.check_spacing;
            this.btnCheckSpacing.Label = "Check Spacing";
            this.btnCheckSpacing.Name = "btnCheckSpacing";
            this.btnCheckSpacing.ShowImage = true;
            this.btnCheckSpacing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckSpacing_Click);
            // 
            // btnClearHighlights
            // 
            this.btnClearHighlights.Image = global::WordAddIn1.Properties.Resources.clear_highlights;
            this.btnClearHighlights.Label = "Clear Highlights";
            this.btnClearHighlights.Name = "btnClearHighlights";
            this.btnClearHighlights.ShowImage = true;
            this.btnClearHighlights.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearHighlights_Click);
            // 
            // btnCheckSpelling
            // 
            this.btnCheckSpelling.Image = global::WordAddIn1.Properties.Resources.check_spelling;
            this.btnCheckSpelling.Label = "Check Spelling";
            this.btnCheckSpelling.Name = "btnCheckSpelling";
            this.btnCheckSpelling.ShowImage = true;
            this.btnCheckSpelling.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckSpelling_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnChuanHoaNoiNhan);
            this.group3.Items.Add(this.btnConvertFooterToTable);
            this.group3.Items.Add(this.btnThemLoaiVanBan);
            this.group3.Items.Add(this.btnXuLyNoiDungDeMucTheo);
            this.group3.Items.Add(this.btnChuanHoaKinhGui);
            this.group3.Name = "group3";
            // 
            // btnChuanHoaNoiNhan
            // 
            this.btnChuanHoaNoiNhan.Image = global::WordAddIn1.Properties.Resources.chuan_hoa_noi_nhan;
            this.btnChuanHoaNoiNhan.Label = "Chuan Hoa Noi Nhan";
            this.btnChuanHoaNoiNhan.Name = "btnChuanHoaNoiNhan";
            this.btnChuanHoaNoiNhan.ShowImage = true;
            this.btnChuanHoaNoiNhan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChuanHoaNoiNhan_Click);
            // 
            // btnConvertFooterToTable
            // 
            this.btnConvertFooterToTable.Image = global::WordAddIn1.Properties.Resources.convert_table;
            this.btnConvertFooterToTable.Label = "Convert Footer To Table";
            this.btnConvertFooterToTable.Name = "btnConvertFooterToTable";
            this.btnConvertFooterToTable.ShowImage = true;
            this.btnConvertFooterToTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConvertFooterToTable_Click);
            // 
            // btnThemLoaiVanBan
            // 
            this.btnThemLoaiVanBan.Label = "Thêm Loại VB";
            this.btnThemLoaiVanBan.Name = "btnThemLoaiVanBan";
            this.btnThemLoaiVanBan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnThemLoaiVanBan_Click);
            // 
            // btnXuLyNoiDungDeMucTheo
            // 
            this.btnXuLyNoiDungDeMucTheo.Label = "Chuẩn hóa đề mục";
            this.btnXuLyNoiDungDeMucTheo.Name = "btnXuLyNoiDungDeMucTheo";
            this.btnXuLyNoiDungDeMucTheo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnXuLyNoiDungDeMucTheo_Click);
            // 
            // btnChuanHoaKinhGui
            // 
            this.btnChuanHoaKinhGui.Label = "Chuẩn hóa kính gửi/kính trình";
            this.btnChuanHoaKinhGui.Name = "btnChuanHoaKinhGui";
            this.btnChuanHoaKinhGui.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChuanHoaKinhGui_Click);
            // 
            // RibbonVFormat
            // 
            this.Name = "RibbonVFormat";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonVFormat_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCanLeA4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBoldNumbering;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckSpelling;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckSpacing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearHighlights;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChuanHoaNoiNhan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertFooterToTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnXuLyNoiDungDeMucTheo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThemLoaiVanBan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChuanHoaKinhGui;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonVFormat RibbonVFormat
        {
            get { return this.GetRibbon<RibbonVFormat>(); }
        }
    }
}
