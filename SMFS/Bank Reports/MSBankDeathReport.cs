﻿using DevExpress.Utils;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraPrinting;
using GeneralLib;
using MyXtraGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
/****************************************************************************************/
namespace SMFS
{
    /****************************************************************************************/
    public partial class MSBankDeathReport : DevExpress.XtraEditors.XtraForm
    {
        private DataTable groupContracts = null;
        private DataTable agentsDt = null;
        private bool runAgents = false;
        private DataTable paymentDetail = null;
        private string bankDetails = "";
        private double beginningBalance = 0D;
        private string workReport = "";
        /****************************************************************************************/
        public MSBankDeathReport(string report)
        {
            InitializeComponent();
            SetupTotalsSummary();
            workReport = report;
        }
        /****************************************************************************************/
        private void SetupTotalsSummary()
        {
            AddSummaryColumn("amtActuallyReceived", null);
        }
        /****************************************************************************************/
        private void AddSummaryColumn(string columnName)
        {
            gridMain.Columns[columnName].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            //gridMain.Columns[columnName].SummaryItem.DisplayFormat = "{0:0,0.00}";
            gridMain.Columns[columnName].SummaryItem.DisplayFormat = "{0:N2}";
        }
        /****************************************************************************************/
        private void AddSummaryColumn(string columnName, DevExpress.XtraGrid.Views.BandedGrid.AdvBandedGridView gMain = null)
        {
            if (gMain == null)
                gMain = gridMain;
            //            gMain.Columns[columnName].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            gMain.Columns[columnName].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            //            gMain.Columns[columnName].SummaryItem.DisplayFormat = "${0:0,0.00}";
            gMain.Columns[columnName].SummaryItem.DisplayFormat = "{0:N2}";
        }
        /****************************************************************************************/
        private void MSBankDeathReport_Load(object sender, EventArgs e)
        {
            this.Text = workReport;

            //PleaseWait pleaseForm = G1.StartWait("Please Wait!\nLoading Information . . .");

            pictureAdd.Hide();
            pictureDelete.Hide();
            chkComboLocation.Hide();
            label2.Hide();
            txtBeginningBalance.Hide();
            label1.Hide();

            btnGetDeposits.Hide();

            btnSave.Hide();
            txtBeginningBalance.Hide();
            label1.Hide();

            chkComboLocation.Hide();
            label2.Hide();

            DateTime now = DateTime.Now;
            now = now.AddMonths(-1);
            now = new DateTime(now.Year, now.Month, 1);

            int days = DateTime.DaysInMonth(now.Year, now.Month);
            DateTime stop = new DateTime(now.Year, now.Month, days);
            this.dateTimePicker1.Value = stop;

            this.Cursor = Cursors.WaitCursor;

            DataTable dt = GetPayments();

            G1.NumberDataTable(dt);
            dgv.DataSource = dt;

            this.Cursor = Cursors.Default;

            //G1.StopWait(ref pleaseForm);
        }
        /***************************************************************************************/
        private DataTable bankDt = null;
        private DataTable GetPayments()
        {
            string cmd = "";
            DataTable dt = null;

            DateTime stopDate = this.dateTimePicker1.Value;
            DateTime startDate = new DateTime(stopDate.Year, stopDate.Month, 1);
            int days = DateTime.DaysInMonth(stopDate.Year, stopDate.Month);
            stopDate = new DateTime(stopDate.Year, stopDate.Month, days);

            string beginDate = startDate.ToString("yyyy-MM-dd 00:00:00");
            string endDate = stopDate.ToString("yyyy-MM-dd 23:59:59");

            //SELECT* FROM cust_payment_details WHERE `type` = 'INSURANCE UNITY'
            //AND (( `dateFiled` = '00010101' AND (`dateReceived` >= '2022-03-01' AND `dateReceived` <= '2022-03-31 23:59:59'))
            //OR(`dateFiled` >= '2022-03-01' AND `dateFiled` <= '2022-03-31 23:59:59')
            //OR(`dateReceived` >= '2022-03-01 00:00:00' AND `dateReceived` <= '2022-03-31 23:59:59'))
            //ORDER BY `dateFiled`;



            cmd = "SELECT * FROM cust_payment_details WHERE `type` = 'TRUST' and `paidFrom` = 'UNITY' ";
            cmd += " AND `localDescription` = 'Mag. St. Bank - Bay Springs' ";
            cmd += " AND (( `dateFiled` = '00010101' AND (`dateReceived` >= '" + beginDate + "' AND `dateReceived` <= '" + endDate + "' ))";
            cmd += " OR (`dateFiled` >= '" + beginDate + "' AND `dateFiled` <= '" + endDate + "') ";
            cmd += " OR (`dateReceived` >= '" + beginDate + "' AND `dateReceived` <= '" + endDate + "'))";
            cmd += " ORDER BY `dateFiled`;"; // Ramma Zamma

            DataTable dx = G1.get_db_data(cmd);
            dx.Columns.Add("name");
            dx.Columns.Add("deceasedDate");
            dx.Columns.Add("serviceId");
            dx.Columns.Add("lastName");
            dx.Columns.Add("firstName");
            dx.Columns.Add("trustNumber");
            dx.Columns.Add("adate");

            string paymentRecord = "";

            string contractNumber = "";
            string prefix = "";
            string fName = "";
            string mName = "";
            string lName = "";
            string suffix = "";
            string name = "";

            for (int i = 0; i < dx.Rows.Count; i++)
            {
                paymentRecord = dx.Rows[i]["paymentRecord"].ObjToString();
                contractNumber = dx.Rows[i]["contractNumber"].ObjToString();
                cmd = "Select * from `fcustomers` WHERE `contractNumber` = '" + contractNumber + "';";
                dt = G1.get_db_data(cmd);
                if (dt.Rows.Count > 0)
                {
                    dx.Rows[i]["serviceId"] = dt.Rows[0]["serviceId"].ObjToString();
                    dx.Rows[i]["deceasedDate"] = dt.Rows[0]["deceasedDate"].ObjToDateTime().ToString("MM/dd/yyyy");

                    prefix = dt.Rows[0]["prefix"].ObjToString();
                    fName = dt.Rows[0]["firstName"].ObjToString();
                    mName = dt.Rows[0]["middleName"].ObjToString();
                    lName = dt.Rows[0]["lastName"].ObjToString();
                    suffix = dt.Rows[0]["suffix"].ObjToString();

                    name = G1.BuildFullName(prefix, fName, mName, lName, suffix);

                    dx.Rows[i]["name"] = name;
                    dx.Rows[i]["lastName"] = lName;
                    dx.Rows[i]["firstName"] = fName;
                }
                if ( !String.IsNullOrEmpty ( paymentRecord ))
                {
                    cmd = "Select * from `cust_payments` where `record` = '" + paymentRecord + "';";
                    dt = G1.get_db_data(cmd);
                    if (dt.Rows.Count > 0)
                        dx.Rows[i]["trustNumber"] = dt.Rows[0]["trust_policy"].ObjToString();
                }
            }
            DataView tempview = dx.DefaultView;
            tempview.Sort = "dateReceived ASC";
            dx = tempview.ToTable();

            return dx;
        }
        /****************************************************************************************/
        private void btnGetDeposits_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            DataTable dt = GetPayments();

            G1.NumberDataTable(dt);
            dgv.DataSource = dt;

            this.Cursor = Cursors.Default;
        }
        /****************************************************************************************/
        private void btnRight_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DateTime now = this.dateTimePicker1.Value;
            now = now.AddMonths(1);
            now = new DateTime(now.Year, now.Month, 1);
            //this.dateTimePicker1.Value = now;
            int days = DateTime.DaysInMonth(now.Year, now.Month);
            this.dateTimePicker1.Value = new DateTime(now.Year, now.Month, days);

            DateTime startDate = now;
            DateTime stopDate = this.dateTimePicker1.Value;

            btnGetDeposits_Click(null, null);

            this.Cursor = Cursors.Default;
        }
        /****************************************************************************************/
        private void btnLeft_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DateTime now = this.dateTimePicker1.Value;
            now = now.AddMonths(-1);
            now = new DateTime(now.Year, now.Month, 1);
            //this.dateTimePicker1.Value = now;
            int days = DateTime.DaysInMonth(now.Year, now.Month);
            this.dateTimePicker1.Value = new DateTime(now.Year, now.Month, days);

            DateTime startDate = now;
            DateTime stopDate = this.dateTimePicker1.Value;

            btnGetDeposits_Click(null, null);

            this.Cursor = Cursors.Default;
        }
        /****************************************************************************************/
        private void gridMain_CustomDrawGroupRow(object sender, DevExpress.XtraGrid.Views.Base.RowObjectCustomDrawEventArgs e)
        {
            //GridGroupRowInfo info = e.Info as GridGroupRowInfo;
            //string location = info.GroupText;
            //int idx = location.LastIndexOf(']');
            //if (idx > 0)
            //{
            //    location = location.Substring(idx+1);
            //    DataRow[] dRows = groupContracts.Select("loc='" + location.Trim() + "'");
            //    if (dRows.Length > 0)
            //        info.GroupText += " " + dRows[0]["contracts"].ObjToString();
            //}
        }
        /***********************************************************************************************/
        private int pageMarginLeft = 0;
        private int pageMarginRight = 0;
        private int pageMarginTop = 0;
        private int pageMarginBottom = 0;
        /***********************************************************************************************/
        private void printPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.components == null)
                this.components = new System.ComponentModel.Container();

            DevExpress.XtraPrinting.PrintingSystem printingSystem1 = new DevExpress.XtraPrinting.PrintingSystem(this.components);
            DevExpress.XtraPrinting.PrintableComponentLink printableComponentLink1 = new DevExpress.XtraPrinting.PrintableComponentLink(this.components);

            printingSystem1.Links.AddRange(new object[] {
            printableComponentLink1});


            printableComponentLink1.Component = dgv;
            printableComponentLink1.PrintingSystemBase = printingSystem1;

            printableComponentLink1.EnablePageDialog = true;

            printableComponentLink1.CreateDetailHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateDetailHeaderArea);
            printableComponentLink1.CreateMarginalHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateMarginalHeaderArea);
            printableComponentLink1.BeforeCreateAreas += new System.EventHandler(this.printableComponentLink1_BeforeCreateAreas);
            printableComponentLink1.AfterCreateAreas += new System.EventHandler(this.printableComponentLink1_AfterCreateAreas);
            printableComponentLink1.Landscape = false;

            Printer.setupPrinterMargins(10, 5, 80, 50);

            pageMarginLeft = Printer.pageMarginLeft;
            pageMarginRight = Printer.pageMarginRight;
            pageMarginTop = Printer.pageMarginTop;
            pageMarginBottom = Printer.pageMarginBottom;

            printableComponentLink1.Margins.Left = pageMarginLeft;
            printableComponentLink1.Margins.Right = pageMarginRight;
            printableComponentLink1.Margins.Top = pageMarginTop;
            printableComponentLink1.Margins.Bottom = pageMarginBottom;

            gridMain.Columns["num"].Visible = false;

            //printingSystem1.Document.AutoFitToPagesWidth = 1;

            printableComponentLink1.CreateDocument();
            printableComponentLink1.ShowPreview();

            gridMain.Columns["num"].Visible = true;
        }
        /***********************************************************************************************/
        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.components == null)
                this.components = new System.ComponentModel.Container();
            DevExpress.XtraPrinting.PrintingSystem printingSystem1 = new DevExpress.XtraPrinting.PrintingSystem(this.components);
            DevExpress.XtraPrinting.PrintableComponentLink printableComponentLink1 = new DevExpress.XtraPrinting.PrintableComponentLink(this.components);

            printingSystem1.Links.AddRange(new object[] {
            printableComponentLink1});

            printableComponentLink1.Component = dgv;
            printableComponentLink1.PrintingSystemBase = printingSystem1;
            printableComponentLink1.CreateDetailHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateDetailHeaderArea);
            printableComponentLink1.CreateMarginalHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateMarginalHeaderArea);
            printableComponentLink1.BeforeCreateAreas += new System.EventHandler(this.printableComponentLink1_BeforeCreateAreas);
            printableComponentLink1.AfterCreateAreas += new System.EventHandler(this.printableComponentLink1_AfterCreateAreas);
            printableComponentLink1.Landscape = false;

            Printer.setupPrinterMargins(10, 5, 80, 50);

            pageMarginLeft = Printer.pageMarginLeft;
            pageMarginRight = Printer.pageMarginRight;
            pageMarginTop = Printer.pageMarginTop;
            pageMarginBottom = Printer.pageMarginBottom;

            printableComponentLink1.Margins.Left = pageMarginLeft;
            printableComponentLink1.Margins.Right = pageMarginRight;
            printableComponentLink1.Margins.Top = pageMarginTop;
            printableComponentLink1.Margins.Bottom = pageMarginBottom;

            gridMain.Columns["num"].Visible = false;

            //printingSystem1.Document.AutoFitToPagesWidth = 1;

            printableComponentLink1.CreateDocument();
            printableComponentLink1.PrintDlg();

            gridMain.Columns["num"].Visible = true;
        }
        /***********************************************************************************************/
        private void printableComponentLink1_BeforeCreateAreas(object sender, EventArgs e)
        {
        }
        /***********************************************************************************************/
        private void printableComponentLink1_AfterCreateAreas(object sender, EventArgs e)
        {
        }
        /***********************************************************************************************/
        private void printableComponentLink1_CreateDetailHeaderArea(object sender, CreateAreaEventArgs e)
        {
        }
        /***********************************************************************************************/
        private void printableComponentLink1_CreateMarginalHeaderArea(object sender, CreateAreaEventArgs e)
        {
            Printer.setupPrinterQuads(e, 2, 3);
            Font font = new Font("Ariel", 16);
            Printer.DrawQuad(1, 1, Printer.xQuads, 2, "South Mississippi Funeral Services, LLC", Color.Black, BorderSide.Top, font, HorizontalAlignment.Center);

            Printer.SetQuadSize(12, 12);

            font = new Font("Ariel", 8);
            Printer.DrawGridDate(2, 3, 2, 3, Color.Black, BorderSide.None, font);
            Printer.DrawGridPage(11, 3, 2, 3, Color.Black, BorderSide.None, font);

            //            Printer.DrawQuad(1, 9, 2, 3, "User : " + LoginForm.username, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Center);

            font = new Font("Ariel", 10, FontStyle.Regular);
            string title = this.Text;
            Printer.DrawQuad(5, 7, 5, 3, title, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            Printer.DrawQuad(1, 9, 2, 3, "User : " + LoginForm.username, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);


            DateTime date = this.dateTimePicker1.Value;
            string workDate = date.Month.ToString("D2") + "/" + date.Year.ToString("D4");
            Printer.SetQuadSize(24, 12);
            font = new Font("Ariel", 9, FontStyle.Regular);
            Printer.DrawQuad(20, 8, 5, 4, "Month Closing - " + workDate, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);

            //Printer.DrawQuad(16, 8, 3, 4, lblPayment.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(19, 8, 3, 4, lblTrust85.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(22, 8, 3, 4, lblTrust100.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);

            Printer.SetQuadSize(12, 12);
            Printer.DrawQuadBorder(1, 1, 12, 12, BorderSide.All, 1, Color.Black);
            Printer.DrawQuadBorder(12, 1, 1, 12, BorderSide.Right, 1, Color.Black);
        }
        /****************************************************************************************/
        private int footerCount = 0;
        private void gridMain_BeforePrintRow(object sender, DevExpress.XtraGrid.Views.Printing.CancelPrintRowEventArgs e)
        {
            if (e.HasFooter)
            {
                footerCount++;
            }
        }
        /****************************************************************************************/
        private void gridMain_AfterPrintRow(object sender, DevExpress.XtraGrid.Views.Printing.PrintRowEventArgs e)
        {
            if (e.HasFooter)
            {
                footerCount++;
                if (footerCount >= 1)
                {
                    footerCount = 0;
                    e.PS.InsertPageBreak(e.Y);
                }
            }
        }
        /****************************************************************************************/
        private void gridMain_CustomRowFilter(object sender, DevExpress.XtraGrid.Views.Base.RowFilterEventArgs e)
        {
        }
        /****************************************************************************************/
        private void pictureBox1_Click(object sender, EventArgs e)
        { // Spy Glass
            G1.ShowHideFindPanel(gridMain);
            //if (this.gridMain.OptionsFind.AlwaysVisible == true)
            //    gridMain.OptionsFind.AlwaysVisible = false;
            //else
            //    gridMain.OptionsFind.AlwaysVisible = true;
        }
        /****************************************************************************************/
        private bool CheckForSave()
        {
            if (!btnSave.Visible)
                return true;
            DialogResult result = MessageBox.Show("***Question***\nInformation has been modified!\nWould you like to save your changes?", "Modified Data Dialog", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Cancel)
                return false;
            if (result == DialogResult.No)
                return true;
            return true;
        }
        /****************************************************************************************/
        private void gridMain_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
        }
        /****************************************************************************************/
        private void gridMain_DoubleClick(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);

            string contractNumber = dr["contractNumber"].ObjToString();
            FunPayments editFunPayments = new FunPayments(this, contractNumber, "", false, true );
            editFunPayments.Show();
        }
        /****************************************************************************************/
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime date = this.dateTimePicker1.Value;
            int days = DateTime.DaysInMonth(date.Year, date.Month);
            date = new DateTime(date.Year, date.Month, days);
            this.dateTimePicker1.Value = date;
        }
        /****************************************************************************************/
        private void pictureAdd_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
                return;
            DataRow dr = gridMain.GetFocusedDataRow();
            DateTime date = dr["dateReceived"].ObjToDateTime();
            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);
            if (row < 0 || row > (dt.Rows.Count - 1))
                return;
            if (rowHandle == (dt.Rows.Count - 1))
                return; // Already at the last row
            DataRow dRow = dt.NewRow();
            dRow["dateReceived"] = G1.DTtoMySQLDT(date);
            //dRow["TDA"] = 0.0D;
            //dRow["IDA"] = 0.0D;
            //dRow["NDA"] = 0.0D;
            //dRow["returns"] = 0.0D;
            //dRow["transfers"] = 0.0D;
            //dRow["dailyTotals"] = 0.00D;
            //dRow["comment"] = "Enter Comment Here";
            //dRow["dow"] = G1.DayOfWeekText(date);
            dt.Rows.InsertAt(dRow, row);
            G1.NumberDataTable(dt);
            dt.AcceptChanges();
            dgv.DataSource = dt;
            gridMain.ClearSelection();
            gridMain.RefreshData();
            gridMain.FocusedRowHandle = rowHandle + 1;
            gridMain.SelectRow(rowHandle + 1);
            dgv.Refresh();
        }
        /****************************************************************************************/
        private void gridMain_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.ToUpper().IndexOf("DATE") >= 0 && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                if (e.DisplayText.IndexOf("0000") >= 0 || e.DisplayText.IndexOf("0001") >= 0)
                    e.DisplayText = "";
                else
                {
                    DateTime date = e.DisplayText.ObjToString().ObjToDateTime();
                    e.DisplayText = date.ToString("MM/dd/yyyy");
                    if (date.Year < 30)
                        e.DisplayText = "";
                }
            }
        }
        /****************************************************************************************/
        private void gridMain_CustomSummaryCalculate(object sender, DevExpress.Data.CustomSummaryEventArgs e)
        {
            //double balance = 0D;
            //DataTable dt = (DataTable)dgv.DataSource;
            //if (dt.Rows.Count > 0)
            //{
            //    int lastRow = dt.Rows.Count - 1;
            //    balance = dt.Rows[lastRow]["balance"].ObjToDouble();
            //}
            //string str = G1.ReformatMoney(balance);
            //e.TotalValue = str;
        }
        /****************************************************************************************/
        private void dateTimePicker1_CloseUp(object sender, EventArgs e)
        {
            DateTime date = this.dateTimePicker1.Value;
            int days = DateTime.DaysInMonth(date.Year, date.Month);
            date = new DateTime(date.Year, date.Month, days);
            this.dateTimePicker1.Value = date;

            if (!CheckForSave())
                return;

            btnGetDeposits_Click(null, null);
        }
        /****************************************************************************************/
        private void gridMain_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {

            //GridView View = sender as GridView;
            //if (e.RowHandle >= 0)
            //{
            //    string column = e.Column.FieldName.ToUpper();
            //    if (column == "NUM")
            //    {
            //        DataTable dt = (DataTable)dgv.DataSource;
            //        int row = gridMain.GetDataSourceRowIndex(e.RowHandle);

            //        string comment = dt.Rows[row]["comment"].ObjToString();
            //        if (comment.Trim().ToUpper() != "BALANCE FORWARD")
            //        {
            //            string adate = dt.Rows[row]["adate"].ObjToString();

            //            if (String.IsNullOrWhiteSpace(adate) && workReport != "Funeral Detail Report")
            //            {
            //                e.Appearance.BackColor = Color.Red;
            //            }
            //        }
            //    }
            //}
        }
        /****************************************************************************************/
        private void pictureDelete_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
                return;
            DataRow dr = gridMain.GetFocusedDataRow();
            DateTime date = dr["dateReceived"].ObjToDateTime();
            string aDate = dr["aDate"].ObjToString();
            if (!String.IsNullOrWhiteSpace(aDate))
                return;
            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);
            if (row < 0 || row > (dt.Rows.Count - 1))
                return;
            dt.Rows.RemoveAt(row);

//            double balance = RecalcTotals(dt);

            G1.NumberDataTable(dt);
            dt.AcceptChanges();
            dgv.DataSource = dt;
            dgv.Refresh();

            btnSave.Show();
        }
        /****************************************************************************************/
        private double originalSize = 0D;
        private Font mainFont = null;
        private Font newFont = null;
        private void ScaleCells()
        {
            if (originalSize == 0D)
            {
                //                originalSize = gridMain.Columns["address1"].AppearanceCell.FontSizeDelta.ObjToDouble();
                originalSize = gridMain.Columns["comment"].AppearanceCell.Font.Size;
                mainFont = gridMain.Columns["comment"].AppearanceCell.Font;
            }
            double scale = txtScale.Text.ObjToDouble();
            double size = scale / 100D * originalSize;
            Font font = new Font(mainFont.Name, (float)size);
            for (int i = 0; i < gridMain.Columns.Count; i++)
            {
                gridMain.Columns[i].AppearanceCell.Font = font;
            }
            gridMain.Appearance.GroupFooter.Font = font;
            gridMain.AppearancePrint.FooterPanel.Font = font;
            newFont = font;
            gridMain.RefreshData();
            gridMain.RefreshEditor(true);
            dgv.Refresh();
            this.Refresh();
        }
        /****************************************************************************************/
        private void txtScale_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string balance = txtScale.Text.Trim();
                if (!G1.validate_numeric(balance))
                {
                    MessageBox.Show("***ERROR*** Scale must be numeric!");
                    return;
                }
                double money = balance.ObjToDouble();
                balance = G1.ReformatMoney(money);
                txtScale.Text = balance;
                ScaleCells();
                return;
            }
            // Initialize the flag to false.
            bool nonNumberEntered = false;

            // Determine whether the keystroke is a number from the top of the keyboard.
            if (e.KeyCode < Keys.D0 || e.KeyCode > Keys.D9)
            {
                // Determine whether the keystroke is a number from the keypad.
                if (e.KeyCode < Keys.NumPad0 || e.KeyCode > Keys.NumPad9)
                {
                    // Determine whether the keystroke is a backspace.
                    if (e.KeyCode != Keys.Back)
                    {
                        // A non-numerical keystroke was pressed.
                        // Set the flag to true and evaluate in KeyPress event.
                        if (e.KeyCode != Keys.OemPeriod)
                            nonNumberEntered = true;
                    }
                }
            }
            //If shift key was pressed, it's not a number.
            if (Control.ModifierKeys == Keys.Shift)
            {
                nonNumberEntered = true;
            }
            if (nonNumberEntered)
            {
                MessageBox.Show("***ERROR*** Key entered must be a number!");
                e.Handled = true;
            }
        }
        /****************************************************************************************/
        private void SaveData()
        {
            this.Cursor = Cursors.WaitCursor;
            DateTime stopDate = this.dateTimePicker1.Value;
            DateTime startDate = new DateTime(stopDate.Year, stopDate.Month, 1);
            string date1 = G1.DateTimeToSQLDateTime(startDate);
            string date2 = G1.DateTimeToSQLDateTime(stopDate);

            DateTime saveDate = DateTime.Now;
            string date = "";

            DataTable dt = (DataTable)dgv.DataSource;

            double tda = 0D;
            double ida = 0D;
            double nda = 0D;
            double dda = 0D;
            double returns = 0D;
            double transfer = 0D;
            double balance = 0D;
            string comment = "";
            string aDate = "";
            string record = "";
            string manual = "";
            string bankAccount = "";
            string oldBankAccount = "";
            double endingBalance = 0D;

            double credits = 0D;
            double debits = 0D;
            string accountTitle = "";

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    bankAccount = dt.Rows[i]["bank_account"].ObjToString();
            //    if (String.IsNullOrWhiteSpace(oldBankAccount))
            //    {
            //        oldBankAccount = bankAccount;
            //        string cmd = "DELETE FROM `lockboxdeposits` where `date` >= '" + date1 + "' AND `date` <= '" + date2 + "' AND `bank_account` = '" + bankAccount + "';";
            //        G1.get_db_data(cmd);
            //    }
            //    if (oldBankAccount != bankAccount)
            //    {
            //        beginningBalance = getBeginningBalance(oldBankAccount, startDate, ref endingBalance, ref accountTitle);
            //        UpdateBankTotals(oldBankAccount, this.dateTimePicker1.Value, credits, debits, beginningBalance, balance);
            //        oldBankAccount = bankAccount;

            //        string cmd = "DELETE FROM `lockboxdeposits` where `date` >= '" + date1 + "' AND `date` <= '" + date2 + "' AND `bank_account` = '" + bankAccount + "';";
            //        G1.get_db_data(cmd);
            //    }

            //    saveDate = dt.Rows[i]["date"].ObjToDateTime();
            //    date = saveDate.ToString("MM/dd/yyyy");
            //    tda = dt.Rows[i]["TDA"].ObjToDouble();
            //    ida = dt.Rows[i]["IDA"].ObjToDouble();
            //    nda = dt.Rows[i]["NDA"].ObjToDouble();
            //    dda = dt.Rows[i]["NDA"].ObjToDouble();
            //    returns = dt.Rows[i]["returns"].ObjToDouble();
            //    transfer = dt.Rows[i]["transfers"].ObjToDouble();
            //    balance = dt.Rows[i]["balance"].ObjToDouble();
            //    comment = dt.Rows[i]["comment"].ObjToString();
            //    if (comment.Trim().ToUpper() == "BALANCE FORWARD")
            //        continue;
            //    aDate = dt.Rows[i]["adate"].ObjToString();
            //    accountTitle = dt.Rows[i]["accountTitle"].ObjToString();
            //    manual = dt.Rows[i]["manual"].ObjToString();

            //    tda = G1.RoundValue(tda);
            //    ida = G1.RoundValue(ida);
            //    nda = G1.RoundValue(nda);
            //    dda = G1.RoundValue(nda);
            //    returns = G1.RoundValue(returns);
            //    transfer = G1.RoundValue(transfer);
            //    balance = G1.RoundValue(balance);

            //    credits += tda + ida + nda;
            //    debits += returns + transfer;

            //    record = G1.create_record("lockboxdeposits", "comment", "-1");
            //    if (G1.BadRecord("lockboxdeposits", record))
            //        break;
            //    G1.update_db_table("lockboxdeposits", "record", record, new string[] { "date", date, "adate", aDate, "comment", comment, "TDA", tda.ToString(), "IDA", ida.ToString(), "NDA", nda.ToString(), "dda", dda.ToString(), "returns", returns.ToString(), "transfers", transfer.ToString(), "balance", balance.ToString(), "manual", manual, "bank_account", bankAccount, "order", i.ToString() });

            //    if (i == (dt.Rows.Count - 1))
            //    {
            //        beginningBalance = getBeginningBalance(bankAccount, startDate, ref endingBalance, ref accountTitle);
            //        UpdateBankTotals(bankAccount, this.dateTimePicker1.Value, credits, debits, beginningBalance, balance);
            //        oldBankAccount = bankAccount;
            //    }
            //}

            this.Cursor = Cursors.Default;
            btnSave.Hide();
        }
        /****************************************************************************************/
        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveData();
        }
        /****************************************************************************************/
        private void btnGetDeposits_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            DataTable dt = GetPayments();

            G1.NumberDataTable(dt);
            dgv.DataSource = dt;
        }
        /****************************************************************************************/
    }
}