﻿using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using GeneralLib;
using DevExpress.XtraPrinting;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;


using System.Collections.Generic;
using System.ComponentModel;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraPrinting;
/****************************************************************************************/
namespace SMFS
{
    /****************************************************************************************/
    public partial class NewBusiness : DevExpress.XtraEditors.XtraForm
    {
        private DataTable groupContracts = null;
        private DataTable agentsDt = null;
        private bool runAgents = false;
        private DataTable originalDt = null;
        /****************************************************************************************/
        public NewBusiness()
        {
            InitializeComponent();
            SetupTotalsSummary();
        }
        /****************************************************************************************/
        private void SetupTotalsSummary()
        {
            AddSummaryColumn("contractValue", null);
            AddSummaryColumn("ibtrust", null);
            AddSummaryColumn("sptrust", null);
            AddSummaryColumn("total", null);
            AddSummaryColumn("lapses", null);
            AddSummaryColumn("reinstates", null);
            AddSummaryColumn("ibtrustytd", null);
            AddSummaryColumn("sptrustytd", null);
            AddSummaryColumn("totalytd", null);
            AddSummaryColumn("lapsesytd", null);
            AddSummaryColumn("reinstatesytd", null);
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
        private void NewBusiness_Load(object sender, EventArgs e)
        {
            gridBand1.Visible = false;
            DateTime now = DateTime.Now;
            now = now.AddMonths(-1);
            now = new DateTime(now.Year, now.Month, 1);
            this.dateTimePicker1.Value = now;
            int days = DateTime.DaysInMonth(now.Year, now.Month);
            DateTime stop = new DateTime(now.Year, now.Month, days);
            this.dateTimePicker2.Value = stop;

            LoadTabAllLocations();
        }
        /****************************************************************************************/
        private void LoadTabAllLocations()
        {
            DataTable locationDt = new DataTable();
            locationDt.Columns.Add("options");
            AddToLocationCombo(locationDt, "Normal");
            AddToLocationCombo(locationDt, "Trusts");
            AddToLocationCombo(locationDt, "Lapses");
            AddToLocationCombo(locationDt, "Reinstates");
            AddToLocationCombo(locationDt, "Simple by Agent");
            AddToLocationCombo(locationDt, "Simple by Location");
            cmbLocationTotals.Properties.DataSource = locationDt;
        }
        /****************************************************************************************/
        private void AddToLocationCombo(DataTable locationDt, string text)
        {
            DataRow ddrx = locationDt.NewRow();
            ddrx["options"] = text;
            locationDt.Rows.Add(ddrx);
        }
        /****************************************************************************************/
        private void checkedComboBoxEdit1_Properties_EditValueChanged(object sender, EventArgs e)
        {
        }
        /****************************************************************************************/
        private void btnRight_Click(object sender, EventArgs e)
        {
            DateTime now = this.dateTimePicker2.Value;
            now = now.AddMonths(1);
            now = new DateTime(now.Year, now.Month, 1);
            this.dateTimePicker1.Value = now;
            int days = DateTime.DaysInMonth(now.Year, now.Month);
            this.dateTimePicker2.Value = new DateTime(now.Year, now.Month, days);
        }
        /****************************************************************************************/
        private void btnLeft_Click(object sender, EventArgs e)
        {
            DateTime now = this.dateTimePicker1.Value;
            now = now.AddMonths(-1);
            now = new DateTime(now.Year, now.Month, 1);
            this.dateTimePicker1.Value = now;
            int days = DateTime.DaysInMonth(now.Year, now.Month);
            this.dateTimePicker2.Value = new DateTime(now.Year, now.Month, days);
        }
        /****************************************************************************************/
        private void btnRunAgents_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            runAgents = true;
            doSimple = false;
            gridMain.OptionsView.ShowFooter = true;
            gridMain.OptionsView.ShowBands = true;

            gridMain.OptionsPrint.PrintBandHeader = true;
            gridMain.OptionsPrint.PrintFooter = true;

            agentsDt = G1.get_db_data("Select * from `agents`");

            DateTime start = this.dateTimePicker1.Value;
            DateTime stop = this.dateTimePicker2.Value;

            DataTable dt = LoadData(start, stop, null, false);

            if (chkYTD.Checked)
            {
//                stop = start.AddDays(-1);
                start = new DateTime(start.Year, 1, 1);
                if (stop > start)
                    dt = LoadData(start, stop, dt, true);
            }

            G1.NumberDataTable(dt);
            int lastRow = dt.Rows.Count;
            dgv.DataSource = dt;

            originalDt = dt;
            chkExpand_CheckedChanged(null, null);
            ScaleCells();
            this.Cursor = Cursors.Default;
        }
        /****************************************************************************************/
        private void btnRun_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            runAgents = false;
            doSimple = false;

            gridMain.OptionsView.ShowFooter = true;
            gridMain.OptionsView.ShowBands = true;

            gridMain.OptionsPrint.PrintBandHeader = true;
            gridMain.OptionsPrint.PrintFooter = true;


            agentsDt = G1.get_db_data("Select * from `agents`");

            DateTime start = this.dateTimePicker1.Value;
            DateTime stop = this.dateTimePicker2.Value;

            DataTable dt = LoadData(start, stop, null, false);

            if (chkYTD.Checked)
            {
//                stop = start.AddDays(-1);
                start = new DateTime(start.Year, 1, 1);
                if (stop > start)
                    dt = LoadData(start, stop, dt, true);
            }

            LoadUpGroupRows(dt);

            G1.NumberDataTable(dt);
            int lastRow = dt.Rows.Count;
            dgv.DataSource = dt;
            originalDt = dt;
            chkExpand_CheckedChanged(null, null);
            ScaleCells();
            this.Cursor = Cursors.Default;
        }
        /****************************************************************************************/
        private DataTable LoadData(DateTime startDate, DateTime stopDate, DataTable mainDt, bool ytd)
        {
            DateTime start = startDate;
            string date1 = G1.DateTimeToSQLDateTime(start);
            DateTime stop = stopDate;
            string date2 = G1.DateTimeToSQLDateTime(stop);
            string contractNumber = "";
            string loc = "";
            string contract = "";
            string trust = "";
            double contractValue = 0D;
            double downPayment = 0D;
            int idx = 0;
            string ch = "";
            string agentCode = "";
            string cmd = "Select * from `contracts` c JOIN `customers` a on c.`contractNumber` = a.`contractNumber` where `issueDate8` >= '" + date1 + "' AND `issueDate8` <='" + date2 + "';";
            DataTable dt = G1.get_db_data(cmd);

            dt.Columns.Add("ibtrust", Type.GetType("System.Double"));
            dt.Columns.Add("sptrust", Type.GetType("System.Double"));
            dt.Columns.Add("total", Type.GetType("System.Double"));
            dt.Columns.Add("contractValue", Type.GetType("System.Double"));
            dt.Columns.Add("lapses", Type.GetType("System.Double"));
            dt.Columns.Add("reinstates", Type.GetType("System.Double"));
            dt.Columns.Add("loc");
            dt.Columns.Add("Location Name");
            dt.Columns.Add("contracts");
            dt.Columns.Add("agentName");

            dt.Columns.Add("ibtrustytd", Type.GetType("System.Double"));
            dt.Columns.Add("sptrustytd", Type.GetType("System.Double"));
            dt.Columns.Add("totalytd", Type.GetType("System.Double"));
            dt.Columns.Add("lapsesytd", Type.GetType("System.Double"));
            dt.Columns.Add("reinstatesytd", Type.GetType("System.Double"));

            cmd = "Select * from `funeralhomes` order by `keycode`;";
            DataTable dd = G1.get_db_data(cmd);

            string ibtrustCol = "ibtrust";
            string sptrustCol = "sptrust";
            string totalCol = "total";
            if (ytd)
            {
                ibtrustCol += "ytd";
                sptrustCol += "ytd";
                totalCol += "ytd";
            }

            DataRow[] dr = null;
            string deceasedDate = "";
            DateTime ddate = DateTime.Now;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                if ( contractNumber == "L21129L")
                {
                }
                contract = Trust85.decodeContractNumber(contractNumber, ref trust, ref loc);
                dr = dd.Select("keycode='" + loc + "'");
                if (dr.Length > 0)
                {
                    //dt.Rows[i]["Location Name"] = dr[0]["name"].ObjToString();
                    dt.Rows[i]["Location Name"] = dr[0]["LocationCode"].ObjToString();
                }
                else
                    dt.Rows[i]["Location Name"] = loc;
                dt.Rows[i]["loc"] = loc;

                agentCode = dt.Rows[i]["agentCode"].ObjToString();
                dr = agentsDt.Select("agentCode='" + agentCode + "'");
                if (dr.Length > 0)
                    dt.Rows[i]["agentName"] = dr[0]["firstName"].ObjToString() + " " + dr[0]["lastName"].ObjToString();

                contractValue = DailyHistory.GetContractValue(dt.Rows[i]);
                contractValue = G1.RoundValue(contractValue);
                dt.Rows[i]["contractValue"] = contractValue;
                if (trust.Length > 0)
                {
                    idx = trust.Length - 1;
                    ch = trust.Substring(idx);
                    if (ch.ToUpper() == "I")
                        dt.Rows[i][ibtrustCol] = contractValue;
                    else
                    {
                        if ( contractValue < 0D)
                        {
                            downPayment = dt.Rows[i]["downPayment"].ObjToDouble();
                            contractValue = downPayment;
                            dt.Rows[i]["contractValue"] = downPayment;
                        }
                        dt.Rows[i][sptrustCol] = contractValue;
                    }
                }
                else
                    dt.Rows[i][sptrustCol] = contractValue;
                dt.Rows[i][totalCol] = contractValue;
            }

            cmd = "Select * from `contracts` c JOIN `customers` a on c.`contractNumber` = a.`contractNumber` JOIN `payments` p ON c.`contractNumber` = p.`contractNumber` where p.`downPayment` > '0' AND p.`payDate8` >= '" + date1 + "' AND p.`payDate8` <='" + date2 + "';";
            DataTable ddt = G1.get_db_data(cmd);
            for ( int i=0; i<ddt.Rows.Count; i++)
            {
                contractNumber = ddt.Rows[i]["contractNumber"].ObjToString();
                dr = dt.Select("contractNumber='" + contractNumber + "'");
                if ( dr.Length <= 0 )
                {
                    dt.ImportRow(ddt.Rows[i]);
                }
            }

            if (mainDt != null && ytd)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                    mainDt.ImportRow(dt.Rows[i]);
                dt = mainDt.Copy();
            }
            for ( int i=dt.Rows.Count-1; i>=0; i--)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                deceasedDate = G1.GetSQLDate(dt, i, "deceasedDate");
                ddate = deceasedDate.ObjToDateTime();
                if ( ddate.Year > 100)
                {
                }
                if (!Commission.ShouldCommissionBePaid(dt, i))
                {
                    dt.Rows.RemoveAt(i);
                }
                else
                {
                    contractValue = dt.Rows[i]["contractValue"].ObjToDouble();
                    if (contractValue < 0D)
                        dt.Rows.RemoveAt(i);
                }
            }

            if (runAgents)
            {
                DataView tempview = dt.DefaultView;
                tempview.Sort = "agentName asc, loc asc";
                dt = tempview.ToTable();

                ProcessLapses(startDate, stopDate, dt, dd, ytd);
                ProcessReinstates(startDate, stopDate, dt, dd, ytd);

                return dt;
            }
            else
            {
                DataView tempview = dt.DefaultView;
                tempview.Sort = "loc asc, contractNumber asc";
                dt = tempview.ToTable();
            }

            groupContracts = new DataTable();
            groupContracts.Columns.Add("loc");
            groupContracts.Columns.Add("contracts");

            string contracts = "";
            string oldLoc = "";
            string lastContract = "";
            int lastRow = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                contractValue = dt.Rows[i]["contractValue"].ObjToDouble();
                if (contractValue <= 0D)
                    continue;
                contract = Trust85.decodeContractNumber(contractNumber, ref trust, ref loc);
                dr = dd.Select("keycode='" + loc + "'");
                if (dr.Length > 0)
                {
                    //dt.Rows[i]["Location Name"] = dr[0]["name"].ObjToString();
                    dt.Rows[i]["Location Name"] = dr[0]["LocationCode"].ObjToString();
                }
                else
                    dt.Rows[i]["Location Name"] = loc;
                dt.Rows[i]["loc"] = loc;
                if (oldLoc != loc)
                {
                    if (groupContracts.Rows.Count > 0)
                    {
                        lastRow = groupContracts.Rows.Count - 1;
                        contracts = groupContracts.Rows[lastRow]["contracts"].ObjToString();
                        contracts += lastContract;
                        groupContracts.Rows[lastRow]["contracts"] = contracts;
                    }
                    contracts = "";
                    oldLoc = loc;
                }
                if (String.IsNullOrWhiteSpace(contracts))
                {
                    contracts = contractNumber + " - ";
                    DataRow dRow = groupContracts.NewRow();
                    dRow["contracts"] = contracts;
                    dRow["loc"] = dt.Rows[i]["Location Name"].ObjToString();
                    groupContracts.Rows.Add(dRow);
                }
                lastContract = contractNumber;
            }

            if (groupContracts.Rows.Count > 0)
            {
                lastRow = groupContracts.Rows.Count - 1;
                contracts = groupContracts.Rows[lastRow]["contracts"].ObjToString();
                contracts += lastContract;
                groupContracts.Rows[lastRow]["contracts"] = contracts;
            }

            ProcessLapses(startDate, stopDate, dt, dd, ytd);
            ProcessReinstates(startDate, stopDate, dt, dd, ytd);
            return dt;
        }
        /****************************************************************************************/
        private void LoadUpGroupRows(DataTable dt)
        {
            string location = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                location = dt.Rows[i]["Location Name"].ObjToString();
                DataRow[] dRows = groupContracts.Select("loc='" + location.Trim() + "'");
                if (dRows.Length > 0)
                    dt.Rows[i]["Location Name"] = location + " [ Contracts: " + dRows[0]["contracts"].ObjToString() + " ]";
            }
        }
        /****************************************************************************************/
        private void chkExpand_CheckedChanged(object sender, EventArgs e)
        {
            gridMain.Columns["num"].Visible = false;
            if (runAgents)
            {
                if (chkExpand.Checked)
                {
                    gridMain.Columns["agentName"].GroupIndex = -1;
                    gridMain.Columns["Location Name"].GroupIndex = -1;
                    gridMain.Columns["agentName"].Visible = true;
                    gridMain.Columns["Location Name"].Visible = true;
                    gridMain.Columns["loc"].Visible = false;
                    gridMain.Columns["contractNumber"].Visible = true;
                    gridMain.Columns["contractValue"].Visible = true;
                    gridMain.Columns["contracts"].Visible = false;
                    this.gridMain.ExpandAllGroups();
                }
                else
                {
                    gridMain.Columns["agentName"].GroupIndex = 0;
                    gridMain.Columns["Location Name"].GroupIndex = 1;
                    gridMain.Columns["agentName"].Visible = true;
                    gridMain.Columns["Location Name"].Visible = true;
                    gridMain.Columns["loc"].Visible = false;
                    gridMain.Columns["contractNumber"].Visible = false;
                    gridMain.Columns["contractValue"].Visible = false;
                    gridMain.Columns["contracts"].Visible = false;
                    //                    this.gridMain.CollapseAllGroups();
                    this.gridMain.ExpandGroupLevel(0);
                    //                    this.gridMain.ExpandGroupLevel(1);
                }
            }
            else
            {
                if (chkExpand.Checked)
                {
                    gridMain.Columns["Location Name"].GroupIndex = -1;
                    gridMain.Columns["agentName"].Visible = false;
                    gridMain.Columns["Location Name"].Visible = true;
                    gridMain.Columns["loc"].Visible = true;
                    gridMain.Columns["contractNumber"].Visible = true;
                    gridMain.Columns["contractValue"].Visible = true;
                    gridMain.Columns["contracts"].Visible = false;
                    this.gridMain.ExpandAllGroups();
                }
                else
                {
                    gridMain.Columns["Location Name"].GroupIndex = 0;
                    gridMain.Columns["agentName"].Visible = false;
                    gridMain.Columns["Location Name"].Visible = false;
                    gridMain.Columns["loc"].Visible = false;
                    gridMain.Columns["contractNumber"].Visible = false;
                    gridMain.Columns["contractValue"].Visible = false;
                    gridMain.Columns["contracts"].Visible = false;
                    this.gridMain.CollapseAllGroups();
                }
            }
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
        /****************************************************************************************/
        private void ProcessLapses(DateTime startDate, DateTime stopDate, DataTable dt, DataTable dd, bool ytd)
        {
            DateTime start = startDate;
            start = start.AddMonths(1);
            string date1 = G1.DateTimeToSQLDateTime(start);
            DateTime stop = stopDate;
            stop = stop.AddMonths(1);
            int days = DateTime.DaysInMonth(stop.Year, stop.Month);
            stop = new DateTime(stop.Year, stop.Month, days);
            string date2 = G1.DateTimeToSQLDateTime(stop);
            string contractNumber = "";
            string loc = "";
            string contract = "";
            string trust = "";
            double contractValue = 0D;
            int idx = 0;
            string ch = "";
            string locationName = "";
            string agentCode = "";
            string lapseCol = "lapses";
            if (ytd)
                lapseCol += "ytd";

            string cmd = "Select * from `contracts` c JOIN `customers` a ON c.`contractNumber` = a.`contractNumber` where `lapseDate8` >= '" + date1 + "' AND `lapseDate8` <='" + date2 + "';";
            DataTable dx = G1.get_db_data(cmd);

            int lastRow = dx.Rows.Count;

            for (int i = 0; i < dx.Rows.Count; i++)
            {
                contractNumber = dx.Rows[i]["contractNumber"].ObjToString();
                if ( contractNumber == "E17069LI")
                {
                }
                contract = Trust85.decodeContractNumber(contractNumber, ref trust, ref loc);
                DataRow[] dr = dd.Select("keycode='" + loc + "'");
                if (dr.Length > 0)
                {
                    //locationName = dr[0]["name"].ObjToString();
                    locationName = dr[0]["LocationCode"].ObjToString();
                }
                else
                    locationName = loc;

                DataRow dRow = dt.NewRow();
                dRow["contractNumber"] = contractNumber;
                contractValue = DailyHistory.GetContractValue(dx.Rows[i]);
                contractValue = G1.RoundValue(contractValue);
                dRow[lapseCol] = contractValue;
                dRow["loc"] = loc;
                dRow["Location Name"] = locationName;

                agentCode = dx.Rows[i]["agentCode"].ObjToString();
                if (String.IsNullOrWhiteSpace(agentCode))
                    agentCode = "XX";
                dr = agentsDt.Select("agentCode='" + agentCode + "'");
                if (dr.Length > 0)
                    dRow["agentName"] = dr[0]["firstName"].ObjToString() + " " + dr[0]["lastName"].ObjToString();
                else
                    dRow["agentName"] = agentCode;

                dt.Rows.Add(dRow);
            }
        }
        /****************************************************************************************/
        private void ProcessReinstates(DateTime startDate, DateTime stopDate, DataTable dt, DataTable dd, bool ytd)
        {
            DateTime start = startDate;
            string date1 = G1.DateTimeToSQLDateTime(start);
            DateTime stop = stopDate;
            string date2 = G1.DateTimeToSQLDateTime(stop);
            string contractNumber = "";
            string loc = "";
            string contract = "";
            string trust = "";
            double contractValue = 0D;
            int idx = 0;
            string ch = "";
            string locationName = "";
            string agentCode = "";
            string reinstateCol = "reinstates";
            if (ytd)
                reinstateCol += "ytd";

            string cmd = "Select * from `contracts` c JOIN `customers` a ON c.`contractNumber` = a.`contractNumber` where `reinstateDate8` >= '" + date1 + "' AND `reinstateDate8` <='" + date2 + "';";
            DataTable dx = G1.get_db_data(cmd);

            for (int i = 0; i < dx.Rows.Count; i++)
            {
                contractNumber = dx.Rows[i]["contractNumber"].ObjToString();
                contract = Trust85.decodeContractNumber(contractNumber, ref trust, ref loc);
                DataRow[] dr = dd.Select("keycode='" + loc + "'");
                if (dr.Length > 0)
                {
                    //locationName = dr[0]["name"].ObjToString();
                    locationName = dr[0]["LocationCode"].ObjToString();
                }
                else
                    locationName = loc;

                DataRow dRow = dt.NewRow();
                dRow["contractNumber"] = contractNumber;
                contractValue = DailyHistory.GetContractValue(dx.Rows[i]);
                contractValue = G1.RoundValue(contractValue);
                dRow[reinstateCol] = contractValue;
                dRow["loc"] = loc;
                dRow["Location Name"] = locationName;

                agentCode = dx.Rows[i]["agentCode"].ObjToString();
                if (String.IsNullOrWhiteSpace(agentCode))
                    agentCode = "XX";
                dr = agentsDt.Select("agentCode='" + agentCode + "'");
                if (dr.Length > 0)
                    dRow["agentName"] = dr[0]["firstName"].ObjToString() + " " + dr[0]["lastName"].ObjToString();
                else
                    dRow["agentName"] = agentCode;
                dt.Rows.Add(dRow);
            }
        }
        /***********************************************************************************************/
        private int pageMarginLeft = 0;
        private int pageMarginRight = 0;
        private int pageMarginTop = 0;
        private int pageMarginBottom = 0;
        /***********************************************************************************************/
        private void printPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!chkExpand.Checked)
                gridMain.OptionsPrint.ExpandAllGroups = false;
            else
                gridMain.OptionsPrint.ExpandAllGroups = true;

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
            printableComponentLink1.Landscape = true;

            Printer.setupPrinterMargins(50, 50, 80, 50);

            pageMarginLeft = Printer.pageMarginLeft;
            pageMarginRight = Printer.pageMarginRight;
            pageMarginTop = Printer.pageMarginTop;
            pageMarginBottom = Printer.pageMarginBottom;

            printableComponentLink1.Margins.Left = pageMarginLeft;
            printableComponentLink1.Margins.Right = pageMarginRight;
            printableComponentLink1.Margins.Top = pageMarginTop;
            printableComponentLink1.Margins.Bottom = pageMarginBottom;

            printingSystem1.Document.AutoFitToPagesWidth = 1;

            Font saveFont = gridMain.AppearancePrint.Row.Font;

            if (doSimple)
            {
                Font newFont = new Font(saveFont.FontFamily, 5F);
                gridMain.Appearance.Row.Font = newFont;
            }

            G1.PrintPreview(printableComponentLink1, gridMain);

            //printableComponentLink1.CreateDocument();
            //printableComponentLink1.ShowPreview();

            gridMain.Appearance.Row.Font = saveFont;
        }
        /***********************************************************************************************/
        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!chkExpand.Checked)
                gridMain.OptionsPrint.ExpandAllGroups = false;
            else
                gridMain.OptionsPrint.ExpandAllGroups = true;

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
            printableComponentLink1.Landscape = true;

            Printer.setupPrinterMargins(50, 50, 80, 50);

            pageMarginLeft = Printer.pageMarginLeft;
            pageMarginRight = Printer.pageMarginRight;
            pageMarginTop = Printer.pageMarginTop;
            pageMarginBottom = Printer.pageMarginBottom;

            printableComponentLink1.Margins.Left = pageMarginLeft;
            printableComponentLink1.Margins.Right = pageMarginRight;
            printableComponentLink1.Margins.Top = pageMarginTop;
            printableComponentLink1.Margins.Bottom = pageMarginBottom;

            printingSystem1.Document.AutoFitToPagesWidth = 1;

            printableComponentLink1.CreateDocument();
            printableComponentLink1.PrintDlg();
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

            Printer.DrawQuad(1, 9, 2, 3, "User : " + LoginForm.username, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Center);

            Printer.SetQuadSize(36, 12);
            font = new Font("Ariel", 11, FontStyle.Bold);
            string title = "New Trust Written - By Location";
            if (runAgents || doSimple )
                title = "New Trust Written - By Agent";
            Printer.DrawQuad(14, 8, 11, 4, title, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);


            DateTime date = this.dateTimePicker1.Value;
            string workDate = date.Month.ToString("D2") + "/" + date.Year.ToString("D4");

            workDate = date.ToString("MMMM yyyy");
            Printer.SetQuadSize(36, 12);
            font = new Font("Ariel", 11, FontStyle.Bold);
            Printer.DrawQuad(29, 8, 12, 4, "Month Closing - " + workDate, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);


            //Printer.SetQuadSize(24, 12);
            //font = new Font("Ariel", 9, FontStyle.Regular);
            //Printer.DrawQuad(20, 8, 5, 4, "Month Closing - " + workDate, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);

            //Printer.DrawQuad(16, 8, 3, 4, lblPayment.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(19, 8, 3, 4, lblTrust85.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(22, 8, 3, 4, lblTrust100.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);

            Printer.SetQuadSize(12, 12);
            Printer.DrawQuadBorder(1, 1, 12, 12, BorderSide.All, 1, Color.Black);
            Printer.DrawQuadBorder(12, 1, 1, 12, BorderSide.Right, 1, Color.Black);
        }
        /****************************************************************************************/
        private bool pageBreak = false;
        private void gridMain_BeforePrintRow(object sender, DevExpress.XtraGrid.Views.Printing.CancelPrintRowEventArgs e)
        {
            int rowHandle = e.RowHandle;
            if (gridMain.IsDataRow(rowHandle))
            {
                try
                {
                    DataTable dt = (DataTable)dgv.DataSource;
                    int row = gridMain.GetDataSourceRowIndex(rowHandle);

                    string newPage = dt.Rows[row]["contracts"].ObjToString();
                    if (newPage.ToUpper() == "BREAK")
                    {
                        pageBreak = true;
                        e.Cancel = true;
                    }
                }
                catch ( Exception ex)
                {
                }
            }
        }
        /****************************************************************************************/
        private void gridMain_AfterPrintRow(object sender, DevExpress.XtraGrid.Views.Printing.PrintRowEventArgs e)
        {
            if ( pageBreak )
                e.PS.InsertPageBreak(e.Y);
            pageBreak = false;
        }
        /****************************************************************************************/
        private void chkYTD_CheckedChanged(object sender, EventArgs e)
        {
            if (chkYTD.Checked)
                gridBand1.Visible = true;
            else
                gridBand1.Visible = false;
        }
        /****************************************************************************************/
        private void gridMain_CustomRowFilter(object sender, DevExpress.XtraGrid.Views.Base.RowFilterEventArgs e)
        {
            if (runAgents)
            {
                int row = e.ListSourceRow;
                if (row >= 0)
                {
                    //if (gridMain.IsDataRow(row))
                    //{
                    //    e.Visible = false;
                    //    e.Handled = true;
                    //    return;
                    //}
                }
            }
        }
        /****************************************************************************************/
        private void pictureBox1_Click(object sender, EventArgs e)
        { // Spy Glass
            G1.ShowHideFindPanel(gridMain);
        }
        /****************************************************************************************/
        private void pieChartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            PlotData plotForm = new PlotData(dt);
            plotForm.Show();
        }
        /****************************************************************************************/
        private void btnSummary_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            SummaryView summaryForm = new SummaryView(dt, this.dateTimePicker2.Value );
            summaryForm.Show();
        }
        /****************************************************************************************/
        private double originalSize = 0D;
        private Font mainFont = null;
        private Font newFont = null;
        private Font HeaderFont = null;
        private double originalHeaderSize = 0D;
        private void ScaleCells()
        {
            if (originalSize == 0D)
            {
                //                originalSize = gridMain.Columns["address1"].AppearanceCell.FontSizeDelta.ObjToDouble();
                originalSize = gridMain.Columns["Location Name"].AppearanceCell.Font.Size;
                mainFont = gridMain.Columns["Location Name"].AppearanceCell.Font;
                HeaderFont = gridMain.Appearance.HeaderPanel.Font;
                originalHeaderSize = gridMain.Appearance.HeaderPanel.Font.Size;
            }
            double scale = txtScale.Text.ObjToDouble();
            double size = scale / 100D * originalSize;
            Font font = new Font(mainFont.Name, (float)size);

            for (int i = 0; i < gridMain.Columns.Count; i++)
            {
                gridMain.Columns[i].AppearanceCell.Font = font;
            }
            gridMain.Appearance.GroupFooter.Font = font;
            gridMain.Appearance.FooterPanel.Font = font;
            gridMain.AppearancePrint.FooterPanel.Font = font;
            gridMain.AppearancePrint.GroupFooter.Font = font;
            newFont = font;
            size = scale / 100D * originalHeaderSize;
            font = new Font(HeaderFont.Name, (float)size, FontStyle.Bold);
            for (int i = 0; i < gridMain.Columns.Count; i++)
            {
                gridMain.Columns[i].AppearanceHeader.Font = font;
            }
            //gridMain.Appearance.HeaderPanel.Font = font;
            //gridMain.AppearancePrint.HeaderPanel.Font = font;
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
        private bool doSimple = false;
        private bool doLocDetail = false;
        private void cmbLocationTotals_EditValueChanged(object sender, EventArgs e)
        {
            doSimple = false;
            doLocDetail = false;
            bool byAgent = false;
            bool byLocation = false;
            string what = "";
            string[] locIDs = this.cmbLocationTotals.EditValue.ToString().Split('|');
            for ( int i=0; i<locIDs.Length; i++)
            {
                what = locIDs[i].ObjToString();
                if (what.ToUpper().ToUpper() == "SIMPLE BY AGENT")
                {
                    byAgent = true;
                    break;
                }
                if (what.ToUpper().ToUpper() == "SIMPLE BY LOCATION")
                {
                    byLocation = true;
                    break;
                }
            }
            if ( byLocation )
            {
                doLocDetail = true;
                simpleByLocation();
                return;
            }
            if ( byAgent )
            {
                doSimple = true;
                if ( agentsDt == null )
                    agentsDt = G1.get_db_data("Select * from `agents`");

                gridMain.Columns["agentName"].GroupIndex = -1;
                gridMain.Columns["Location Name"].GroupIndex = -1;

                DateTime start = this.dateTimePicker1.Value;
                DateTime stop = this.dateTimePicker2.Value;

                DataTable dt = LoadData(start, stop, null, false);

                DataView tempview = dt.DefaultView;
                tempview.Sort = "agentName asc, loc asc";
                dt = tempview.ToTable();

                double totalIB = 0D;
                double totalSP = 0D;
                double agentIB = 0D;
                double agentSP = 0D;
                double locIB = 0D;
                double locSP = 0D;
                double totalTotalIB = 0D;
                double totalTotalSP = 0D;
                double IB = 0D;
                double SP = 0D;
                string agent = "";
                string groupAgent = "";
                string loc = "";
                string groupLoc = "";
                bool firstAgent = true;
                bool firstLoc = true;

                bool doTotals = false;

                DataRow dR = null;

                DataTable dx = dt.Clone();

                for ( int i=0; i<dt.Rows.Count; i++)
                {
                    try
                    {
                        agent = dt.Rows[i]["agentName"].ObjToString();
                        if (agent.ToUpper() == "RONNIE KNOTTS")
                        {
                        }
                        loc = dt.Rows[i]["Location Name"].ObjToString();
                        if (String.IsNullOrWhiteSpace(groupAgent))
                        {
                            groupAgent = agent;
                            groupLoc = loc;
                            firstAgent = true;
                            firstLoc = true;
                        }
                        if (i == (dt.Rows.Count - 1))
                        {
                            IB = dt.Rows[i]["ibtrust"].ObjToDouble();
                            SP = dt.Rows[i]["sptrust"].ObjToDouble();

                            locIB += IB;
                            locSP += SP;
                            agent = "";
                            doTotals = true;
                        }

                        if (agent != groupAgent)
                        {
                            if ((locIB + locSP) != 0D)
                            {
                                dR = dx.NewRow();
                                if (firstAgent)
                                    dR["agentName"] = groupAgent;
                                dR["Location Name"] = groupLoc;
                                dR["ibtrust"] = locIB;
                                dR["sptrust"] = locSP;
                                dR["total"] = locIB + locSP;
                                dx.Rows.Add(dR);
                            }

                            agentIB += locIB;
                            agentSP += locSP;

                            if (!firstAgent)
                            {
                                dR = dx.NewRow();
                                //dR["agentName"] = groupAgent;
                                dR["Location Name"] = "                 Agent Totals";
                                dR["ibtrust"] = agentIB;
                                dR["sptrust"] = agentSP;
                                dR["total"] = agentIB + agentSP;
                                dx.Rows.Add(dR);
                            }
                            firstAgent = true;
                            firstLoc = true;

                            if ((locIB + locSP) != 0D)
                            {
                                dR = dx.NewRow();
                                dx.Rows.Add(dR);
                            }

                            totalIB += agentIB;
                            totalSP += agentSP;

                            if (doTotals)
                            {
                                dR = dx.NewRow();
                                dR["Location Name"] = "                 Grand Totals";
                                dR["ibtrust"] = totalIB;
                                dR["sptrust"] = totalSP;
                                dR["total"] = totalIB + totalSP;
                                dx.Rows.Add(dR);
                                break;
                            }
                            agentIB = 0D;
                            agentSP = 0D;
                            locIB = 0D;
                            locSP = 0D;

                            groupAgent = agent;
                            groupLoc = loc;
                        }
                        else if (loc != groupLoc)
                        {
                            dR = dx.NewRow();
                            if (firstAgent)
                                dR["agentName"] = groupAgent;
                            dR["Location Name"] = groupLoc;
                            dR["ibtrust"] = locIB;
                            dR["sptrust"] = locSP;
                            dR["total"] = locIB + locSP;
                            dx.Rows.Add(dR);

                            firstAgent = false;
                            firstLoc = true;

                            agentIB += locIB;
                            agentSP += locSP;

                            locIB = 0D;
                            locSP = 0D;

                            groupLoc = loc;
                        }

                        IB = dt.Rows[i]["ibtrust"].ObjToDouble();
                        SP = dt.Rows[i]["sptrust"].ObjToDouble();

                        locIB += IB;
                        locSP += SP;
                    }
                    catch ( Exception ex)
                    {
                    }
                }

                G1.NumberDataTable(dx);
                dgv.DataSource = dx;
                dgv.Refresh();

                gridMain.Columns["lapses"].Visible = false;
                gridMain.Columns["reinstates"].Visible = false;
                gridMain.Columns["num"].Visible = false;
                gridMain.Columns["loc"].Visible = false;
                gridMain.Columns["Location Name"].Visible = true;
                gridMain.Columns["agentName"].Visible = true;

                gridMain.OptionsView.ShowFooter = false;
                gridMain.OptionsView.ShowBands = false;

                gridMain.OptionsPrint.PrintBandHeader = false;
                gridMain.OptionsPrint.PrintFooter = false;

                gridMain.Appearance.Row.Font = new Font("Tahoma", 9F);
                gridMain.AppearancePrint.Row.Font = new Font("Tahoma", 9F);
                doSimple = true;
            }
        }
        /****************************************************************************************/
        private void simpleByLocation()
        {
            if (agentsDt == null)
                agentsDt = G1.get_db_data("Select * from `agents`");

            gridMain.Columns["agentName"].GroupIndex = -1;
            gridMain.Columns["Location Name"].GroupIndex = -1;

            DateTime start = this.dateTimePicker1.Value;
            DateTime stop = this.dateTimePicker2.Value;

            DataTable dt = LoadData(start, stop, null, false);

            DataView tempview = dt.DefaultView;
            tempview.Sort = "loc asc";
            dt = tempview.ToTable();

            double totalIB = 0D;
            double totalSP = 0D;
            double agentIB = 0D;
            double agentSP = 0D;
            double locIB = 0D;
            double locSP = 0D;
            double totalTotalIB = 0D;
            double totalTotalSP = 0D;
            double IB = 0D;
            double SP = 0D;
            string agent = "";
            string groupAgent = "";
            string loc = "";
            string groupLoc = "";
            bool firstAgent = true;
            bool firstLoc = true;

            string firstName = "";
            string lastName = "";
            string contractNumber = "";
            bool first = true;
            string str = "";

            bool doTotals = false;

            DataRow dR = null;
            DataRow[] dRows = null;

            DataTable dx = dt.Clone();

            DataTable agentTotals = dt.Clone();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                    if ( contractNumber == "C21026LI")
                    {
                    }

                    agent = dt.Rows[i]["agentName"].ObjToString();
                    loc = dt.Rows[i]["Location Name"].ObjToString();
                    if (String.IsNullOrWhiteSpace(groupLoc))
                    {
                        groupLoc = loc;
                        firstLoc = true;
                    }
                    if (i == (dt.Rows.Count - 1))
                        doTotals = true;

                    if (loc != groupLoc || doTotals )
                    {
                        dR = dx.NewRow();
                        dx.Rows.Add(dR);

                        for ( int j=0; j<agentTotals.Rows.Count; j++)
                        {
                            str = agentTotals.Rows[j]["agentName"].ObjToString();
                            //if ( agent.IndexOf ( " Totals" ) < 0 )
                                str += " Totals";
                            agentTotals.Rows[j]["agentName"] = "    " + str;
                            dx.ImportRow(agentTotals.Rows[j]);
                        }

                        agentTotals.Rows.Clear();

                        if (locIB != 0D || locSP != 0D)
                        {
                            dR = dx.NewRow();
                            dR["agentName"] = "";
                            dR["Location Name"] = "        " + groupLoc + " Totals";
                            dR["ibtrust"] = locIB;
                            dR["sptrust"] = locSP;
                            dR["total"] = locIB + locSP;
                            dx.Rows.Add(dR);

                            dR = dx.NewRow();
                            dR["contracts"] = "BREAK";
                            dx.Rows.Add(dR);
                        }

                        firstLoc = true;

                        locIB = 0D;
                        locSP = 0D;
                        groupLoc = loc;

                        agentTotals.Rows.Clear();
                        if (doTotals)
                            break;
                    }

                    IB = dt.Rows[i]["ibtrust"].ObjToDouble();
                    SP = dt.Rows[i]["sptrust"].ObjToDouble();

                    if (IB == 0D && SP == 0D)
                        continue;

                    if ( firstLoc )
                    {
                        dR = dx.NewRow();
                        dR["Location Name"] = groupLoc;
                        //dR["agentName"] = groupLoc;
                        dx.Rows.Add(dR);
                        firstLoc = false;
                    }


                    dRows = agentTotals.Select("agentName='" + agent + "'");
                    if ( dRows.Length <= 0 )
                    {
                        dR = agentTotals.NewRow();
                        if (firstLoc)
                            dR["Location Name"] = loc;
                        dR["agentName"] = agent;
                        dR["ibtrust"] = 0D;
                        dR["sptrust"] = 0D;
                        dR["total"] = 0D;
                        agentTotals.Rows.Add(dR);
                        firstLoc = false;
                    }
                    dRows = agentTotals.Select("agentName='" + agent + "'");
                    if ( dRows.Length > 0 )
                    {
                        dRows[0]["agentName"] = agent;
                        agentIB = dRows[0]["ibtrust"].ObjToDouble();
                        agentIB += IB;
                        dRows[0]["ibtrust"] = agentIB;
                        agentSP = dRows[0]["sptrust"].ObjToDouble();
                        agentSP += SP;
                        dRows[0]["sptrust"] = agentSP;
                        dRows[0]["total"] = agentIB + agentSP;
                    }

                    firstName = dt.Rows[i]["firstName"].ObjToString();
                    if (firstName.Length > 1)
                        firstName = firstName.Substring(0, 1) + ".";
                    lastName = dt.Rows[i]["lastName"].ObjToString();
                    dt.Rows[i]["contracts"] = firstName + " " + lastName;
                    dt.Rows[i]["Location Name"] = "";

                    dx.ImportRow(dt.Rows[i]);


                    locIB += IB;
                    locSP += SP;
                }
                catch (Exception ex)
                {
                }
            }

            G1.NumberDataTable(dx);
            dgv.DataSource = dx;
            dgv.Refresh();

            gridMain.Columns["lapses"].Visible = false;
            gridMain.Columns["reinstates"].Visible = false;
            gridMain.Columns["num"].Visible = false;
            gridMain.Columns["loc"].Visible = false;
            gridMain.Columns["Location Name"].Visible = true;
            gridMain.Columns["agentName"].Visible = true;
            gridMain.Columns["contractNumber"].Visible = true;
            gridMain.Columns["contractValue"].Visible = true;

            gridMain.Columns["contracts"].Caption = "Name";

            gridMain.OptionsView.ShowFooter = false;
            gridMain.OptionsView.ShowBands = false;

            gridMain.OptionsPrint.PrintBandHeader = false;
            gridMain.OptionsPrint.PrintFooter = false;

            gridMain.Appearance.Row.Font = new Font("Tahoma", 9F);
            gridMain.AppearancePrint.Row.Font = new Font("Tahoma", 9F);
            doSimple = true;
        }
        /****************************************************************************************/
        private void gridMain_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (!doSimple && !doLocDetail )
                return;
            if (e.ListSourceRowIndex == DevExpress.XtraGrid.GridControl.InvalidRowHandle)
                return;
            if (doLocDetail)
            { // RAMMA
                DataTable dt = (DataTable)dgv.DataSource;
                int row = e.ListSourceRowIndex;
                string contractNumber = dt.Rows[row]["contractNumber"].ObjToString();
                if (!String.IsNullOrWhiteSpace(contractNumber))
                {
                    if (e.Column.FieldName.ToUpper() == "IBTRUST")
                    {
                        double value = e.DisplayText.Replace(",", "").ObjToDouble();
                        if (value > 0D)
                            e.DisplayText = "        IB";
                        else if (e.DisplayText.Trim() == "0.00")
                            e.DisplayText = "-     ";
                    }
                    else if (e.Column.FieldName.ToUpper() == "SPTRUST")
                    {
                        double value = e.DisplayText.Replace(",", "").ObjToDouble();
                        if (value > 0D)
                            e.DisplayText = "        SP";
                        else if (e.DisplayText.Trim() == "0.00")
                            e.DisplayText = "-     ";
                    }
                    else if (e.Column.FieldName.ToUpper() == "TOTAL")
                    {
                        //double value = e.DisplayText.Replace(",", "").ObjToDouble();
                        //if (value > 0D)
                        //    e.DisplayText = "        SP";
                        //else if (e.DisplayText.Trim() == "0.00")
                        e.DisplayText = " ";
                    }
                }
            }
            else if (doSimple)
            {
                if (e.Column.FieldName.ToUpper() == "IBTRUST")
                {
                    if (e.DisplayText.Trim() == "0.00")
                        e.DisplayText = "-     ";
                }
                else if (e.Column.FieldName.ToUpper() == "SPTRUST")
                {
                    if (e.DisplayText.Trim() == "0.00")
                        e.DisplayText = "-     ";
                }
                else if (e.Column.FieldName.ToUpper() == "TOTAL")
                {
                    if (e.DisplayText.Trim() == "0.00")
                        e.DisplayText = "-     ";
                }
            }
        }
        /****************************************************************************************/
        private void gridMain_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridView View = sender as GridView;
            if (!doSimple)
                return;
            if (e.RowHandle >= 0)
            {
                string location = View.GetRowCellDisplayText(e.RowHandle, View.Columns["Location Name"]);
                if ( location.Trim().ToUpper() == "GRAND TOTALS")
                {
                    Font f = e.Appearance.Font;
                    string name = f.Name.ObjToString();
                    Font font = new Font(name, e.Appearance.Font.Size, FontStyle.Bold);
                    e.Appearance.Font = font;
                }
            }

        }
        /****************************************************************************************/
        private void gridMain_DoubleClick(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            if (!String.IsNullOrWhiteSpace(contract))
            {
                this.Cursor = Cursors.WaitCursor;
                CustomerDetails clientForm = new CustomerDetails(contract);
                clientForm.ShowDialog();
                this.Cursor = Cursors.Default;
            }
        }
        /****************************************************************************************/
        private void gridMain_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName.ToUpper().IndexOf("DATE") >= 0)
            {
                if (e.RowHandle >= 0)
                {
                    DateTime date = e.DisplayText.ObjToDateTime();
                    e.DisplayText = date.Month.ToString("D2") + "/" + date.Day.ToString("D2") + "/" + date.Year.ToString("D4");
                    if (date.Year == 1)
                        e.DisplayText = "";
                }
            }
            if ( e.Column.FieldName.ToUpper() == "CONTRACTS")
            {
                if (e.RowHandle >= 0)
                {
                    if (e.DisplayText.ToUpper() == "BREAK")
                        e.DisplayText = "";
                }
            }
        }
        /****************************************************************************************/
    }
}