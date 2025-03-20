﻿using DevExpress.CodeParser;
using DevExpress.Pdf;
using DevExpress.Utils;
using DevExpress.Utils.Drawing;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraEditors.ViewInfo;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Controls;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraPrinting;
using DevExpress.XtraReports.UI;
using GeneralLib;
using System;
using System.Data;
using System.Drawing;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using GridView = DevExpress.XtraGrid.Views.Grid.GridView;
using iTextSharp.text.pdf;
using System.IO;
using System.Text.RegularExpressions;
using DevExpress.Data;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using MySql.Data.MySqlClient;
/***********************************************************************************************/
namespace SMFS
{
    /***********************************************************************************************/
    public partial class FunManager : DevExpress.XtraEditors.XtraForm
    {
        /***********************************************************************************************/
        private Color menuBackColor = Color.Gray;
        private DataTable workDt = null;
        private DataTable originalDt = null;
        private string workReport = "";
        private string workManager = "";
        private string workLocation = "";
        private string workArranger = "";
        private string workWho = "";
        private string workArrangerFirstName = "";
        private string workArrangerLastName = "";
        private bool workingManagers = false;
        private bool workingArrangers = false;
        private DataTable oDt = null;
        private bool loading = false;
        private bool insurance = false;
        private bool foundLocalPreference = false;
        private GridView originalGridView = null;
        private bool showFooters = true;
        private string serviceLocs = "";
        private bool workPDF = false;
        private DataTable[] Alldbs = new DataTable [200];
        private int dbCount = 0;
        private bool byPass = false;
        private DataTable summaryDt = null;
        private DataTable funeralHomes = null;
        private bool historicBonus = false;
        public static DataTable timJonesDt = null;
        private DataTable badDt = null;
        /***********************************************************************************************/
        public FunManager( DataTable dt, string manager, string who, string location = "" )
        {
            InitializeComponent();
            //workReport = report;
            workDt = dt;
            workManager = manager;
            workLocation = location;
            workWho = who;
        }
        /***********************************************************************************************/
        RepositoryItemDateEdit ri = null;
        private void FunManager_Load(object sender, EventArgs e)
        {
            menuBackColor = menuStrip1.BackColor;

            btnSaveBonus.Hide();

            btnCompare.Hide();
            btnZeros.Hide();

            loading = true;

            barImport.Hide();

            btnRunCommission.Hide();

            ri = new RepositoryItemDateEdit();
            ri.VistaEditTime = DevExpress.Utils.DefaultBoolean.True;
            ri.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTime;
            ri.Mask.UseMaskAsDisplayFormat = true;
            ri.Mask.EditMask = @"yyyy-MM-dd hh-mm";

            string name = G1.GetUserFullName();

            string title = "Funeral Commission for ";
            if (workDt == null)
            {
                if (workWho.ToUpper() == "M")
                {
                    title += " " + workManager + " as Manager";
                    workingManagers = true;
                }
                else if (workWho.ToUpper() == "A")
                {
                    workingArrangers = true;
                    if (!String.IsNullOrWhiteSpace(workArranger))
                        title += " " + workArranger + " as Arranger";
                    else
                        title += " " + workManager + " as Arranger";
                }
            }
            else
            {
                btnRefresh.Hide();
                btnRunCommission.Show();
                btnRunCommission.Text = "Run Commissions";
                string ma = workDt.Rows[0]["ma"].ObjToString();
                if (ma.ToUpper() == "M")
                {
                    title += "All Managers";
                    workingManagers = true;
                }
                else if (ma.ToUpper() == "A")
                {
                    title += "All Arrangers";
                    workingArrangers = true;
                }
            }

            this.Text = title;

            showFooters = true;
            string preference = G1.getPreference(LoginForm.username, "Funerals CB Chooser", "Allow Access");
            if (preference != "YES")
                showFooters = false;

            string prefix = "";
            string suffix = "";
            string firstName = "";
            string lastName = "";
            string mi = "";

            G1.ParseOutName(workManager, ref prefix, ref firstName, ref lastName, ref mi, ref suffix);

            if (!String.IsNullOrWhiteSpace(firstName))
                workArrangerFirstName = firstName;
            if (!String.IsNullOrWhiteSpace(lastName))
                workArrangerLastName = lastName;

            workArranger = firstName + " " + lastName;

            loadLocatons();

            SetupTotalsSummary();

            cmbReportType.Text = workReport;
            loading = false;

            DateTime today = DateTime.Now;
            int days = DateTime.DaysInMonth(today.Year, today.Month);
            this.dateTimePicker2.Value = new DateTime(today.Year, today.Month, days);
            this.dateTimePicker1.Value = new DateTime(today.Year, today.Month, 1);

            this.Refresh();
            gridMain.RefreshEditor(true);

            //cmbSelectColumns_SelectedIndexChanged(cmbSelectColumns, null);

            gridMain.ShowCustomizationForm += GridMain_ShowCustomizationForm;


            if (showFooters)
            {
                //this.gridMain.GroupSummary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[]
                //{
                //new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "balanceDue", this.balanceDue6, "${0:0,0.00}"),
                //new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "grossAmountReceived", this.grossAmountReceived12, "${0:0,0.00}"),
                //new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "amountDiscount", this.amountDiscount6, "${0:0,0.00}"),
                //new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "trustAmountFiled", this.trustAmountFiled2, "${0:0,0.00}"),
                //new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "trustAmountReceived", this.trustAmountReceived12, "${0:0,0.00}"),
                //new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "insAmountFiled", this.insAmountFiled6, "${0:0,0.00}"),
                //new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "insAmountReceived", this.insAmountReceived18, "${0:0,0.00}"),
                //new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "amountGrowth1", this.amountGrowth12, "${0:0,0.00}")});
            }

            SetAllClaims();

            SetupServiceLocs();
            //this.Refresh();
            //gridMain.RefreshEditor(true);
        }
        /****************************************************************************************/
        private void SetupServiceLocs ()
        {
            serviceLocs = "";
            string cmd = "Select * from `funeralhomes` where `manager` = '" + workManager + "' ";
            if (!String.IsNullOrWhiteSpace(workLocation))
                cmd += " AND `LocationCode` = '" + workLocation + "' ";
            cmd += " ;";

            DataTable funDt = G1.get_db_data( cmd );
            if (funDt.Rows.Count <= 0)
                return;

            string atNeedCode = "";
            for ( int i=0; i<funDt.Rows.Count; i++)
            {
                atNeedCode = funDt.Rows[i]["atneedcode"].ObjToString();
                if ( !String.IsNullOrWhiteSpace ( atNeedCode ))
                {
                    if (!String.IsNullOrWhiteSpace(serviceLocs))
                        serviceLocs += ",";
                    serviceLocs += "'" + atNeedCode + "'";
                }
            }
        }
        /****************************************************************************************/
        private void SetAllClaims()
        {
            if (!chkSingle.Checked)
                return;
            //if (1 == 1)
            //    return;

            ClearAllPositions(gridMain);
            int i = 0;
            G1.SetColumnPosition(gridMain, "num", ++i, 50 );
            G1.SetColumnPosition(gridMain, "bad", ++i, 65);
            G1.SetColumnPosition(gridMain, "serviceId", ++i, 55 );
            G1.SetColumnPosition(gridMain, "name", ++i, 100 );
            G1.SetColumnPosition(gridMain, "deceasedDate", ++i, 55);
            G1.SetColumnPosition(gridMain, "funeralType", ++i, 40);
            //G1.SetColumnPosition(gridMain, "netFuneral", ++i, 80);
            G1.SetColumnPosition(gridMain, "gotPackage", ++i, 40);
            G1.SetColumnPosition(gridMain, "casket", ++i, 40);
            G1.SetColumnPosition(gridMain, "vault", ++i, 30);
            G1.SetColumnPosition(gridMain, "currentServices", ++i, 70);
            G1.SetColumnPosition(gridMain, "casketAmount", ++i, 70);
            G1.SetColumnPosition(gridMain, "urn", ++i, 40);
            G1.SetColumnPosition(gridMain, "vaultAmount", ++i, 70);
            G1.SetColumnPosition(gridMain, "cashAdvance", ++i, 70);
            G1.SetColumnPosition(gridMain, "currentPrice", ++i, 80);
            G1.SetColumnPosition(gridMain, "classa", ++i, 70);
            G1.SetColumnPosition(gridMain, "totalDiscount", ++i, 80);
            G1.SetColumnPosition(gridMain, "endingBalance", ++i, 80);
            G1.SetColumnPosition(gridMain, "balanceDue", ++i, 80);
            G1.SetColumnPosition(gridMain, "casketgauge", ++i, 50);
            G1.SetColumnPosition(gridMain, "caskettype", ++i, 50);
            G1.SetColumnPosition(gridMain, "SRVLOC", ++i, 25);
            G1.SetColumnPosition(gridMain, "casketCost", ++i, 70);
            G1.SetColumnPosition(gridMain, "vaultCost", ++i, 70);
            G1.SetColumnPosition(gridMain, "upgrade", ++i, 50);
            G1.SetColumnPosition(gridMain, "premiumType", ++i, 40);
            G1.SetColumnPosition(gridMain, "amountReceived", ++i, 80);
            G1.SetColumnPosition(gridMain, "otherBonuses", ++i, 60);
            G1.SetColumnPosition(gridMain, "netFuneral", ++i, 60);
            G1.SetColumnPosition(gridMain, "casketdesc", ++i, 100);
            //G1.SetColumnPosition(gridMain, "contractNumber", ++i, 60);

            gridMain.Columns["empty"].Visible = false;

            AddSummaryItem("netFuneral");
            AddSummaryItem("currentServices");
            AddSummaryItem("casketAmount");
            AddSummaryItem("vaultAmount");
            AddSummaryItem("cashAdvance");
            AddSummaryItem("currentPrice");
            AddSummaryItem("classa");
            AddSummaryItem("totalDiscount");
            AddSummaryItem("endingBalance");
            AddSummaryItem("balanceDue");
            AddSummaryItem("casketCost");
            AddSummaryItem("vaultCost");
            AddSummaryItem("upgrade");
            AddSummaryItem("amountReceived");
            AddSummaryItem("otherBonuses");

            gridMain.Columns["bad"].Visible = false;

            chkPreparePrint.Show();
            chkPreparePrint.Refresh();

            chkHonorFilter.Show();
            chkHonorFilter.Refresh();

            gridMain.RefreshEditor(true);
        }
        /****************************************************************************************/
        private void AddSummaryItem ( string fieldName )
        {
            bool found = false;
            string field = "";
            for (int i = 0; i < gridMain.GroupSummary.Count; i++)
            {
                field = gridMain.GroupSummary[i].FieldName;
                if ( field == fieldName )
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                if (G1.getGridColumnIndex(gridMain, fieldName) >= 0)
                {
                    GridGroupSummaryItem item = new GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, fieldName, gridMain.Columns[fieldName], "{0:N2}");
                    gridMain.GroupSummary.Add(item);
                }
            }
        }
        /****************************************************************************************/
        private void ClearAllPositions(DevExpress.XtraGrid.Views.BandedGrid.BandedGridView gMain = null)
        {
            if (gMain == null)
                gMain = gridMain;
            string name = "";
            for (int i = 0; i < gMain.Columns.Count; i++)
            {
                name = gMain.Columns[i].Name.ToUpper();
                if ( name != "NUM" )
                    gMain.Columns[i].Visible = false;
                else
                    gMain.Columns[i].Visible = true;
                gridMain.Columns[i].OptionsColumn.FixedWidth = true;
            }
        }
        /****************************************************************************************/
        private void ClearAllPositions(DevExpress.XtraGrid.Views.BandedGrid.AdvBandedGridView gMain = null)
        {
            if (gMain == null)
                gMain = (AdvBandedGridView) gridMain;
            string name = "";
            for (int i = 0; i < gMain.Columns.Count; i++)
            {
                name = gMain.Columns[i].Name.ToUpper();
                if (name != "NUM")
                    gMain.Columns[i].Visible = false;
                else
                    gMain.Columns[i].Visible = true;
                gridMain.Columns[i].OptionsColumn.FixedWidth = true;
            }
        }
        /****************************************************************************************/
        //private void ClearAllPositions(DevExpress.XtraGrid.Views.BandedGrid.BandedGridView gMain = null)
        //{
        //    if (gMain == null)
        //        gMain = gridMain;
        //    string name = "";
        //    for (int i = 0; i < gMain.Columns.Count; i++)
        //    {
        //        name = gMain.Columns[i].Name.ToUpper();
        //        if (name != "NUM")
        //            gMain.Columns[i].Visible = false;
        //        else
        //            gMain.Columns[i].Visible = true;
        //        gridMain.Columns[i].OptionsColumn.FixedWidth = true;
        //    }
        //}
        /***********************************************************************************************/
        private void GridMain_ShowCustomizationForm(object sender, EventArgs e)
        {
            if (!showFooters)
            {
                gridMain.DestroyCustomization();
            }
        }
        /***********************************************************************************************/
        private void loadLocatons()
        {
            string cmd = "Select * from `funeralhomes` order by `keycode`;";
            DataTable locDt = G1.get_db_data(cmd);

            DataTable newLocDt = locDt.Clone();

            string assignedLocations = "";

            cmd = "Select * from `users` where `username` = '" + LoginForm.username + "';";
            DataTable userDt = G1.get_db_data(cmd);
            if ( userDt.Rows.Count > 0 )
                assignedLocations = userDt.Rows[0]["assignedLocations"].ObjToString();

            string locationCode = "";
            string keyCode = "";
            string[] Lines = null;
            string locations = "";
            string location = "";

            for ( int i=locDt.Rows.Count-1; i>=0; i--)
            {
                keyCode = locDt.Rows[i]["keycode"].ObjToString();
                if (keyCode.IndexOf("-") > 0)
                    locDt.Rows.RemoveAt(i);
            }
            for (int i = 0; i < locDt.Rows.Count; i++)
            {
                locationCode = locDt.Rows[i]["locationCode"].ObjToString();
                if (String.IsNullOrWhiteSpace(locationCode))
                    continue;
                Lines = assignedLocations.Split('~');
                for (int j = 0; j < Lines.Length; j++)
                {
                    location = Lines[j].Trim();
                    if (String.IsNullOrWhiteSpace(location))
                        continue;
                    if (location.ToUpper() == locationCode.ToUpper())
                    {
                        location = locDt.Rows[i]["atNeedCode"].ObjToString();
                        locations += location + "|";
                        newLocDt.ImportRow(locDt.Rows[i]);
                    }
                }
            }
            if (!LoginForm.administrator)
                locDt = newLocDt;

            chkComboLocation.Properties.DataSource = locDt;

            locations = locations.TrimEnd('|');
            chkComboLocation.EditValue = locations;
            chkComboLocation.Text = locations;
        }
        /****************************************************************************************/
        private void SetupTotalsSummary()
        {
            if (!showFooters)
            {
                gridMain.GroupSummary.Clear();
                return;
            }
            //AddSummaryColumn("payment", null);
            AddSummaryColumn("amountReceived", null, "Custom" );
            AddSummaryColumn("amountFiled", null, "Custom");
            AddSummaryColumn("custPrice", null, "Custom");
            AddSummaryColumn("custMerchandise", null, "Custom");
            AddSummaryColumn("custServices", null, "Custom");
            AddSummaryColumn("totalDiscount", null, "Custom");
            AddSummaryColumn("currentPrice", null, "Custom");
            AddSummaryColumn("currentMerchandise", null, "Custom");
            AddSummaryColumn("currentServices", null, "Custom");
            AddSummaryColumn("balanceDue", null, "Custom");
            AddSummaryColumn("additionalDiscount", null, "Custom");
            AddSummaryColumn("classa", null, "Custom");
            AddSummaryColumn("grossAmountReceived", null, "Custom");
            AddSummaryColumn("amountDiscount", null, "Custom");
            AddSummaryColumn("amountGrowth1", null, "Custom");
            AddSummaryColumn("cashAdvance", null, "Custom");
            AddSummaryColumn("trustAmountFiled", null, "Custom");
            AddSummaryColumn("trustAmountReceived", null, "Custom");
            AddSummaryColumn("insAmountFiled", null, "Custom");
            AddSummaryColumn("insAmountReceived", null, "Custom");
            AddSummaryColumn("trustPayment", null, "Custom");
            AddSummaryColumn("netFuneral", null, "Custom");
            AddSummaryColumn("cashCheck", null, "Custom");
            AddSummaryColumn("cc", null, "Custom");
            AddSummaryColumn("casketAmount", null, "Custom");
            AddSummaryColumn("vaultAmount", null, "Custom");
            AddSummaryColumn("casketCost", null, "Custom");
            AddSummaryColumn("vaultCost", null, "Custom");
            AddSummaryColumn("endingBalance", null, "Custom");
            AddSummaryColumn("upgrade", null, "Custom");
            AddSummaryColumn("otherBonuses", null, "Custom");
            AddSummaryColumn("urn", null, "Custom");

            gridMain.Columns["SRVLOC"].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Custom;
            gridMain.Columns["SRVLOC"].SummaryItem.DisplayFormat = "{0:N0}";

            //AddSummaryColumn("currentprice", null);
            //AddSummaryColumn("difference", null);
        }
        /****************************************************************************************/
        private void AddSummaryColumn(string columnName, DevExpress.XtraGrid.Views.Grid.GridView gMain = null, string summaryItemType = "", string format = "" )
        {
            if (gMain == null)
                gMain = gridMain;
            if (String.IsNullOrWhiteSpace(format))
                format = "{0:N2}";
            if (summaryItemType.ToUpper() == "CUSTOM")
            {
                gMain.Columns[columnName].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Custom;
                GridSummaryItem item = null;
                bool found = false;
                for ( int i=0; i<gMain.GroupSummary.Count; i++)
                {
                    item = gMain.GroupSummary[i];
                    if ( item.FieldName == columnName)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    gMain.GroupSummary.Add(new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Custom, columnName, gMain.Columns[columnName], format));
                }
            }
            else
                gMain.Columns[columnName].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
            gMain.Columns[columnName].SummaryItem.DisplayFormat = format;
        }
        /*******************************************************************************************/
        private string getLocationQuery()
        {
            DataRow[] dRows = null;
            DataTable locDt = (DataTable) this.chkComboLocation.Properties.DataSource;
            string procLoc = "";
            string jewelLoc = "";
            string[] locIDs = this.chkComboLocation.EditValue.ToString().Split('|');
            for (int i = 0; i < locIDs.Length; i++)
            {
                if (!String.IsNullOrWhiteSpace(locIDs[i]))
                {
                    if (procLoc.Trim().Length > 0)
                        procLoc += ",";
                    procLoc += "'" + locIDs[i].Trim() + "'";
                    dRows = locDt.Select("atneedcode='" + locIDs[i].Trim() + "'");
                    if ( dRows.Length > 0 )
                    {
                        jewelLoc = dRows[0]["merchandiseCode"].ObjToString();
                        if ( !String.IsNullOrWhiteSpace ( jewelLoc ))
                            procLoc += ",'" + jewelLoc.Trim() + "'";
                    }
                }
            }
            return procLoc.Length > 0 ? " serviceLoc IN (" + procLoc + ") " : "";
        }
        /***********************************************************************************************/
        private DataTable LoadData() // Ramma Zamma
        {
            this.Cursor = Cursors.WaitCursor;

            DataTable dt = null;

            string contractsFile = "fcontracts";
            string customersFile = "fcustomers";
            insurance = false;

            workReport = cmbReportType.Text.Trim().ToUpper();

            DateTime date = this.dateTimePicker2.Value;
            DateTime firstDate = new DateTime(date.Year, date.Month, 1);

            try
            {
                string cmd = "";
                //            string cmd = "SELECT * FROM `fcust_extended` e LEFT JOIN `fcontracts` p ON p.`contractNumber` = e.`contractNumber` left join `fcustomers` d ON e.`contractNumber` = d.`contractNumber` LEFT JOIN `icontracts` i ON i.`contractNumber` = e.`contractNumber` LEFT JOIN `icustomers` j ON j.`contractNumber` = e.`contractNumber` WHERE e.`ServiceID` <> '' ";
                cmd = "SELECT * FROM `fcust_extended` e LEFT JOIN `fcontracts` p ON p.`contractNumber` = e.`contractNumber` left join `fcustomers` d ON d.`contractNumber` = e.`contractNumber` LEFT JOIN `icontracts` i ON i.`contractNumber` = e.`contractNumber` LEFT JOIN `icustomers` j ON j.`contractNumber` = e.`contractNumber` ";
                //cmd += " LEFT JOIN `cust_payments` c ON c.`contractNumber` = e.`contractNumber` LEFT JOIN `cust_payment_details` x ON c.`record` = x.`paymentRecord` ";
                //cmd += " WHERE e.`ServiceID` <> '' AND ( e.`serviceLoc` IN (" + serviceLocs + ") OR e.`Funeral Arranger` LIKE '" + workArrangerFirstName + "%' AND e.`Funeral Arranger` LIKE  '%" + workArrangerLastName + "%' )";
                cmd += " WHERE e.`ServiceID` <> '' ";

                if ( !String.IsNullOrWhiteSpace ( serviceLocs ))
                    cmd += " AND ( e.`serviceLoc` IN (" + serviceLocs + ") OR e.`Funeral Arranger` LIKE '" + workArrangerFirstName + "%' AND e.`Funeral Arranger` LIKE  '%" + workArrangerLastName + "%' )";
                else
                    cmd += " AND e.`Funeral Arranger` LIKE '" + workArrangerFirstName + "%' AND e.`Funeral Arranger` LIKE  '%" + workArrangerLastName + "%' ";

                //OR e.`Funeral Arranger` LIKE 'Arthur%' AND e.`Funeral Arranger` LIKE  '%Newman%' )
                string paidInFull = cmbPaidInFull.Text;
                if (!String.IsNullOrWhiteSpace(paidInFull))
                {
                }
                if (chkUseDates.Checked || chkDeceasedDate.Checked)
                {
                    string date1 = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    if ( firstDate != this.dateTimePicker1.Value )
                        date1 = firstDate.ToString("yyyy-MM-dd");

                    string date2 = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    if (chkDeceasedDate.Checked)
                        cmd += " AND ( (p.`deceasedDate` >= '" + date1 + "' AND p.`deceasedDate` <= '" + date2 + "' ) OR ( e.`bonusDate` >= '" + date1 + "' AND e.`bonusDate` <= '" + date2 + "' ) )";
                    else
                        cmd += " AND ( (e.`serviceDate` >= '" + date1 + "' AND e.`serviceDate` <= '" + date2 + "' ) OR ( e.`bonusDate` >= '" + date1 + "' AND e.`bonusDate` <= '" + date2 + "' ) )";
                }

                if (chkBalanceDue.Checked)
                    cmd += " AND e.`balanceDue` <> '0.00' ";

                //cmd += " AND e.`serviceDate` >= '2015-01-01' ";
                //string locations = getLocationQuery();
                //if (!String.IsNullOrWhiteSpace(locations))
                //    cmd += " AND " + locations;
                cmd += " ORDER BY e.`serviceDate` DESC ";
                cmd += ";";

                dt = G1.get_db_data(cmd);

                DataColumn Col1 = dt.Columns.Add("runDate");
                Col1.SetOrdinal(0);// to put the column in position 0;
                string str = this.dateTimePicker2.Value.ToString("yyyyMMdd");
                for ( int i=0; i<dt.Rows.Count; i++)
                    dt.Rows[i]["runDate"] = str;


                if ( !String.IsNullOrWhiteSpace ( serviceLocs ))
                {
                    //DataRow[] dRows = dt.Select("serviceLoc IN (" + serviceLocs + ")");
                    DataRow[] dRows = dt.Select("SRVLOC IN (" + serviceLocs + ")");
                    if (dRows.Length > 0)
                        dt = dRows.CopyToDataTable();
                }
                if ( !String.IsNullOrWhiteSpace ( workLocation ))
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                        dt.Rows[i]["SRVLOC"] = workLocation;

                }

                //dt.Columns.Add("casket");
                //dt.Columns.Add("vault");
                //dt.Columns.Add("serviceAmount", Type.GetType("System.Double"));
                //dt.Columns.Add("casketAmount", Type.GetType("System.Double"));
                //dt.Columns.Add("vaultAmount", Type.GetType("System.Double"));
                //dt.Columns.Add("cashAdvance", Type.GetType("System.Double"));
                dt.Columns.Add("netFuneral", Type.GetType("System.Double"));
                dt.Columns.Add("otherBonuses", Type.GetType("System.Double"));
                dt.Columns.Add("funeralType");
                dt.Columns.Add("bad");




                workReport = cmbReportType.Text.Trim().ToUpper();
                this.Text = "Funerals for Manager " + workManager + " Location " + workLocation;

                string title = "Funeral Commission for ";
                if (workDt == null)
                {
                    if (workWho.ToUpper() == "M")
                        title += " " + workManager + " as Manager for Location " + workLocation;
                    else if (workWho.ToUpper() == "A")
                    {
                        if (!String.IsNullOrWhiteSpace(workArranger))
                            title += " " + workArranger + " as Arranger";
                        else
                            title += " " + workManager + " as Arranger";
                    }
                    this.Text = title;
                }


                SetAllClaims();

                //Trust85.FindContract(dt, "F0372");


                PreProcessData(dt);

                //SetupPaidUpCheck(dt);

                //DetermineLapsed(dt);

                //if (workReport == "CASH BALANCE REPORT")
                //    dt = ProcessCashBalance(dt);

                dt = loadManagers(dt);

                ProcessExcludes(dt);

                CombineLocations(dt, "HH", "HH-TY");
                CombineLocations(dt, "TY", "HH-TY");

                CombineLocations(dt, "LR", "LR-RA");
                CombineLocations(dt, "RA", "LR-RA");

                CombineLocations(dt, "MA", "MA-TV");
                CombineLocations(dt, "TV", "MA-TV");

                CombineLocations(dt, "WC", "WC-WR");
                CombineLocations(dt, "WR", "WC-WR");

                G1.NumberDataTable(dt);
                originalDt = dt.Copy();
                if (workDt == null)
                {
                    dgv.DataSource = dt;
                    ScaleCells();

                    gridMain.ExpandAllGroups();
                }
            }
            catch ( Exception ex)
            {
            }

            this.Cursor = Cursors.Default;

            if ( workDt == null )
            {
                Alldbs[0] = dt;
                dbCount = 1;
                CalcSummaryAverages();
                dgv.DataSource = dt;
            }

            return dt;
        }
        /***********************************************************************************************/
        private DataTable CombineLocations ( DataTable dt, string fromLoc, string toLoc )
        {
            if (dt == null)
                return dt;
            DataRow[] dRows = dt.Select("serviceLoc='" + fromLoc + "'");
            if ( dRows.Length > 0 )
            {
                for (int i = 0; i < dRows.Length; i++)
                    dRows[i]["serviceLoc"] = toLoc;
            }
            return dt;
        }
        /***********************************************************************************************/
        private void ProcessExcludes ( DataTable dt )
        {
            string classification = "";
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                classification = dt.Rows[i]["funeral_classification"].ObjToString();
                if (classification.ToUpper().IndexOf("PICKUP") >= 0)
                    dt.Rows[i]["Exclude"] = "Y";
                else 
                {
                    if (classification.ToUpper().IndexOf("OTHER -") >= 0)
                    {
                        if (classification.ToUpper().IndexOf("TENT AND CHAIR") >= 0)
                            dt.Rows[i]["Exclude"] = "Y";
                    }
                }
            }
        }
        /***********************************************************************************************/
        private DataTable loadManagers ( DataTable dt )
        {
            if (G1.get_column_number(dt, "manager") < 0)
                dt.Columns.Add("manager");

            DataTable funDt = G1.get_db_data("Select * from `funeralhomes`;");
            if (funDt.Rows.Count <= 0)
                return dt;

            string serviceId = "";
            string atneedCode = "";
            string manager = "";
            string contract = "";
            string trust = "";
            string loc = "";

            DataRow[] dRows = null;

            for ( int i=0; i<dt.Rows.Count; i++)
            {
                serviceId = dt.Rows[i]["serviceId"].ObjToString();
                if (String.IsNullOrWhiteSpace(serviceId))
                    continue;

                contract = Trust85.decodeContractNumber(serviceId, ref trust, ref loc);
                if (String.IsNullOrWhiteSpace(loc))
                    continue;

                dRows = funDt.Select("atNeedCode='" + loc + "'");
                if ( dRows.Length > 0 )
                {
                    manager = dRows[0]["manager"].ObjToString();
                    dt.Rows[i]["manager"] = manager;
                }
            }
            return dt;
        }
        /***********************************************************************************************/
        private DataTable trust2013Dt = null;
        private void PullTheData ()
        {
            trust2013Dt = null;
            DateTime date = this.dateTimePicker2.Value;
            string myDate = date.ToString("yyyy-MM-dd");
            string cmd = "Select * from `trust2013r` where `payDate8` <= '" + myDate + "' ORDER BY `payDate8` DESC LIMIT 1;";
            DataTable dx = G1.get_db_data(cmd);
            if ( dx.Rows.Count > 0 )
            {
                myDate = dx.Rows[0]["payDate8"].ObjToDateTime().ToString("yyyy-MM-dd");
                cmd = "Select * from `trust2013r` where `payDate8` = '" + myDate + "' ORDER BY `payDate8`;";
                trust2013Dt = G1.get_db_data(cmd);
            }
        }
        /***********************************************************************************************/
        private string GetFuneralHome ( string serviceId )
        {
            string funeralHome = "";
            if (String.IsNullOrWhiteSpace(serviceId))
                return funeralHome;

            string trust = "";
            string loc = "";
            Trust85.decodeContractNumber(serviceId, ref trust, ref loc);

            DataTable funDt = (DataTable)chkComboLocation.Properties.DataSource;
            DataRow[] dRows = funDt.Select("atneedcode='" + loc + "'");
            if (dRows.Length > 0)
                funeralHome = dRows[0]["LocationCode"].ObjToString();
            return funeralHome;
        }
        ///***********************************************************************************************/
        //private void PreProcessDataOld(DataTable dt)
        //{

        //    //PullTheData();

        //    DataTable localDt = dt.Clone();

        //    string contractNumber = "";
        //    string lastName = "";
        //    string firstName = "";
        //    string payer = "";
        //    string serviceId = "";
        //    bool noFirstName1 = false;
        //    DateTime dueDate = DateTime.Now;
        //    string cmd = "";
        //    DataTable dx = null;
        //    double endingBalance = 0D;
        //    double amountDBR = 0D;
        //    double trustAmountFiled = 0D;
        //    double amountReceived = 0D;
        //    string type = "";
        //    string tbbLoc = "";
        //    string is2000 = "";
        //    string locind = "";
        //    string cnum = "";
        //    DataRow[] dRows = null;
        //    double balanceDue = 0D;
        //    double contractValue = 0D;
        //    double growth = 0D;
        //    double paid = 0D;
        //    string paymentRecord = "";
        //    string trust_policy = "";
        //    DateTime issueDate = DateTime.Now;

        //    double casketCost = 0D;
        //    double vaultCost = 0D;
        //    double urnCost = 0D;
        //    double cashAdvance = 0D;
        //    double totalFuneral = 0D;
        //    double totalDiscount = 0D;
        //    double preneedDiscount = 0D;
        //    double additionalDiscount = 0D;
        //    double netFuneral = 0D;
        //    double classa = 0D;
        //    string funeralType = "";
        //    string SRVLOC = "";
        //    string arranger = "";
        //    bool isManager = false;
        //    bool isArranger = false;
        //    double asService = 0D;
        //    double asCash = 0D;
        //    double asNothing = 0D;
        //    double asMerc = 0D;
        //    double fromService = 0D;
        //    double fromMerc = 0D;
        //    double otherBonuses = 0D;

        //    double currentServices = 0D;
        //    double casketAmount = 0D;
        //    string comment = "";
        //    DataTable tempDt = null;
        //    double pendingBalance = 0D;

        //    double trustPayments = 0D;
        //    double trustPaymentsReceived = 0D;
        //    double trustDifference = 0D;

        //    double insPayments = 0D;
        //    double insPaymentsReceived = 0D;
        //    double insDifference = 0D;
        //    try
        //    {
        //        if (G1.get_column_number(dt, "name") < 0)
        //            dt.Columns.Add("name");
        //        if (G1.get_column_number(dt, "premiumType") < 0)
        //            dt.Columns.Add("premiumType");
        //        if (G1.get_column_number(dt, "paidUp") < 0)
        //            dt.Columns.Add("paidUp");
        //        if (G1.get_column_number(dt, "trustee") < 0)
        //            dt.Columns.Add("trustee");
        //        if (G1.get_column_number(dt, "endingBalance") < 0)
        //            dt.Columns.Add("endingBalance", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "dbr") < 0)
        //            dt.Columns.Add("dbr", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "tbbLoc") < 0)
        //            dt.Columns.Add("tbbLoc");
        //        if (G1.get_column_number(dt, "principleReceived") < 0)
        //            dt.Columns.Add("principleReceived", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "amountGrowth1") < 0)
        //            dt.Columns.Add("amountGrowth1", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "notes") < 0)
        //            dt.Columns.Add("notes");
        //        if (G1.get_column_number(dt, "names") < 0)
        //            dt.Columns.Add("names");
        //        if (G1.get_column_number(dt, "class_a") < 0)
        //            dt.Columns.Add("class_a");
        //        if (G1.get_column_number(dt, "amountFiled") < 0)
        //            dt.Columns.Add("amountFiled", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "dateInsFiled") < 0)
        //            dt.Columns.Add("dateInsFiled");
        //        if (G1.get_column_number(dt, "trustPayment") < 0)
        //            dt.Columns.Add("trustPayment", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "netFuneral") < 0)
        //            dt.Columns.Add("netFuneral", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "cashCheck") < 0)
        //            dt.Columns.Add("cashCheck", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "depositNumber") < 0)
        //            dt.Columns.Add("depositNumber");
        //        if (G1.get_column_number(dt, "depositDate") < 0)
        //            dt.Columns.Add("depositDate");
        //        if (G1.get_column_number(dt, "cc") < 0)
        //            dt.Columns.Add("cc", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "ccDepositNumber") < 0)
        //            dt.Columns.Add("ccDepositNumber");
        //        if (G1.get_column_number(dt, "ccDepositDate") < 0)
        //            dt.Columns.Add("ccDepositDate");
        //        if (G1.get_column_number(dt, "remainingBalance") < 0)
        //            dt.Columns.Add("remainingBalance", Type.GetType("System.Double"));
        //        if (G1.get_column_number(dt, "notes") < 0)
        //            dt.Columns.Add("notes");

        //        for (int i = 0; i < dt.Rows.Count; i++)
        //        {
        //            asService = 0D;
        //            asCash = 0D;
        //            asNothing = 0D;
        //            asMerc = 0D;
        //            fromService = 0D;
        //            fromMerc = 0D;

        //            serviceId = dt.Rows[i]["serviceId"].ObjToString();
        //            if (serviceId == "CT25004")
        //            {
        //            }
        //            isArranger = false;
        //            isManager = false;
        //            //SRVLOC = dt.Rows[i]["serviceLoc"].ObjToString().ToUpper();
        //            //dt.Rows[i]["SRVLOC"] = dt.Rows[i]["serviceLoc"].ObjToString();
        //            arranger = dt.Rows[i]["Funeral Arranger"].ObjToString().ToUpper();
        //            if (arranger.IndexOf(workArrangerFirstName.ToUpper()) >= 0 && arranger.IndexOf(workArrangerLastName.ToUpper()) > 0)
        //                isArranger = true;
        //            if (serviceLocs.ToUpper().IndexOf(SRVLOC) >= 0)
        //                isManager = true;
        //            if (isManager && isArranger)
        //                dt.Rows[i]["funeralType"] = "MA";
        //            else if (isManager)
        //                dt.Rows[i]["funeralType"] = "M";
        //            else
        //                dt.Rows[i]["funeralType"] = "A";

        //            firstName = dt.Rows[i]["firstName"].ObjToString();
        //            lastName = dt.Rows[i]["lastName"].ObjToString();
        //            if (!String.IsNullOrWhiteSpace(firstName) && !String.IsNullOrWhiteSpace(lastName))
        //                dt.Rows[i]["name"] = lastName + ", " + firstName;
        //            contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
        //            if (contractNumber.ToUpper().IndexOf("UI") > 0 || contractNumber.ToUpper().IndexOf("LI") > 0)
        //                dt.Rows[i]["premiumType"] = "M";
        //            else
        //                dt.Rows[i]["premiumType"] = "SP";
        //            dueDate = dt.Rows[i]["dueDate8"].ObjToDateTime();
        //            if (dueDate.ToString("MM/dd/yyyy") == "12/31/2039")
        //                dt.Rows[i]["paidUp"] = "YES";
        //            else
        //                dt.Rows[i]["paidUp"] = "NO";

        //            contractNumber = dt.Rows[i]["contractNumber"].ObjToString();


        //            asService = dt.Rows[i]["asService"].ObjToDouble();
        //            asCash = dt.Rows[i]["asCash"].ObjToDouble();
        //            asNothing = dt.Rows[i]["asNothing"].ObjToDouble();
        //            asMerc = dt.Rows[i]["asMerc"].ObjToDouble();
        //            fromService = dt.Rows[i]["fromService"].ObjToDouble();
        //            fromMerc = dt.Rows[i]["fromMerc"].ObjToDouble();

        //            trustAmountFiled = dt.Rows[i]["trustAmountFiled"].ObjToDouble();
        //            amountReceived = dt.Rows[i]["amountReceived"].ObjToDouble();
        //            balanceDue = dt.Rows[i]["balanceDue"].ObjToDouble();
        //            if (balanceDue > 0D)
        //            {
        //                balanceDue = balanceDue - trustAmountFiled;
        //                if (balanceDue < 0D)
        //                    balanceDue = 0D;
        //                if (balanceDue == 0D)
        //                {
        //                    trustPaymentsReceived = dt.Rows[i]["trustPaymentsReceived"].ObjToDouble();
        //                    //amountReceived += trustAmountFiled;
        //                    amountReceived = trustPaymentsReceived;
        //                    dt.Rows[i]["amountReceived"] = amountReceived;
        //                }
        //            }
        //            else
        //            {
        //                trustPaymentsReceived = dt.Rows[i]["trustPaymentsReceived"].ObjToDouble();
        //                dt.Rows[i]["amountReceived"] = trustPaymentsReceived;
        //            }
        //            dt.Rows[i]["balanceDue"] = balanceDue;
        //            dt.Rows[i]["endingBalance"] = trustAmountFiled;
        //            contractValue = DailyHistory.GetContractValuePlus(dt.Rows[i]);
        //            dt.Rows[i]["principleReceived"] = contractValue - balanceDue;

        //            casketCost = dt.Rows[i]["casketCost"].ObjToDouble();
        //            vaultCost = dt.Rows[i]["vaultCost"].ObjToDouble();
        //            urnCost = dt.Rows[i]["urn"].ObjToDouble();
        //            if (urnCost > 0D)
        //                casketCost += urnCost;
        //            cashAdvance = dt.Rows[i]["cashAdvance"].ObjToDouble();
        //            totalFuneral = dt.Rows[i]["currentPrice"].ObjToDouble();

        //            //totalFuneral = totalFuneral - asNothing;
        //            if (asCash > 0D)
        //            {
        //                cashAdvance = cashAdvance + asCash;
        //                dt.Rows[i]["cashAdvance"] = cashAdvance;
        //            }
        //            if (fromService > 0D)
        //            {
        //                cashAdvance = cashAdvance + fromService;
        //                dt.Rows[i]["cashAdvance"] = cashAdvance;
        //            }
        //            if (fromMerc > 0D)
        //            {
        //                cashAdvance = cashAdvance + fromMerc;
        //                dt.Rows[i]["cashAdvance"] = cashAdvance;
        //            }

        //            currentServices = dt.Rows[i]["currentServices"].ObjToDouble();
        //            casketAmount = dt.Rows[i]["casketAmount"].ObjToDouble();
        //            if (asCash > 0D || asNothing > 0D || asService > 0D || asMerc > 0D || fromService > 0D || fromMerc > 0D)
        //            {
        //                currentServices = currentServices - asCash - asNothing + asService - asMerc - fromService;
        //                dt.Rows[i]["currentServices"] = currentServices;
        //                if (asMerc > 0D)
        //                {
        //                    casketAmount = casketAmount + asMerc;
        //                    dt.Rows[i]["casketAmount"] = casketAmount;
        //                }
        //                if (fromMerc > 0D)
        //                {
        //                    casketAmount = casketAmount - fromMerc;
        //                    dt.Rows[i]["casketAmount"] = casketAmount;
        //                }
        //            }

        //            totalDiscount = dt.Rows[i]["totalDiscount"].ObjToDouble();
        //            preneedDiscount = dt.Rows[i]["preneedDiscount"].ObjToDouble();
        //            preneedDiscount = Math.Abs(preneedDiscount);
        //            if (preneedDiscount > 0D)
        //                totalDiscount = preneedDiscount;
        //            additionalDiscount = dt.Rows[i]["additionalDiscount"].ObjToDouble();
        //            trustPayments = dt.Rows[i]["trustPayments"].ObjToDouble();
        //            trustPaymentsReceived = dt.Rows[i]["trustPaymentsReceived"].ObjToDouble();
        //            if (additionalDiscount > 0D && preneedDiscount <= 0D)
        //            {
        //                totalDiscount += additionalDiscount;
        //                dt.Rows[i]["totalDiscount"] = totalDiscount;
        //            }

        //            trustDifference = trustPayments - trustPaymentsReceived;
        //            if (trustDifference > 0D && trustPaymentsReceived > 0D)
        //            {
        //                totalDiscount += trustDifference;
        //                dt.Rows[i]["totalDiscount"] = totalDiscount;
        //            }
        //            insPayments = dt.Rows[i]["insAmountFiled"].ObjToDouble();
        //            insPaymentsReceived = dt.Rows[i]["insAmountReceived"].ObjToDouble();
        //            insDifference = insPayments - insPaymentsReceived;
        //            if (insDifference > 0D && insPaymentsReceived > 0D)
        //            {
        //                totalDiscount += insDifference;
        //                dt.Rows[i]["totalDiscount"] = totalDiscount;
        //            }
        //            classa = dt.Rows[i]["classa"].ObjToDouble();
        //            netFuneral = totalFuneral - totalDiscount - casketCost - vaultCost - cashAdvance - classa;
        //            if (netFuneral <= 0D)
        //                netFuneral = 0D;
        //            dt.Rows[i]["netFuneral"] = netFuneral;

        //            localDt.Rows.Clear();

        //            localDt.ImportRow(dt.Rows[i]);

        //            tempDt = LoadArrangerData(localDt);
        //            otherBonuses = 0D;
        //            for (int k = 0; k < tempDt.Rows.Count; k++)
        //            {
        //                otherBonuses += tempDt.Rows[k]["commission"].ObjToDouble();
        //            }

        //            if (balanceDue > 0D)
        //            {
        //                comment = dt.Rows[i]["pendingComment"].ObjToString();
        //                if (comment.ToUpper().IndexOf("PENDING TRUST") >= 0)
        //                {
        //                    pendingBalance = GetPendingTrust(dt.Rows[i]["contractNumber"].ObjToString());
        //                    issueDate = dt.Rows[i]["issueDate8"].ObjToDateTime();
        //                    if (issueDate.Year >= 2006)
        //                        otherBonuses += pendingBalance * 0.15D;
        //                    else
        //                        otherBonuses += pendingBalance * 0.50D;
        //                }
        //            }
        //            dt.Rows[i]["otherBonuses"] = otherBonuses;

        //            if (otherBonuses > 0D)
        //            {
        //                netFuneral = netFuneral - otherBonuses;
        //                if (netFuneral <= 0D)
        //                    netFuneral = 0D;
        //                dt.Rows[i]["netFuneral"] = netFuneral;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //    }
        //}
        /***********************************************************************************************/
        private void PreProcessData(DataTable dt)
        {

            //PullTheData();

            DataTable localDt = dt.Clone();

            string contractNumber = "";
            string lastName = "";
            string firstName = "";
            string payer = "";
            string serviceId = "";
            bool noFirstName1 = false;
            DateTime dueDate = DateTime.Now;
            string cmd = "";
            DataTable dx = null;
            double endingBalance = 0D;
            double amountDBR = 0D;
            double trustAmountFiled = 0D;
            double amountReceived = 0D;
            string type = "";
            string tbbLoc = "";
            string is2000 = "";
            string locind = "";
            string cnum = "";
            DataRow[] dRows = null;
            double balanceDue = 0D;
            double contractValue = 0D;
            double growth = 0D;
            double paid = 0D;
            string paymentRecord = "";
            string trust_policy = "";
            DateTime issueDate = DateTime.Now;

            double casketCost = 0D;
            double vaultCost = 0D;
            double urnCost = 0D;
            double cashAdvance = 0D;
            double totalFuneral = 0D;
            double totalDiscount = 0D;
            double preneedDiscount = 0D;
            double additionalDiscount = 0D;
            double netFuneral = 0D;
            double classa = 0D;
            string funeralType = "";
            string SRVLOC = "";
            string arranger = "";
            bool isManager = false;
            bool isArranger = false;
            double asService = 0D;
            double asCash = 0D;
            double asNothing = 0D;
            double asMerc = 0D;
            double fromService = 0D;
            double fromMerc = 0D;
            double otherBonuses = 0D;

            double currentServices = 0D;
            double casketAmount = 0D;
            string comment = "";
            DataTable tempDt = null;
            double pendingBalance = 0D;

            double trustPayments = 0D;
            double trustPaymentsReceived = 0D;
            double trustDifference = 0D;

            double insPayments = 0D;
            double insPaymentsReceived = 0D;
            double insDifference = 0D;
            DataRow dRow = null;
            try
            {
                if (G1.get_column_number(dt, "name") < 0)
                    dt.Columns.Add("name");
                if (G1.get_column_number(dt, "premiumType") < 0)
                    dt.Columns.Add("premiumType");
                if (G1.get_column_number(dt, "paidUp") < 0)
                    dt.Columns.Add("paidUp");
                if (G1.get_column_number(dt, "trustee") < 0)
                    dt.Columns.Add("trustee");
                if (G1.get_column_number(dt, "endingBalance") < 0)
                    dt.Columns.Add("endingBalance", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "dbr") < 0)
                    dt.Columns.Add("dbr", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "tbbLoc") < 0)
                    dt.Columns.Add("tbbLoc");
                if (G1.get_column_number(dt, "principleReceived") < 0)
                    dt.Columns.Add("principleReceived", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "amountGrowth1") < 0)
                    dt.Columns.Add("amountGrowth1", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "notes") < 0)
                    dt.Columns.Add("notes");
                if (G1.get_column_number(dt, "names") < 0)
                    dt.Columns.Add("names");
                if (G1.get_column_number(dt, "class_a") < 0)
                    dt.Columns.Add("class_a");
                if (G1.get_column_number(dt, "amountFiled") < 0)
                    dt.Columns.Add("amountFiled", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "dateInsFiled") < 0)
                    dt.Columns.Add("dateInsFiled");
                if (G1.get_column_number(dt, "trustPayment") < 0)
                    dt.Columns.Add("trustPayment", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "netFuneral") < 0)
                    dt.Columns.Add("netFuneral", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "cashCheck") < 0)
                    dt.Columns.Add("cashCheck", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "depositNumber") < 0)
                    dt.Columns.Add("depositNumber" );
                if (G1.get_column_number(dt, "depositDate") < 0)
                    dt.Columns.Add("depositDate");
                if (G1.get_column_number(dt, "cc") < 0)
                    dt.Columns.Add("cc", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "ccDepositNumber") < 0)
                    dt.Columns.Add("ccDepositNumber");
                if (G1.get_column_number(dt, "ccDepositDate") < 0)
                    dt.Columns.Add("ccDepositDate");
                if (G1.get_column_number(dt, "remainingBalance") < 0)
                    dt.Columns.Add("remainingBalance", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "notes") < 0)
                    dt.Columns.Add("notes");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dRow = dt.Rows[i];
                    ProcessRow ( dt, ref dRow );
                }
            }
            catch ( Exception ex)
            {
            }
        }
        /***********************************************************************************************/
        private void ProcessRow ( DataTable dt, ref DataRow dRow )
        {
            string contractNumber = "";
            string lastName = "";
            string firstName = "";
            string payer = "";
            string serviceId = "";
            bool noFirstName1 = false;
            DateTime dueDate = DateTime.Now;
            string cmd = "";
            DataTable dx = null;
            double endingBalance = 0D;
            double amountDBR = 0D;
            double trustAmountFiled = 0D;
            double amountReceived = 0D;
            double cashReceived = 0D;
            string type = "";
            string tbbLoc = "";
            string is2000 = "";
            string locind = "";
            string cnum = "";
            DataRow[] dRows = null;
            double balanceDue = 0D;
            double contractValue = 0D;
            double growth = 0D;
            double paid = 0D;
            string paymentRecord = "";
            string trust_policy = "";
            DateTime issueDate = DateTime.Now;

            double casketCost = 0D;
            double vaultCost = 0D;
            double urnCost = 0D;
            double cashAdvance = 0D;
            double totalFuneral = 0D;
            double totalDiscount = 0D;
            double preneedDiscount = 0D;
            double additionalDiscount = 0D;
            double compDiscount = 0D;
            double netFuneral = 0D;
            double classa = 0D;
            string funeralType = "";
            string SRVLOC = "";
            string arranger = "";
            bool isManager = false;
            bool isArranger = false;
            double asService = 0D;
            double asCash = 0D;
            double asNothing = 0D;
            double asMerc = 0D;
            double fromService = 0D;
            double fromMerc = 0D;
            double otherBonuses = 0D;
            double grossAmountReceived = 0D;
            double refunds = 0D;
            double money = 0D;
            double dbr = 0D;

            double currentServices = 0D;
            double casketAmount = 0D;
            string comment = "";
            DataTable tempDt = null;
            double pendingBalance = 0D;

            double trustPayments = 0D;
            double trustPaymentsReceived = 0D;
            double trustDifference = 0D;

            double insPayments = 0D;
            double insPaymentsReceived = 0D;
            double insDifference = 0D;
            string casketDesc = "";

            asService = 0D;
            asCash = 0D;
            asNothing = 0D;
            asMerc = 0D;
            fromService = 0D;
            fromMerc = 0D;

            serviceId = dRow["serviceId"].ObjToString();
            if (serviceId == "BN25023")
            {
            }
            isArranger = false;
            isManager = false;
            //SRVLOC = dt.Rows[i]["serviceLoc"].ObjToString().ToUpper();
            //dt.Rows[i]["SRVLOC"] = dt.Rows[i]["serviceLoc"].ObjToString();
            arranger = dRow["Funeral Arranger"].ObjToString().ToUpper();
            if (arranger.IndexOf(workArrangerFirstName.ToUpper()) >= 0 && arranger.IndexOf(workArrangerLastName.ToUpper()) > 0)
                isArranger = true;
            if (serviceLocs.ToUpper().IndexOf(SRVLOC) >= 0)
                isManager = true;
            if (isManager && isArranger)
                dRow["funeralType"] = "MA";
            else if (isManager)
                dRow["funeralType"] = "M";
            else
                dRow["funeralType"] = "A";

            firstName = dRow["firstName"].ObjToString();
            lastName = dRow["lastName"].ObjToString();
            if (!String.IsNullOrWhiteSpace(firstName) && !String.IsNullOrWhiteSpace(lastName))
                dRow["name"] = lastName + ", " + firstName;
            contractNumber = dRow["contractNumber"].ObjToString();
            if (contractNumber.ToUpper().IndexOf("UI") > 0 || contractNumber.ToUpper().IndexOf("LI") > 0)
                dRow["premiumType"] = "M";
            else
                dRow["premiumType"] = "SP";
            dueDate = dRow["dueDate8"].ObjToDateTime();
            if (dueDate.ToString("MM/dd/yyyy") == "12/31/2039")
                dRow["paidUp"] = "YES";
            else
                dRow["paidUp"] = "NO";

            contractNumber = dRow["contractNumber"].ObjToString();


            asService = dRow["asService"].ObjToDouble();
            asCash = dRow["asCash"].ObjToDouble();
            asNothing = dRow["asNothing"].ObjToDouble();
            asMerc = dRow["asMerc"].ObjToDouble();
            fromService = dRow["fromService"].ObjToDouble();
            fromMerc = dRow["fromMerc"].ObjToDouble();
            casketDesc = dRow["casketDesc"].ObjToString();

            trustAmountFiled = dRow["trustAmountFiled"].ObjToDouble();
            trustPayments = dRow["trustPayments"].ObjToDouble();
            amountReceived = dRow["amountReceived"].ObjToDouble();
            cashReceived = dRow["cashReceived"].ObjToDouble();
            insPaymentsReceived = dRow["insAmountReceived"].ObjToDouble();
            amountReceived += cashReceived + insPaymentsReceived;
            balanceDue = dRow["balanceDue"].ObjToDouble();
            if (balanceDue > 0D)
            {
                balanceDue = balanceDue - trustAmountFiled;
                if (balanceDue < 0D)
                    balanceDue = 0D;
                if (balanceDue == 0D)
                {
                    trustPaymentsReceived = dRow["trustPaymentsReceived"].ObjToDouble();
                    //amountReceived += trustAmountFiled;
                    amountReceived = trustPaymentsReceived;
                    dRow["amountReceived"] = amountReceived + cashReceived + insPaymentsReceived;
                }
            }
            else
            {
                trustPaymentsReceived = dRow["trustPaymentsReceived"].ObjToDouble();
                dRow["amountReceived"] = trustPaymentsReceived + cashReceived + insPaymentsReceived;
            }
            dRow["balanceDue"] = balanceDue;
            dRow["endingBalance"] = trustPayments;
            contractValue = DailyHistory.GetContractValuePlus(dRow);
            dRow["principleReceived"] = contractValue - balanceDue;

            casketCost = dRow["casketCost"].ObjToDouble();
            vaultCost = dRow["vaultCost"].ObjToDouble();
            urnCost = dRow["urnCost"].ObjToDouble();
            if (urnCost > 0D)
            {
                //casketCost += urnCost;
                //dRow["casketDesc"] = dRow["urnDesc"].ObjToString();

                vaultCost += urnCost;
                dRow["vaultCost"] = vaultCost;
                dRow["urnCost"] = 0D;
                urnCost = 0D;
            }
            cashAdvance = dRow["cashAdvance"].ObjToDouble();
            //if ( casketDesc.ToUpper().IndexOf ( "INFANT CASKET") >= 0 )
            //{
            //    cashAdvance += dRow["casketCost"].ObjToDouble();
            //    dRow["casketCost"] = 0D;
            //    casketCost = 0D;
            //    dRow["cashAdvance"] = cashAdvance;
            //}
            totalFuneral = dRow["currentPrice"].ObjToDouble();

            //totalFuneral = totalFuneral - asNothing;
            if (asCash > 0D)
            {
                cashAdvance = cashAdvance + asCash;
                dRow["cashAdvance"] = cashAdvance;
            }
            if (fromService > 0D)
            {
                cashAdvance = cashAdvance + fromService;
                dRow["cashAdvance"] = cashAdvance;
            }
            if (fromMerc > 0D)
            {
                cashAdvance = cashAdvance + fromMerc;
                dRow["cashAdvance"] = cashAdvance;
                casketCost = casketCost - fromMerc;
                dRow["casketCost"] = casketCost;
            }

            currentServices = dRow["currentServices"].ObjToDouble();
            casketAmount = dRow["casketAmount"].ObjToDouble(); 
            if (asCash > 0D || asNothing > 0D || asService > 0D || asMerc > 0D || fromService > 0D || fromMerc > 0D)
            {
                currentServices = currentServices - asCash - asNothing + asService - asMerc - fromService;
                dRow["currentServices"] = currentServices;
                if (asMerc > 0D)
                {
                    casketAmount = casketAmount + asMerc;
                    dRow["casketAmount"] = casketAmount;
                }
                if (fromMerc > 0D)
                {
                    casketAmount = casketAmount - fromMerc;
                    dRow["casketAmount"] = casketAmount;
                }
            }




            //if (asCash > 0D || asNothing > 0D || asService > 0D || asMerc > 0D)
            //{
            //    currentServices = currentServices - asCash - asNothing + asService - asMerc;
            //    dRow["currentServices"] = currentServices;
            //    if (asMerc > 0D)
            //    {
            //        casketAmount = casketAmount + asMerc;
            //        dRow["casketAmount"] = casketAmount;
            //    }
            //}

            totalDiscount = dRow["totalDiscount"].ObjToDouble();
            preneedDiscount = dRow["preneedDiscount"].ObjToDouble();
            preneedDiscount = Math.Abs(preneedDiscount);
            if (preneedDiscount > 0D)
                totalDiscount = preneedDiscount;
            additionalDiscount = dRow["additionalDiscount"].ObjToDouble();
            compDiscount = dRow["compDiscount"].ObjToDouble();
            trustPayments = dRow["trustPayments"].ObjToDouble();
            trustPaymentsReceived = dRow["trustPaymentsReceived"].ObjToDouble();
            if (additionalDiscount != 0D )
            {
                totalDiscount += Math.Abs (additionalDiscount);
                dRow["totalDiscount"] = totalDiscount;
            }

            trustDifference = trustPayments - trustPaymentsReceived;
            if (trustDifference > 0D && trustPaymentsReceived > 0D)
            {
                totalDiscount += trustDifference;
                dRow["totalDiscount"] = totalDiscount;
            }
            else
            {
                if (trustPaymentsReceived > 0)
                {
                    totalDiscount -= Math.Abs(trustDifference);
                    dRow["totalDiscount"] = totalDiscount;
                }
            }
            insPayments = dRow["insAmountFiled"].ObjToDouble();
            insPaymentsReceived = dRow["insAmountReceived"].ObjToDouble();
            insDifference = insPayments - insPaymentsReceived;
            if (insDifference > 0D && insPaymentsReceived > 0D)
            {
                totalDiscount += insDifference;
                dRow["totalDiscount"] = totalDiscount;
            }
            else
            {
                if (insPaymentsReceived > 0)
                {
                    totalDiscount -= Math.Abs(insDifference);
                    dRow["totalDiscount"] = totalDiscount;
                }
            }

            if ( totalDiscount < 0D )
            {
                totalDiscount = 0D;
                dRow["totalDiscount"] = 0D;
            }

            //dRow["totalDiscount"] = totalDiscount;
            //classa = dRow["classa"].ObjToDouble();
            //netFuneral = totalFuneral - totalDiscount - casketCost - vaultCost - cashAdvance - classa;
            //if (netFuneral <= 0D)
            //    netFuneral = 0D;
            //dRow["netFuneral"] = netFuneral;

            //DataTable dt = (DataTable)dgv.DataSource;

            otherBonuses = 0D;
            if (totalDiscount <= 0D)
            {
                DataTable localDt = dt.Clone();

                localDt.ImportRow(dRow);

                tempDt = LoadArrangerData(localDt);
                for (int k = 0; k < tempDt.Rows.Count; k++)
                {
                    otherBonuses += tempDt.Rows[k]["commission"].ObjToDouble();
                }
            }

            double trustPending = 0D;

            if (balanceDue > 0D)
            {
                comment = dRow["pendingComment"].ObjToString();
                if (comment.ToUpper().IndexOf("PENDING TRUST") >= 0)
                {
                    pendingBalance = GetPendingTrust(dRow["contractNumber"].ObjToString());
                    issueDate = dRow["issueDate8"].ObjToDateTime();
                    if (issueDate.Year >= 2006)
                        trustPending += pendingBalance * 0.15D;
                    else
                        trustPending += pendingBalance * 0.50D;
                    totalDiscount += trustPending;
                }
            }

            grossAmountReceived = dRow["grossAmountReceived"].ObjToDouble();
            classa = dRow["classa"].ObjToDouble();
            money = dRow["money"].ObjToDouble();
            dbr = dRow["dbr"].ObjToDouble();
            refunds = dRow["refund"].ObjToDouble();
            if (grossAmountReceived > 0D && preneedDiscount <= 0D )
            {
                //totalDiscount = totalFuneral - grossAmountReceived - classa - money - dbr + refunds - additionalDiscount;
                totalDiscount = totalFuneral - grossAmountReceived - classa - money - dbr + refunds;
            }

            if (totalDiscount < 0D)
                totalDiscount = 0D;

            dRow["otherBonuses"] = otherBonuses;

            dRow["totalDiscount"] = totalDiscount;
            classa = dRow["classa"].ObjToDouble();
            netFuneral = totalFuneral - totalDiscount - casketCost - vaultCost - cashAdvance - classa;
            //if (netFuneral <= 0D)
            //    netFuneral = 0D;
            dRow["netFuneral"] = netFuneral;

            if (otherBonuses > 0D)
            {
                netFuneral = netFuneral - otherBonuses;
                if (netFuneral <= 0D)
                    netFuneral = 0D;
                dRow["netFuneral"] = netFuneral;
            }
        }
        /***********************************************************************************************/
        private double GetPendingTrust ( string contractNumber )
        {
            double pending = 0D;
            if (String.IsNullOrWhiteSpace(contractNumber))
                return pending;

            string cmd = "Select * from `cust_payments` where `contractNumber` = '" + contractNumber + "' AND `type` = 'TRUST' AND `status` = 'PENDING';";
            DataTable dt = G1.get_db_data(cmd);
            if ( dt.Rows.Count > 0 )
            {
                pending = dt.Rows[0]["payment"].ObjToDouble();
            }
            return pending;
        }
        /***********************************************************************************************/
        private void LoadServiceDetails ( DataTable dt, int i )
        {
            try
            {
                string contractNumber = dt.Rows[i]["contractNumber"].ObjToString();

                if (String.IsNullOrWhiteSpace(contractNumber))
                    return;

                FunServices serviceForm = new FunServices(contractNumber);
                DataTable serviceDt = serviceForm.funServicesDT;
                if (serviceDt.Rows.Count <= 0)
                    return;

                string merchandise = "";
                string cmd = "";
                string type = "";
                DataTable mDt = null;
                string casketCode = "";
                string ignore = "";
                string[] Lines = null;

                double totalServices = 0D;
                double totalMerchandise = 0D;
                double cashAdvance = 0D;
                double upgrade = 0D;
                double price = 0D;

                for (int j = 0; j < serviceDt.Rows.Count; j++)
                {
                    ignore = serviceDt.Rows[j]["ignore"].ObjToString().ToUpper();
                    if (ignore == "Y")
                        continue;
                    type = serviceDt.Rows[j]["type"].ObjToString().ToUpper();
                    upgrade = serviceDt.Rows[j]["upgrade"].ObjToDouble();
                    price = serviceDt.Rows[j]["price"].ObjToDouble();
                    price = Math.Abs(price);

                    if (type == "CASH ADVANCE")
                        cashAdvance += price;
                    else if (type == "SERVICE")
                        totalServices += price;
                    else if (type == "MERCHANDISE")
                        totalMerchandise += price;
                }

                dt.Rows[i]["serviceAmount"] = totalServices;
                dt.Rows[i]["cashAdvance"] = cashAdvance;
                dt.Rows[i]["totalFuneral"] = totalServices + totalMerchandise + cashAdvance;

                DataRow[] dRows = serviceDt.Select("type='Merchandise'");
                if (dRows.Length <= 0)
                    return;
                try
                {
                    DataTable sDt = dRows.CopyToDataTable();
                    for (int j = 0; j < sDt.Rows.Count; j++)
                    {
                        merchandise = sDt.Rows[j]["service"].ObjToString();
                        if (merchandise.ToUpper().IndexOf("ACKNOW") >= 0)
                            continue;
                        if (merchandise.ToUpper().IndexOf("GRAVE MARKER") >= 0)
                            continue;
                        if (merchandise.ToUpper().IndexOf("REGISTER BOOK") >= 0)
                            continue;
                        if (merchandise.ToUpper().IndexOf("V") == 0)
                        {
                            Lines = merchandise.Split(' ');
                            merchandise = merchandise.Replace(Lines[0].Trim(), "").Trim();
                        }
                        cmd = "SELECT * FROM `casket_master` WHERE `casketDesc` = '" + merchandise + "';";
                        mDt = G1.get_db_data(cmd);
                        if (mDt.Rows.Count <= 0)
                        {
                            Lines = merchandise.Split(' ');
                            merchandise = merchandise.Replace(Lines[0].Trim(), "").Trim();
                            cmd = "SELECT * FROM `casket_master` WHERE `casketDesc` = '" + merchandise + "';";
                            mDt = G1.get_db_data(cmd);
                            if ( mDt.Rows.Count <= 0 )
                                continue;
                        }
                        casketCode = mDt.Rows[0]["casketcode"].ObjToString();
                        if (casketCode.ToUpper().IndexOf("V") == 0)
                        {
                            dt.Rows[i]["vault"] = casketCode;
                            dt.Rows[i]["vaultAmount"] = sDt.Rows[j]["price"].ObjToDouble();
                        }
                        else
                        {
                            dt.Rows[i]["casket"] = casketCode;
                            dt.Rows[i]["casketAmount"] = sDt.Rows[j]["price"].ObjToDouble();
                        }
                    }
                }
                catch (Exception ex)
                {
                }
            }
            catch ( Exception ex)
            {
            }
        }
        /***********************************************************************************************/
        private DataTable ProcessCashBalance ( DataTable dt )
        {
            DataView tempview = dt.DefaultView;
            tempview.Sort = "contractNumber";
            dt = tempview.ToTable();

            string lastContract = "";
            string contractNumber = "";
            string type = "";
            string name = "";
            string record = "";
            string records = "";
            string trust_policy = "";
            string classa = "";
            double classaTotal = 0D;
            double dValue = 0D;
            int lastRow = -1;
            double preneedDiscount = 0D;
            double packageDiscount = 0D;
            string discountType = "";
            string reason = "";
            string status = "";
            string dateInsFiled = "";
            double amountFiled = 0D;
            double trustPayment = 0D;
            double netFuneral = 0D;
            double cashCheck = 0D;
            double creditCard = 0D;
            double remainingBalance = 0D;
            string str = "";
            string depositNumber = "";
            string depositDate = "";
            string ccDepositNumber = "";
            string ccDepositDate = "";
            string notes = "";
            DateTime date = DateTime.Now;

            if (G1.get_column_number(dt, "discountType") < 0)
                dt.Columns.Add("discountType");

            if (G1.get_column_number(dt, "records") < 0)
                dt.Columns.Add("records");

            Trust85.FindContract(dt, "F0372");

            for ( int i=0; i<dt.Rows.Count; i++)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                type = dt.Rows[i]["type"].ObjToString().ToUpper();
                status = dt.Rows[i]["status"].ObjToString().ToUpper();
                if (contractNumber == "L0343" && type == "TRUST" )
                {
                }
                name = dt.Rows[i]["names"].ObjToString();
                record = dt.Rows[i]["record5"].ObjToString();
                if ( contractNumber != lastContract )
                {
                    if (lastRow >= 0)
                    {
                        records = records.TrimEnd(',');
                        dt.Rows[lastRow]["records"] = records;
                        netFuneral = dt.Rows[lastRow]["currentPrice"].ObjToDouble() - dt.Rows[lastRow]["classa"].ObjToDouble() - dt.Rows[lastRow]["totalDiscount"].ObjToDouble();
                        dt.Rows[lastRow]["netFuneral"] = netFuneral;
                    }
                    lastRow = i;
                    lastContract = contractNumber;
                    records = "";
                    if (type == "CASH" || type == "CHECK" || type == "CREDIT CARD")
                        records = record + ",";
                    if (type.ToUpper().IndexOf("INSURANCE") >= 0 || type.ToUpper() == "CLASS A")
                        dt.Rows[i]["names"] = name;
                    else
                        dt.Rows[i]["names"] = "";
                    if (type.ToUpper() == "CLASS A")
                    {
                        trust_policy = dt.Rows[i]["trust_policy"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(trust_policy))
                            dt.Rows[i]["class_a"] = trust_policy;
                    }
                    discountType = "";
                    preneedDiscount = dt.Rows[i]["preneedDiscount"].ObjToDouble();
                    packageDiscount = dt.Rows[i]["packageDiscount"].ObjToDouble();
                    if (preneedDiscount > 0D)
                        discountType = "PN";
                    if (packageDiscount > 0D)
                    {
                        if (!String.IsNullOrWhiteSpace(discountType))
                            discountType += ", ";
                        discountType += "PK";
                    }
                    if ( type.ToUpper() == "DISCOUNT")
                    {
                        reason = dt.Rows[i]["description"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(reason))
                        {
                            if (!String.IsNullOrWhiteSpace(discountType))
                                discountType += ", ";
                            discountType += reason;
                        }
                    }
                    dt.Rows[i]["discountType"] = discountType;
                    if ( status == "FILED")
                    {
                        dt.Rows[i]["dateInsFiled"] = dt.Rows[i]["dateEntered"].ObjToDateTime().ToString("MM/dd/yyyy");
                        dt.Rows[i]["amountFiled"] = dt.Rows[i]["payment"].ObjToDouble();
                    }
                    if (type == "TRUST")
                        dt.Rows[i]["trustPayment"] = dt.Rows[i]["payment"].ObjToDouble();
                    if (type == "CASH" || type == "CHECK")
                    {
                        dt.Rows[i]["cashCheck"] = dt.Rows[i]["payment"].ObjToDouble();
                        depositNumber = dt.Rows[i]["depositNumber2"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(depositNumber))
                            dt.Rows[i]["depositNumber"] = depositNumber;
                        else
                            dt.Rows[i]["depositNumber"] = "";
                        date = dt.Rows[i]["dateReceived"].ObjToDateTime();
                        if (date.Year > 100)
                            dt.Rows[i]["depositDate"] = date.ToString("MM/dd/yyyy");
                        else
                            dt.Rows[i]["depositDate"] = "";
                    }
                    if (type == "CREDIT CARD")
                    {
                        dt.Rows[i]["cc"] = dt.Rows[i]["payment"].ObjToDouble();
                        depositNumber = dt.Rows[i]["depositNumber2"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(depositNumber))
                            dt.Rows[i]["ccDepositNumber"] = depositNumber;
                        else
                            dt.Rows[i]["ccDepositNumber"] = "";
                        date = dt.Rows[i]["dateReceived"].ObjToDateTime();
                        if (date.Year > 100)
                            dt.Rows[i]["ccDepositDate"] = date.ToString("MM/dd/yyyy");
                        else
                            dt.Rows[i]["ccDepositDate"] = "";
                    }

                    notes = dt.Rows[i]["notes2"].ObjToString();
                    if (notes == "-1")
                        notes = "";
                    dt.Rows[i]["notes"] = notes;

                    continue;
                }

                if (type == "CASH" || type == "CHECK" || type == "CREDIT CARD")
                    records += record + ",";
                if (type.ToUpper().IndexOf("INSURANCE") >= 0 || type.ToUpper() == "CLASS A")
                    dt.Rows[lastRow]["names"] += " " + name;
                if ( type.ToUpper() == "CLASS A")
                {
                    trust_policy = dt.Rows[i]["trust_policy"].ObjToString();
                    if (!String.IsNullOrWhiteSpace(trust_policy))
                    {
                        classa = dt.Rows[lastRow]["class_a"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(classa))
                            classa += ", ";
                        classa += trust_policy;
                        dt.Rows[lastRow]["class_a"] = classa;
                        classaTotal = dt.Rows[lastRow]["classa"].ObjToDouble();
                        dValue = dt.Rows[i]["payment"].ObjToDouble();
                        classaTotal += dValue;
                        dt.Rows[lastRow]["classa"] = classaTotal;
                    }
                }
                if ( type == "DISCOUNT")
                {
                    if (type.ToUpper() == "DISCOUNT")
                    {
                        reason = dt.Rows[i]["description"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(reason))
                        {
                            discountType = dt.Rows[lastRow]["discountType"].ObjToString();
                            if (!String.IsNullOrWhiteSpace(discountType))
                                discountType += ", ";
                            discountType += reason;
                        }
                    }
                    dt.Rows[lastRow]["discountType"] = discountType;
                }
                if (status == "FILED")
                {
                    dateInsFiled = dt.Rows[lastRow]["dateInsFiled"].ObjToString();
                    if (!String.IsNullOrWhiteSpace(dateInsFiled))
                        dateInsFiled += ", ";
                    dateInsFiled += dt.Rows[i]["dateEntered"].ObjToDateTime().ToString("MM/dd/yyyy");
                    dt.Rows[lastRow]["dateInsFiled"] = dateInsFiled;
                    amountFiled = dt.Rows[lastRow]["amountFiled"].ObjToDouble();
                    amountFiled += dt.Rows[i]["payment"].ObjToDouble();
                    dt.Rows[lastRow]["amountFiled"] = amountFiled;
                }
                if (type == "TRUST")
                {
                    trustPayment = dt.Rows[lastRow]["trustPayment"].ObjToDouble();
                    trustPayment += dt.Rows[i]["payment"].ObjToDouble();
                    dt.Rows[lastRow]["trustPayment"] = trustPayment;
                }
                if (type == "CASH" || type == "CHECK")
                {
                    cashCheck = dt.Rows[lastRow]["cashCheck"].ObjToDouble() + dt.Rows[i]["payment"].ObjToDouble();
                    dt.Rows[lastRow]["cashCheck"] = cashCheck;

                    str = dt.Rows[i]["depositNumber2"].ObjToString();
                    if (!String.IsNullOrWhiteSpace(str))
                    {
                        depositNumber = dt.Rows[lastRow]["depositNumber"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(depositNumber))
                            depositNumber += ", ";
                        depositNumber += str;
                        dt.Rows[lastRow]["depositNumber"] = depositNumber;

                        date = dt.Rows[i]["dateReceived"].ObjToDateTime();
                        if (date.Year > 100)
                        {
                            depositDate = dt.Rows[lastRow]["depositDate"].ObjToString();
                            if (!String.IsNullOrWhiteSpace(depositDate))
                                depositDate += " ";
                            depositDate += date.ToString("MM/dd/yyyy");
                            dt.Rows[lastRow]["depositDate"] = depositDate;
                        }
                    }
                }

                if (type == "CREDIT CARD" )
                {
                    if (contractNumber == "L0343")
                    {
                    }

                    creditCard = dt.Rows[lastRow]["cc"].ObjToDouble() + dt.Rows[i]["payment"].ObjToDouble();
                    dt.Rows[lastRow]["cc"] = creditCard;

                    depositNumber = dt.Rows[i]["depositNumber2"].ObjToString();
                    if (!String.IsNullOrWhiteSpace(depositNumber))
                    {
                        ccDepositNumber = dt.Rows[lastRow]["ccDepositNumber"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(ccDepositNumber))
                            ccDepositNumber += ", ";
                        ccDepositNumber += depositNumber;
                        dt.Rows[lastRow]["ccDepositNumber"] = ccDepositNumber;
                    }
                    date = dt.Rows[i]["dateReceived"].ObjToDateTime();
                    if (date.Year > 100)
                    {
                        ccDepositDate = dt.Rows[lastRow]["ccDepositDate"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(ccDepositDate))
                            ccDepositDate += ", ";
                        ccDepositDate += date.ToString("MM/dd/yyyy");
                        dt.Rows[lastRow]["ccDepositDate"] = ccDepositDate;
                    }
                }

                notes = dt.Rows[i]["notes2"].ObjToString();
                if (notes == "-1")
                    notes = "";
                if (!String.IsNullOrWhiteSpace(notes))
                {
                    str = dt.Rows[lastRow]["notes"].ObjToString();
                    if (!String.IsNullOrWhiteSpace(str))
                        str += ", ";
                    str += notes;
                    dt.Rows[lastRow]["notes"] = str;
                }

                dt.Rows[i]["contractNumber"] = "";
            }
            for ( int i=(dt.Rows.Count-1); i>=0; i--)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                if (String.IsNullOrWhiteSpace(contractNumber))
                    dt.Rows.RemoveAt(i);
            }
            return dt;
        }
        /***********************************************************************************************/
        private DataTable FilterTrustClaims ( DataTable dt)
        {
            string contractNumber = "";
            string serviceId = "";
            string serviceId1 = "";
            string type = "";
            for ( int i=dt.Rows.Count-1; i>=0; i--)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                if (contractNumber.ToUpper().IndexOf("SX") == 0)
                    dt.Rows.RemoveAt(i);
                else if (contractNumber.ToUpper().IndexOf("ZZ") == 0)
                    dt.Rows.RemoveAt(i);
                else
                {
                    type = dt.Rows[i]["type"].ObjToString().ToUpper();
                    if (type != "TRUST")
                        dt.Rows.RemoveAt(i);
                    else
                    {
                        serviceId = dt.Rows[i]["serviceId"].ObjToString();
                        serviceId1 = dt.Rows[i]["serviceId1"].ObjToString();
                        if (String.IsNullOrWhiteSpace(serviceId))
                            dt.Rows.RemoveAt(i);
                        else if (serviceId != serviceId1)
                            dt.Rows.RemoveAt(i);
                    }
                }
            }
            return dt;
        }
        /***********************************************************************************************/
        private DataTable FilterUnityClaims(DataTable dt)
        {
            string contractNumber = "";
            string serviceId = "";
            string serviceId1 = "";
            string paidFrom = "";
            Trust85.FindContract(dt, "SX21096");
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                //if (contractNumber.ToUpper().IndexOf("SX") == 0)
                //    dt.Rows.RemoveAt(i);
                if (contractNumber.ToUpper().IndexOf("ZZ") == 0)
                    dt.Rows.RemoveAt(i);
                else
                {
                    paidFrom = dt.Rows[i]["paidFrom"].ObjToString();
                    if ( paidFrom.ToUpper() != "UNITY")
                        dt.Rows.RemoveAt(i);
                }
            }
            Trust85.FindContract(dt, "SX21096");
            return dt;
        }
        /***********************************************************************************************/
        private DataTable FilterNonUnityClaims(DataTable dt)
        {
            string contractNumber = "";
            string serviceId = "";
            string serviceId1 = "";
            Trust85.FindContract(dt, "L0343");
            string type = "";
            string paidFrom = "";
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                if (contractNumber.ToUpper().IndexOf("SX") == 0)
                    dt.Rows.RemoveAt(i);
                else if (contractNumber.ToUpper().IndexOf("ZZ") == 0)
                    dt.Rows.RemoveAt(i);
                else
                {
                    type = dt.Rows[i]["type1"].ObjToString().ToUpper();
                    paidFrom = dt.Rows[i]["paidFrom"].ObjToString().ToUpper();
                    if (paidFrom.ToUpper() == "FILED DIRECT")
                        continue;
                    if (type == "3RD PARTY")
                        continue;
                    dt.Rows.RemoveAt(i);
                }
            }
            return dt;
        }
        /***********************************************************************************************/
        private DataTable FilterCashBalance(DataTable dt)
        {
            string contractNumber = "";
            string serviceId = "";
            string serviceId1 = "";
            string type = "";
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                if (contractNumber.ToUpper().IndexOf("SX") == 0)
                    dt.Rows.RemoveAt(i);
                else if (contractNumber.ToUpper().IndexOf("ZZ") == 0)
                    dt.Rows.RemoveAt(i);
                else
                {
                    type = dt.Rows[i]["type"].ObjToString().ToUpper();
                    if (type != "CASH" && type != "CHECK" && type != "CREDIT CARD")
                        dt.Rows.RemoveAt(i);
                }
            }
            return dt;
        }
        /***********************************************************************************************/
        private void DetermineLapsed(DataTable dt)
        {
            if (G1.get_column_number(dt, "lapsed1") < 0)
                return;
            string lapse = "";
            string lapse1 = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                lapse = dt.Rows[i]["lapsed"].ObjToString();
                lapse1 = dt.Rows[i]["lapsed1"].ObjToString();
                if (String.IsNullOrWhiteSpace(lapse))
                    lapse = " ";
                lapse += lapse1;
                dt.Rows[i]["lapsed"] = lapse;
            }
        }
        /***********************************************************************************************/
        private void CalcPaid(DataTable dt)
        {
            double serviceTotal = 0D;
            double merchandiseTotal = 0D;
            double totalPurchase = 0D;
            double balanceDue = 0D;
            double paid = 0D;
            double totalPaid = 0D;
            double contractValue = 0D;
            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    serviceTotal = dt.Rows[i]["serviceTotal"].ObjToDouble();
                    merchandiseTotal = dt.Rows[i]["merchandiseTotal"].ObjToDouble();
                    balanceDue = dt.Rows[i]["balanceDue"].ObjToDouble();
                    totalPurchase = serviceTotal + merchandiseTotal;
                    paid = totalPurchase - balanceDue;
                    dt.Rows[i]["paid"] = paid;
                    dt.Rows[i]["purchase"] = totalPurchase;
                    dt.Rows[i]["contractValue"] = DailyHistory.GetContractValue(dt.Rows[i]);
                    totalPaid += paid;
                }
            }
            catch (Exception ex)
            {
            }
        }
        /***********************************************************************************************/
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if ( dgv.Visible )
                G1.ShowHideFindPanel(gridMain);
            else if ( dgv2.Visible )
                G1.ShowHideFindPanel(gridMain2);
            else if (dgv3.Visible)
                G1.ShowHideFindPanel(gridMain3);
        }
        /***********************************************************************************************/
        private void SetupPaidUpCheck ( DataTable dt)
        {
            DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit selectnew = this.repPaidCheckEdit;
            selectnew.NullText = "";
            selectnew.ValueChecked = "1";
            selectnew.ValueUnchecked = "0";
            selectnew.ValueGrayed = "";
            string paid = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                paid = dt.Rows[i]["paidInFull1"].ObjToString();
                if ( paid == "1" )
                    dt.Rows[i]["paidInFull1"] = "1";
                else
                    dt.Rows[i]["paidInFull1"] = "0";
            }
        }
        /***********************************************************************************************/
        private void repCheckEdit1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = (DataTable)(dgv.DataSource);
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            string value = dr["agreement"].ObjToString();
            string record = dr["picRecord"].ObjToString();
            if (value == "1")
            {
                string filename = "";
                string title = "Agreement for (" + contract + ") ";
                string cmd = "Select * from `customers` where `contractNumber` = '" + contract + "';";
                DataTable dx = G1.get_db_data(cmd);
                if (dx.Rows.Count > 0)
                {
                    filename = dx.Rows[0]["agreementFile"].ObjToString();
                    string firstName = dx.Rows[0]["firstName"].ObjToString();
                    string lastName = dx.Rows[0]["lastName"].ObjToString();
                    title = "Agreement for (" + contract + ") " + firstName + " " + lastName;
                    if (!String.IsNullOrWhiteSpace(record))
                    {
                        if (!String.IsNullOrWhiteSpace(record))
                            Customers.ShowPDfImage(record, title, filename);
                    }
                }
            }
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            historicBonus = false;
            LoadData();
            SetAllClaims();
            btnRunCommission.Show();
            btnRunCommission.Refresh();
        }
        /***********************************************************************************************/
        private int pageMarginLeft = 0;
        private int pageMarginRight = 0;
        private int pageMarginTop = 0;
        private int pageMarginBottom = 0;
        /***********************************************************************************************/
        private void printPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printPreview(false);
        }
        /***********************************************************************************************/
        private int printRow = 0;
        private void printPreview ( bool batch = true )
        {
            printRow = 0;

            if (this.components == null)
                this.components = new System.ComponentModel.Container();

            DevExpress.XtraPrinting.PrintingSystem printingSystem1 = new DevExpress.XtraPrinting.PrintingSystem(this.components);
            DevExpress.XtraPrinting.PrintableComponentLink printableComponentLink1 = new DevExpress.XtraPrinting.PrintableComponentLink(this.components);

            printingSystem1.Links.AddRange(new object[] {
            printableComponentLink1});


            printableComponentLink1.Component = dgv;
            if ( dgv2.Visible )
                printableComponentLink1.Component = dgv2;
            else if (dgv3.Visible)
                printableComponentLink1.Component = dgv3;
            else if (dgv4.Visible)
                printableComponentLink1.Component = dgv4;

            printableComponentLink1.PrintingSystemBase = printingSystem1;

            printableComponentLink1.EnablePageDialog = true;

            printableComponentLink1.CreateDetailHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateDetailHeaderArea);
            printableComponentLink1.CreateMarginalHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateMarginalHeaderArea);
            printableComponentLink1.BeforeCreateAreas += new System.EventHandler(this.printableComponentLink1_BeforeCreateAreas);
            printableComponentLink1.AfterCreateAreas += new System.EventHandler(this.printableComponentLink1_AfterCreateAreas);
            printableComponentLink1.Landscape = true;

//            Printer.setupPrinterMargins(50, 50, 80, 50);
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

            if (workPDF && batch )
            {
                string filename = "";
                //string filename = path + @"\" + workReport + "_" + today.Year.ToString("D4") + today.Month.ToString("D2") + today.Day.ToString("D2") + ".pdf";

                filename = @"C:/rag/pdfDaily.pdf";
                //filename = workPDFfile;
                if (File.Exists(filename))
                {
                    File.SetAttributes(filename, FileAttributes.Normal);
                    File.Delete(filename);
                }
                printableComponentLink1.ExportToPdf(filename);
            }
            else
                printableComponentLink1.ShowPreview();

        }
        /***********************************************************************************************/
        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printRow = 0;

            if (this.components == null)
                this.components = new System.ComponentModel.Container();

            DevExpress.XtraPrinting.PrintingSystem printingSystem1 = new DevExpress.XtraPrinting.PrintingSystem(this.components);
            DevExpress.XtraPrinting.PrintableComponentLink printableComponentLink1 = new DevExpress.XtraPrinting.PrintableComponentLink(this.components);

            printingSystem1.Links.AddRange(new object[] {
            printableComponentLink1});

            printableComponentLink1.Component = dgv;
            if (dgv2.Visible)
                printableComponentLink1.Component = dgv2;
            else if (dgv3.Visible)
                printableComponentLink1.Component = dgv3;
            else if (dgv4.Visible)
                printableComponentLink1.Component = dgv4;

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

            font = new Font("Ariel", 10, FontStyle.Regular);
            string title = this.Text;
            if (dgv2.Visible)
                title += " Services";
            else if (dgv3.Visible)
                title += " Merchandise";
            else if (dgv4.Visible)
                title += " Differences";

            Printer.DrawQuad(5, 8, 6, 4, title, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);

            //            Printer.DrawQuadTicks();
            DateTime date = DateTime.Now;
            string workDate = date.Month.ToString("D2") + "/" + date.Year.ToString("D4");
            Printer.SetQuadSize(24, 12);
            font = new Font("Ariel", 9, FontStyle.Regular);
            //Printer.DrawQuad(20, 8, 5, 4, "Report Month:" + workDate, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(16, 8, 3, 4, lblPayment.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(19, 8, 3, 4, lblTrust85.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(22, 8, 3, 4, lblTrust100.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);

            Printer.SetQuadSize(12, 12);
            Printer.DrawQuadBorder(1, 1, 12, 12, BorderSide.All, 1, Color.Black);
            Printer.DrawQuadBorder(12, 1, 1, 12, BorderSide.Right, 1, Color.Black);
        }
        /***********************************************************************************************/
        private void gridMain_DoubleClick_2(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            if (!String.IsNullOrWhiteSpace(contract))
            {
                this.Cursor = Cursors.WaitCursor;

                Form form = G1.IsFormOpen("EditCust", contract);
                if (form != null)
                {
                    form.Show();
                    form.WindowState = FormWindowState.Maximized;
                    form.Visible = true;
                    form.BringToFront();
                }
                else
                {
                    EditCust custForm = new EditCust(contract);
                    custForm.Tag = contract;
                    //custForm.custClosing += CustForm_custClosing;
                    custForm.Show();
                }
                //FunPayments editFunPayments = new FunPayments(null, contract, "", false, true );
                //editFunPayments.Show();

                //CustomerDetails clientForm = new CustomerDetails(contract);
                //clientForm.Show();
                this.Cursor = Cursors.Default;
            }
        }
        /***********************************************************************************************/
        private void CustForm_custClosing(string contractNumber, double amountFiled, double amountReceived)
        {
            string extendedRecord = "";
            string record = "";
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt == null)
                return;
            string cmd = "SELECT *, (SELECT `pdfimages`.`record` FROM `pdfimages` WHERE `pdfimages`.`contractNumber` = e.`contractNumber` ) AS `picRecord` FROM `fcust_extended` e LEFT JOIN `fcontracts` p ON p.`contractNumber` = e.`contractNumber` left join `fcustomers` d ON e.`contractNumber` = d.`contractNumber` WHERE e.`contractNumber` = '" + contractNumber + "';";
            //cmd += " AND e.`record` = '" + record + "';";

            DataTable dx = G1.get_db_data(cmd);
            if (dx.Rows.Count > 0)
            {
                extendedRecord = dx.Rows[0]["record"].ObjToString();
                record = extendedRecord;
                //G1.NumberDataTable(dt);
                //SetupAgreementIcon(dt);
                //CalcPaid(dt);
                //DetermineLapsed(dt);
                //if (oDt == null)
                //    oDt = dt.Copy();
                //originalDt = dt.Copy();
                //dgv.DataSource = dt;


                string mRecord = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    mRecord = dt.Rows[i]["record"].ObjToString();
                    if (mRecord == record)
                    {
                        G1.HardCopyDtRow(dx, 0, dt, i);
                        break;
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void SetupSelectedColumns(string procType, string group, DevExpress.XtraGrid.GridControl dgv)
        {
            if (String.IsNullOrWhiteSpace(group))
                return;
            if (String.IsNullOrWhiteSpace(procType))
                procType = "AR " + workReport;
            string user = LoginForm.username;
            if (group.ToUpper().IndexOf("(C)") >= 0)
            {
                user = "Common";
                group = group.Replace("(C) ", "");
            }
            string cmd = "Select * from procfiles where name = '" + group + "' and ProcType = '" + procType + "' AND `user` = '" + user + "' order by seq";
            DataTable dt = G1.get_db_data(cmd);
            if (dt.Rows.Count <= 0)
            {
                cmd = "Select * from procfiles where name = '" + group + "' and ProcType = '" + procType + "' AND ( `user` = 'Common' OR `user` = '' ) order by seq";
                dt = G1.get_db_data(cmd);
            }
            if (dt.Rows.Count <= 0)
                return;
            DevExpress.XtraGrid.Views.Grid.GridView gridMain = (DevExpress.XtraGrid.Views.Grid.GridView)dgv.MainView;
            for (int i = 0; i < gridMain.Columns.Count; i++)
                gridMain.Columns[i].Visible = false;
            DataTable ddx = (DataTable)dgv.DataSource;
            int idx = 0;
            string name = "";
            int index = 0;
            for (int i = (dt.Rows.Count - 1); i >= 0; i--)
            {
                name = dt.Rows[i]["Description"].ToString();
                index = dt.Rows[i]["seq"].ObjToInt32();
                try
                {
                    idx = G1.get_column_number(gridMain, name);
                    if (idx >= 0)
                        gridMain.Columns[name].Visible = true;
                }
                catch
                {
                }
            }
        }
        /***********************************************************************************************/
        private void cmbSelectColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loading)
                return;
            ComboBox combo = (ComboBox)sender;
            string comboName = combo.Text;
            string skinName = "";
            if (!String.IsNullOrWhiteSpace(comboName))
            {
                SetupSelectedColumns("AR", comboName, dgv);
                string name = "AR " + comboName;
                foundLocalPreference = G1.RestoreGridLayout(this, this.dgv, gridMain, LoginForm.username, name, ref skinName);
                SetupTotalsSummary();
                gridMain.OptionsView.ShowFooter = showFooters;
            }
            else
            {
                SetupSelectedColumns("AR", "Primary", dgv);
                string name = "AR Primary";
                foundLocalPreference = G1.RestoreGridLayout(this, this.dgv, gridMain, LoginForm.username, name, ref skinName);
                gridMain.OptionsView.ShowFooter = showFooters;
                SetupTotalsSummary();
            }

            CleanupFieldColumns();

            string format = chkComboLocation.Text;
            if (!String.IsNullOrWhiteSpace(format))
                chkComboLocation_EditValueChanged(null, null);
            if (chkSort.Checked)
                ForceGroups();
        }
        /***********************************************************************************************/
        private void CleanupFieldColumns()
        {
            if (LoginForm.classification.ToUpper() != "FIELD")
                return;
            gridMain.Columns["amountGrowth1"].Visible = false;
            gridMain.Columns["amountDiscount"].Visible = false;
            gridMain.Columns["contractNumber"].Visible = false;
        }
        /***********************************************************************************************/
        void sform_Done(DataTable dt)
        {
            string name = cmbSelectColumns.Text.Trim();
            if (String.IsNullOrWhiteSpace(name))
                name = "Primary";
            string saveName = "AR " + name;
            string skinName = "";
            SetupSelectedColumns("AR", name, dgv);
            foundLocalPreference = G1.RestoreGridLayout(this, this.dgv, gridMain, LoginForm.username, saveName, ref skinName);
            gridMain.OptionsView.ShowFooter = showFooters;
            SetupTotalsSummary();
            string field = "";
            string select = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                field = dt.Rows[i]["field"].ObjToString();
                select = dt.Rows[i]["select"].ObjToString();
                if (G1.get_column_number(gridMain, field) >= 0)
                {
                    if (select == "0")
                        gridMain.Columns[field].Visible = false;
                    else
                        gridMain.Columns[field].Visible = true;
                }
            }
            dgv.Refresh();
            this.Refresh();
        }
        /***********************************************************************************************/
        private void btnSelectColumns_Click_1(object sender, EventArgs e)
        {
            string actualName = cmbSelectColumns.Text;
            if (actualName.ToUpper().IndexOf("(C)") >= 0 && !LoginForm.administrator)
            {
                MessageBox.Show("***Warning*** You do not have permission to modify a Common Display Format!", "Display Format Dialog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            //            SelectColumns sform = new SelectColumns(dgv, "Funerals", "Primary", actualName);
            string user = LoginForm.username;
            SelectDisplayColumns sform = new SelectDisplayColumns(dgv, "AR " + workReport , "Primary", actualName, LoginForm.username);
            sform.Done += new SelectDisplayColumns.d_void_selectionDone(sform_Done);
            sform.Show();
        }
        /***********************************************************************************************/
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string name = cmbSelectColumns.Text.Trim();
            if (String.IsNullOrWhiteSpace(name))
                name = "Primary";
            string saveName = "AR " + workReport + " " + name;
            G1.SaveLocalPreferences(this, gridMain, LoginForm.username, saveName);
        }
        /***********************************************************************************************/
        private DataTable groupDt = null;
        private bool calcSummary = false;
        private int columnCount = 0;
        private bool allowSummary = false;
        /***********************************************************************************************/
        //private double groupTotal = 0D;
        private void gridMain_CustomSummaryCalculate(object sender, DevExpress.Data.CustomSummaryEventArgs e)
        {
            //if (!allowSummary)
            //    return;

            string field = (e.Item as GridSummaryItem).FieldName.ObjToString();
            if (groupDt == null)
            {
                groupDt = new DataTable();
                groupDt.Columns.Add ( "location");
            }

            if (G1.get_column_number(groupDt, field) < 0)
                groupDt.Columns.Add(field, Type.GetType("System.Double"));

            DataTable dt = null;
            DataTable newDt = null;
            int row = 0;
            string location = "";
            DataRow dRow = null;
            DataRow[] dRows = null;
            DataRow[] xRows = null;
            double dValue = 0D;
            double dTotal = 0D;
            double aValue = 0D;
            bool pass = true;
            try
            {
                dt = (DataTable)dgv.DataSource;
                if (dt.Rows.Count <= 0)
                    return;

                row = e.RowHandle;
                row = gridMain.GetDataSourceRowIndex(row);
                if (row < 0)
                    return;
                location = dt.Rows[row]["SRVLOC"].ObjToString();
                location = dt.Rows[row]["serviceLOC"].ObjToString();
                if (String.IsNullOrWhiteSpace(location))
                    return;
                dRows = groupDt.Select("location='" + location + "'");
                if ( field == "currentServices")
                {

                }
            }
            catch ( Exception ex)
            {
            }

            try
            {
                if (dRows.Length <= 0)
                {
                    dRow = groupDt.NewRow();
                    dRow["location"] = location;
                    groupDt.Rows.Add(dRow);

                    dRows = groupDt.Select("location='" + location + "'");
                }

                if (e.SummaryProcess == CustomSummaryProcess.Start)
                {
                    dRows[0][field] = 0D;

                    for (int i = 0; i < groupDt.Rows.Count; i++)
                        groupDt.Rows[i][field] = 0D;
                    return;
                }
                else if (e.SummaryProcess == CustomSummaryProcess.Calculate)
                {
                    string exclude = dt.Rows[row]["exclude"].ObjToString();
                    if ( chkHonorFilter.Checked )
                    {
                        if (location == "BN")
                        {
                        }
                        string xxxx = dt.Rows[row]["bonusDate"].ObjToString();
                        if (xxxx != "0/0/0000")
                        {
                            DateTime bonusDate = dt.Rows[row]["bonusDate"].ObjToDateTime();
                            if (bonusDate.Year > 1000)
                            {
                                if (bonusDate < this.dateTimePicker1.Value || bonusDate > this.dateTimePicker2.Value)
                                    exclude = "Y";
                            }
                        }
                    }
                    aValue = dt.Rows[row][field].ObjToDouble();
                    if (String.IsNullOrWhiteSpace(exclude))
                    {
                        if (field == "SRVLOC")
                        {
                            dValue = dRows[0][field].ObjToDouble();
                            dValue += 1D;
                            dRows[0][field] = dValue;
                        }
                        else
                        {
                            dValue = dRows[0][field].ObjToDouble();
                            dValue += aValue;
                            dRows[0][field] = dValue;
                        }
                    }
                    return;
                }
                else if (e.IsGroupSummary )
                {
                    if ( location == "BN")
                    {
                        if ( field == "classa")
                        {
                        }
                    }
                    e.TotalValue = dRows[0][field].ObjToDouble();
                    return;
                }
                else if (e.SummaryProcess == CustomSummaryProcess.Finalize)
                {
                    if ( field == "cashAdvance")
                    {
                    }
                    dTotal = 0D;
                    for ( int i=0; i<groupDt.Rows.Count; i++)
                    {
                        dValue = groupDt.Rows[i][field].ObjToDouble();
                        dTotal += dValue;

                    }

                    e.TotalValue = dTotal;

                    return;
                }
            }
            catch ( Exception ex)
            {
            }
        }
        /*******************************************************************************************/
        private DataTable funeralsDt = null;
        private void LoadFuneralLocations(DataTable dt)
        {
            if (funeralsDt == null)
                funeralsDt = G1.get_db_data("Select * from `funeralhomes`;");
            string contract = "";
            string contractNumber = "";
            string trust = "";
            string loc = "";
            DataRow[] dR = null;
            if (G1.get_column_number(dt, "loc") < 0)
                dt.Columns.Add("loc");
            if (G1.get_column_number(dt, "location") < 0)
                dt.Columns.Add("location");

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                    if (contractNumber == "L17035UI")
                    {
                    }
                    contract = dt.Rows[i]["serviceId"].ObjToString();
                    contract = Trust85.decodeContractNumber(contract, ref trust, ref loc);

                    if (String.IsNullOrWhiteSpace(loc))
                        continue;
                    //dR = funeralsDt.Select("keycode='" + loc + "'");
                    dR = funeralsDt.Select("atneedcode='" + loc + "'");
                    if (dR.Length > 0)
                    {
                        dt.Rows[i]["loc"] = loc;
                        dt.Rows[i]["location"] = dR[0]["LocationCode"].ObjToString();
                    }
                    else
                    {
                        dR = funeralsDt.Select("keycode='" + loc + "'");
                        if (dR.Length > 0)
                        {
                            dt.Rows[i]["loc"] = dR[0]["atneedcode"].ObjToString();
                            dt.Rows[i]["location"] = dR[0]["LocationCode"].ObjToString();
                        }
                        else
                        {
                            dt.Rows[i]["loc"] = loc;
                            dt.Rows[i]["location"] = loc;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }
        }
        /***********************************************************************************************/
        private void ForceGroups()
        {
            if (dgv.DataSource == null)
                return;
            DataTable dt = (DataTable)dgv.DataSource;
            DataView tempview = dt.DefaultView;
            tempview.Sort = "location, lastName, firstName";
            dt = tempview.ToTable();
            G1.NumberDataTable(dt);
            dgv.DataSource = dt;

            gridMain.Columns["location"].GroupIndex = 0;
            this.gridMain.ExpandAllGroups();
            SetupTotalsSummary();
            gridMain.RefreshData();
            dgv.Refresh();
            gridMain.OptionsView.ShowFooter = showFooters;
        }
        /***********************************************************************************************/
        private void chkSort_CheckedChanged(object sender, EventArgs e)
        {
            if (dgv.DataSource == null)
                return;
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
                return;
            try
            {
                if (byPass)
                    return;
                if (chkArranger.Checked || chkManager.Checked || chkDirector.Checked )
                {
                    byPass = true;
                    chkArranger.Checked = false;
                    chkManager.Checked = false;
                    chkDirector.Checked = false;
                    gridMain.Columns["Funeral Arranger"].GroupIndex = -1;
                    gridMain.Columns["Funeral Director"].GroupIndex = -1;
                    gridMain.Columns["manager"].GroupIndex = -1;
                    byPass = false;
                }
                if (chkSort.Checked)
                {
                    DataView tempview = dt.DefaultView;
                    tempview.Sort = "serviceLoc, lastName, firstName";
                    dt = tempview.ToTable();
                    G1.NumberDataTable(dt);
                    dgv.DataSource = dt;

                    gridMain.Columns["serviceLoc"].GroupIndex = 0;
                    //if (!showFooters)
                    //{
                    //    gridMain.OptionsView.ShowFooter = false;
                    //    gridMain.Appearance.FooterPanel.Dispose();
                    //    gridMain.Appearance.GroupFooter.Dispose();
                    //    gridMain.Columns["balanceDue"].SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.None;
                    //    gridMain.Columns["balanceDue"].SummaryItem.DisplayFormat = "";
                    //    gridMain.Columns["balanceDue"].Summary.Clear();

                    //}

                    gridMain.OptionsView.ShowFooter = true;
                    //gridMain.OptionsView.ShowFooter = false;

                    this.gridMain.ExpandAllGroups();
                    SetupTotalsSummary();
                }
                else
                {
                    gridMain.Columns["location"].GroupIndex = -1;
                    gridMain.Columns["serviceLoc"].GroupIndex = -1;
                    gridMain.CollapseAllGroups();
                }
                allowSummary = false;
                gridMain.RefreshData();
                gridMain.RefreshEditor(true);
            }
            catch ( Exception ex)
            {
            }
            dgv.Refresh();
        }
        /***********************************************************************************************/
        private void chkComboLocation_EditValueChanged(object sender, EventArgs e)
        {
            if (originalDt == null)
                return;
            string names = getLocationQuery();
            DataRow[] dRows = originalDt.Select(names);
            DataTable dt = originalDt.Clone();
            for (int i = 0; i < dRows.Length; i++)
                dt.ImportRow(dRows[i]);
            G1.NumberDataTable(dt);
            dgv.DataSource = dt;
            this.gridMain.ExpandAllGroups();
            gridMain.OptionsView.ShowFooter = showFooters;
        }
        /***********************************************************************************************/
        private void latestChangesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            FuneralsChanges fForm = new FuneralsChanges();
            fForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void recalculateBalanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);
            string contractNumber = dr["contractNumber"].ObjToString();

            string cmd = "Select * from `fcust_extended` where `contractNumber` = '" + contractNumber + "';";
            DataTable dt = G1.get_db_data(cmd);
            if (dt.Rows.Count <= 0)
                return;
            string record = dt.Rows[0]["record"].ObjToString();
            string number = dr["num"].ObjToString();
            string serviceId = dr["serviceId"].ObjToString();
            string srvloc = dr["SRVLOC"].ObjToString();
            string serviceLoc = dr["serviceLoc"].ObjToString();

            Funerals.CalculateCustomerDetails(contractNumber, record, dr, true );

            this.Cursor = Cursors.WaitCursor;
            dt = (DataTable)dgv.DataSource;

            ProcessRow( dt, ref dr);

            dr["num"] = number;
            dr["serviceId"] = serviceId;
            dr["SRVLOC"] = srvloc;
            dr["serviceLoc"] = serviceLoc;

            CompareTimJones();

            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void recalculateBalance(string contractNumber, DataRow dr)
        {
            //DataRow dr = gridMain.GetFocusedDataRow();
            //int rowHandle = gridMain.FocusedRowHandle;
            //int row = gridMain.GetDataSourceRowIndex(rowHandle);
            //string contractNumber = dr["contractNumber"].ObjToString();

            string cmd = "Select * from `fcust_extended` where `contractNumber` = '" + contractNumber + "';";
            DataTable dt = G1.get_db_data(cmd);
            if (dt.Rows.Count <= 0)
                return;
            string record = dt.Rows[0]["record"].ObjToString();
            CalculateCustomerDetails(contractNumber, record, dr);
        }
        /***********************************************************************************************/
        private void CalculateCustomerDetails(string contractNumber, string custExtendedRecord, DataRow dR)
        {
            if (1 == 1)
                return;
            //this.Cursor = Cursors.WaitCursor;

            //string cmd = "Select * from `cust_payments` where `contractNumber` = '" + contractNumber + "';";
            //DataTable dt = G1.get_db_data(cmd);

            //double oldDiscount = dR["totalDiscount"].ObjToDouble();

            //string amountFiled = "";
            //string amountReceived = "";
            //string amountDiscount = "";
            //string amountGrowth = "";
            //string grossAmountReceived = "";
            //double totalFiled = 0D;
            //double totalReceived = 0D;
            //double totalDiscount = 0D;
            //double totalAmountGrowth = 0D;
            //double totalAmountDiscount = 0D;
            //double totalGross = 0D;
            //double payment = 0D;
            //double totalPayments = 0D;

            //double trustAmountFiled = 0D;
            //double insAmountFiled = 0D;
            //double trustAmountReceived = 0D;
            //double insAmountReceived = 0D;

            //string str = "";
            //string type = "";
            //string cash = "";
            //string status = "";
            //string deposit = "";
            //string creditCard = "";
            //string ccDepNumber = "";
            //double dValue = 0D;
            //double balanceDue = 0D;
            //double discount = 0D;
            //double classa = 0D;
            //string approvedBy = "";
            //DateTime dateEntered = DateTime.Now;
            //DateTime dateModified = DateTime.Now;
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    try
            //    {
            //        status = dt.Rows[i]["status"].ObjToString().ToUpper();
            //        if (status.ToUpper() == "CANCELLED")
            //            continue;
            //        amountFiled = dt.Rows[i]["amountFiled"].ObjToString();
            //        amountReceived = dt.Rows[i]["amountReceived"].ObjToString();
            //        amountDiscount = dt.Rows[i]["amountDiscount"].ObjToString();
            //        amountGrowth = dt.Rows[i]["amountGrowth"].ObjToString();
            //        grossAmountReceived = dt.Rows[i]["grossAmountReceived"].ObjToString();
            //        payment = dt.Rows[i]["payment"].ObjToDouble();
            //        totalFiled += amountFiled.ObjToDouble();
            //        totalReceived += amountReceived.ObjToDouble();
            //        totalAmountDiscount += amountDiscount.ObjToDouble();
            //        totalAmountGrowth += amountGrowth.ObjToDouble();
            //        totalGross += grossAmountReceived.ObjToDouble();

            //        type = dt.Rows[i]["type"].ObjToString().ToUpper();
            //        if (type.ToUpper() == "DISCOUNT" || status.ToUpper() == "DEPOSITED")
            //            totalPayments += payment;
            //        if (type == "CASH")
            //        {
            //            dValue = dt.Rows[i]["payment"].ObjToDouble();
            //            str = G1.ReformatMoney(dValue);
            //            cash += "CA - " + str + " ";
            //        }
            //        else if (type == "CREDIT CARD")
            //        {
            //            dValue = dt.Rows[i]["payment"].ObjToDouble();
            //            str = G1.ReformatMoney(dValue);
            //            creditCard += "CC - " + str + " ";
            //        }
            //        else if (type == "CLASSA")
            //        {
            //            classa += dt.Rows[i]["payment"].ObjToDouble();
            //        }
            //        str = dt.Rows[i]["depositNumber"].ObjToString();
            //        if (!String.IsNullOrWhiteSpace(str))
            //        {
            //            if (type == "CASH")
            //                deposit += str + " ";
            //            else if (type == "CREDIT CARD")
            //                ccDepNumber += str + " ";
            //        }
            //        if (type == "DISCOUNT")
            //        {
            //            discount += dt.Rows[i]["payment"].ObjToDouble();
            //            str = dt.Rows[i]["approvedBy"].ObjToString();
            //            if (!String.IsNullOrWhiteSpace(str))
            //                approvedBy += str + " ";
            //        }
            //        if (type == "TRUST")
            //        {
            //            if (status == "PENDING")
            //                trustAmountFiled += amountFiled.ObjToDouble();
            //            else if (status == "DEPOSITED")
            //                trustAmountReceived += amountReceived.ObjToDouble();
            //        }
            //        else if (type == "INSURANCE")
            //        {
            //            if (status == "PENDING")
            //                insAmountFiled += amountFiled.ObjToDouble();
            //            else if (status == "DEPOSITED")
            //                insAmountReceived += amountReceived.ObjToDouble();
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //    }
            //}

            //string serviceId = dR["ServiceId"].ObjToString();
            //EditCust.DetermineActiveGroups(contractNumber, serviceId);

            //string myActiveFuneralHomeGroup = EditCust.activeFuneralHomeGroup;
            //string myActiveFuneralHomeCasketGroup = EditCust.activeFuneralHomeCasketGroup;

            //EditCustomer.activeFuneralHomeGroup = myActiveFuneralHomeGroup;
            //EditCustomer.activeFuneralHomeCasketGroup = myActiveFuneralHomeCasketGroup;


            //FunServices funForm = new FunServices(contractNumber);
            //DataTable funDt = funForm.funServicesDT;
            //double price = 0D;
            //double totalServices = 0D;
            //double totalMerchandise = 0D;
            //double totalCashAdvance = 0D;
            //double totalCost = 0D;
            //double difference = 0D;
            //double totalDifference = 0D;
            //double currentPrice = 0D;
            //double totalCurrentPrice = 0D;

            //string service = "";

            //double currentServices = 0D;
            //double currentMerchandise = 0D;

            //double merchandiseDiscount = 0D;
            //double servicesDiscount = 0D;

            //double totalPackagePrice = 0D;
            //double packagePrice = 0D;
            //double packageDiscount = 0D;
            //bool gotPackage = false;

            //if (funDt != null)
            //{
            //    for (int i = 0; i < funDt.Rows.Count; i++)
            //    {
            //        price = funDt.Rows[i]["price"].ObjToDouble();
            //        currentPrice = funDt.Rows[i]["currentPrice"].ObjToDouble();
            //        difference = currentPrice - price;

            //        service = funDt.Rows[i]["service"].ObjToString().ToUpper();
            //        if (service == "TOTAL LISTED PRICE")
            //        {
            //            totalPackagePrice = funDt.Rows[i]["price"].ObjToDouble();
            //            continue;
            //        }
            //        else if (service == "PACKAGE PRICE")
            //        {
            //            gotPackage = true;
            //            packagePrice = funDt.Rows[i]["price"].ObjToDouble();
            //            continue;
            //        }
            //        else if (service == "PACKAGE DISCOUNT")
            //        {
            //            packageDiscount = funDt.Rows[i]["price"].ObjToDouble();
            //            packageDiscount = Math.Abs(packageDiscount);
            //            continue;
            //        }

            //        type = funDt.Rows[i]["type"].ObjToString().ToUpper();
            //        if (type == "SERVICE")
            //        {
            //            totalServices += price;
            //            servicesDiscount += difference;
            //            currentServices += currentPrice;
            //        }
            //        else if (type == "MERCHANDISE")
            //        {
            //            totalMerchandise += price;
            //            merchandiseDiscount += difference;
            //            currentMerchandise += currentPrice;
            //        }
            //        else if (type == "CASH ADVANCE")
            //            totalCashAdvance += price;
            //        totalCurrentPrice += currentPrice;
            //        totalDifference += (currentPrice - price);
            //    }
            //}

            //totalCost = totalCurrentPrice - totalDifference;
            //if (gotPackage)
            //    totalCost = packagePrice;
            //balanceDue = totalCost - totalPayments;
            //totalDiscount = servicesDiscount + merchandiseDiscount;
            //if (gotPackage)
            //    totalDiscount = packageDiscount;
            //if (oldDiscount != totalDiscount)
            //{
            //}

            //if (!String.IsNullOrWhiteSpace(custExtendedRecord))
            //{
            //    G1.update_db_table("fcust_extended", "record", custExtendedRecord, new string[] { "amountFiled", totalFiled.ToString(), "amountReceived", totalReceived.ToString(), "cash", cash, "depositNumber", deposit, "balanceDue", balanceDue.ToString(), "additionalDiscount", discount.ToString(), "approvedBy", approvedBy, "creditCard", creditCard, "ccDepNumber", ccDepNumber, "grossAmountReceived", totalGross.ObjToString(), "classa", classa.ToString(), "amountDiscount", totalAmountDiscount.ObjToString(), "amountGrowth", totalAmountGrowth.ObjToString() });
            //    G1.update_db_table("fcust_extended", "record", custExtendedRecord, new string[] { "custPrice", totalCost.ToString(), "custMerchandise", totalMerchandise.ToString(), "custServices", totalServices.ToString(), "merchandiseDiscount", merchandiseDiscount.ToString(), "servicesDiscount", servicesDiscount.ToString(), "totalDiscount", totalDiscount.ToString(), "currentPrice", totalCurrentPrice.ToString(), "currentMerchandise", currentMerchandise.ToString(), "currentServices", currentServices.ToString() });
            //    G1.update_db_table("fcust_extended", "record", custExtendedRecord, new string[] { "trustAmountFiled", trustAmountFiled.ObjToString(), "trustAmountReceived", trustAmountReceived.ObjToString(), "insAmountFiled", insAmountFiled.ObjToString(), "insAmountReceived", insAmountReceived.ObjToString() });
            //}
            //try
            //{
            //    dR["balanceDue"] = balanceDue;
            //    dR["custPrice"] = totalCost;
            //    dR["custMerchandise"] = totalMerchandise;
            //    dR["custServices"] = totalServices;
            //    dR["merchandiseDiscount"] = merchandiseDiscount;
            //    dR["servicesDiscount"] = servicesDiscount;
            //    dR["additionalDiscount"] = discount;
            //    dR["totalDiscount"] = totalDiscount;

            //    dR["amountFiled"] = totalFiled;
            //    dR["amountReceived"] = totalReceived;

            //    dR["currentPrice"] = totalCurrentPrice;
            //    dR["currentMerchandise"] = currentMerchandise;
            //    dR["currentServices"] = currentServices;

            //    dR["grossAmountReceived"] = totalGross;

            //    dR["amountDiscount"] = totalAmountDiscount;
            //    dR["amountGrowth"] = totalAmountGrowth;
            //}
            //catch (Exception ex)
            //{
            //}

            //this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripRemoveFormat_Click(object sender, EventArgs e)
        {
            string comboName = cmbSelectColumns.Text;
            if (!String.IsNullOrWhiteSpace(comboName))
            {
                string name = "AR " + workReport + " " + comboName;
                G1.RemoveLocalPreferences(LoginForm.username, name);
                foundLocalPreference = false;
            }
        }
        /***********************************************************************************************/
        private void btnSelectPosition_Click(object sender, EventArgs e)
        {
        }
        /***********************************************************************************************/
        private void gridMain_FilterEditorCreated(object sender, FilterControlEventArgs e)
        {
            e.FilterControl.BeforeShowValueEditor += new DevExpress.XtraEditors.Filtering.ShowValueEditorEventHandler(FilterControl_BeforeShowValueEditor);
        }
        /***********************************************************************************************/
        void FilterControl_BeforeShowValueEditor(object sender, DevExpress.XtraEditors.Filtering.ShowValueEditorEventArgs e)
        {
            if ( e.PropertyName.ToUpper() == "TMSTAMP")
                e.CustomRepositoryItem = ri;
        }
        /***********************************************************************************************/
        private void casketUsageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CasketUsageReport usageForm = new CasketUsageReport();
            usageForm.Show();
        }
        /***********************************************************************************************/
        private void repPaidCheckEdit_Click(object sender, EventArgs e)
        {
            if (loading)
                return;

            DataTable dt = (DataTable)(dgv.DataSource);
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            string value = dr["paidInFull1"].ObjToString();
            string record = dr["record5"].ObjToString();
            int row = gridMain.GetDataSourceRowIndex(gridMain.FocusedRowHandle);
            if (value == "0")
                value = "1";
            else
                value = "0";
            dr["paidInFull1"] = value;
            dt.Rows[row]["paidInFull1"] = value;
            dgv.DataSource = dt;
            if (workReport.ToUpper() == "CASH BALANCE REPORT")
            {
                string records = dr["records"].ObjToString();
                if (!String.IsNullOrWhiteSpace(records))
                {
                    string[] Lines = records.Split(',');
                    for (int i = 0; i < Lines.Length; i++)
                    {
                        record = Lines[i].Trim();
                        if (String.IsNullOrWhiteSpace(record))
                            continue;
                        G1.update_db_table("cust_payments", "record", record, new string[] { "paidInFull", value });
                    }
                }
                //return;
            }
            else
                G1.update_db_table("cust_payments", "record", record, new string[] { "paidInFull", value });
            gridMain.PostEditor();
            gridMain.RefreshData();
            gridMain.RefreshEditor(true);
            dgv.Refresh();
        }
        /***********************************************************************************************/
        private void repPaidCheckEdit_CheckStateChanged(object sender, EventArgs e)
        {
            if (loading)
                return;

            DataTable dt = (DataTable)(dgv.DataSource);
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            string value = dr["paidInFull1"].ObjToString();
            string record = dr["record5"].ObjToString();
            int row = gridMain.GetDataSourceRowIndex(gridMain.FocusedRowHandle);
            if (value == "0")
                value = "1";
            else
                value = "0";
            dr["paidInFull1"] = value;
            dt.Rows[row]["paidInFull1"] = value;
            dgv.DataSource = dt;
            //G1.update_db_table("cust_payments", "record", record, new string[] { "paidInFull", value });
            gridMain.PostEditor();
        }
        /***********************************************************************************************/
        private void trustsClaimsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
        }
        /***********************************************************************************************/
        private void gridMain_CalcRowHeightx(object sender, RowHeightEventArgs e)
        {
            GridView View = sender as GridView;
            if (e.RowHandle >= 0)
            {
                string cancelled = View.GetRowCellDisplayText(e.RowHandle, View.Columns["notes"]);
                if (!String.IsNullOrWhiteSpace(cancelled))
                {
                   // DataGridViewCell cell = View[e.RowHandle, "notes"];

                    string text = cancelled;
                    Size textSize = TextRenderer.MeasureText(text, View.Columns["notes"].AppearanceCell.Font);
                    int charWidth = textSize.Width / text.Length;
                    int charCount = View.Columns["notes"].Width / charWidth;

                    int originalRowHeight = e.RowHeight;
                    cancelled = cancelled.TrimEnd('\n');
                    string[] Lines = cancelled.Split('\n');
                    int count = Lines.Length;
                    count = charWidth;
                    if (count > 1)
                        e.RowHeight = originalRowHeight * count;
                }
            }
        }
        /***********************************************************************************************/
        private void gridMain_CalcRowHeight(object sender, RowHeightEventArgs e)
        {
            if (1 == 1)
                return;
            GridView View = sender as GridView;
            if (e.RowHandle >= 0)
            {
                int maxHeight = 0;
                int newHeight = 0;
                string name = "";
                bool doit = false;
                foreach (GridColumn column in gridMain.Columns)
                {
                    doit = true;
                    if (workReport.ToUpper() == "CASH BALANCE REPORT")
                    {
                        doit = false;
                        name = column.FieldName.ToUpper();
                        if (name == "NAMES" || name == "CLASS_A" || name == "DISCOUNTTYPE" || name == "DATEINSFILED" || name == "DEPOSITNUMBER" || name == "DEPOSITDATE" || name == "CCDEPOSITNUMBER" || name == "CCDEPOSITDATE" || name == "NOTES" )
                            doit = true;
                        if (name == "DEPOSITDATE")
                        {
                        }
                    }
                    if (doit)
                    {
                        using (RepositoryItemMemoEdit edit = new RepositoryItemMemoEdit())
                        {
                            using (MemoEditViewInfo viewInfo = edit.CreateViewInfo() as MemoEditViewInfo)
                            {
                                viewInfo.EditValue = gridMain.GetRowCellValue(e.RowHandle, column.FieldName);
                                viewInfo.Bounds = new Rectangle(0, 0, column.VisibleWidth, dgv.Height);
                                using (Graphics graphics = dgv.CreateGraphics())
                                using (GraphicsCache cache = new GraphicsCache(graphics))
                                {
                                    viewInfo.CalcViewInfo(graphics);
                                    var height = ((IHeightAdaptable)viewInfo).CalcHeight(cache, column.VisibleWidth);
                                    newHeight = Math.Max(height, maxHeight);
                                    if (newHeight > maxHeight)
                                        maxHeight = newHeight;
                                }
                            }
                        }
                    }
                }

                if ( maxHeight > 0 )
                    e.RowHeight = maxHeight;
            }
        }
        /***********************************************************************************************/
        private void gridMain_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            ColumnView view = sender as ColumnView;

            string columnName = e.Column.FieldName.ToUpper();
            if (columnName == "DATEINSFILED")
                return;
            if (columnName == "DEPOSITDATE")
                return;
            if (columnName == "CCDEPOSITDATE")
                return;

            if (e.Column.FieldName.ToUpper().IndexOf("CASECREATEDDATE") >= 0 && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                string str = e.DisplayText.Trim();
                if (String.IsNullOrWhiteSpace(str))
                {
                    DataTable dt = (DataTable)dgv.DataSource;
                    int row = e.ListSourceRowIndex;
                    DateTime date = dt.Rows[row]["tmstamp"].ObjToDateTime();
                    e.DisplayText = date.ToString("yyyy-MM-dd");
                }
                else
                {
                    DateTime date = e.DisplayText.ObjToDateTime();
                    if (date.Year < 100)
                        e.DisplayText = "";
                    else
                        e.DisplayText = date.ToString("yyyy-MM-dd");
                }
            }
            else if (e.Column.FieldName.ToUpper().IndexOf("DATE") >= 0 && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                if (e.DisplayText.IndexOf("0000") >= 0 || e.DisplayText.IndexOf("0001") >= 0)
                    e.DisplayText = "";
                else
                {
                    DateTime date = e.DisplayText.ObjToString().ObjToDateTime();
                    if (date.Year > 100)
                        e.DisplayText = date.ToString("MM/dd/yyyy");
                    else
                        e.DisplayText = "";
                }
            }
            else if (e.Column.FieldName.ToUpper().IndexOf("TMSTAMP") >= 0 && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                DateTime date = e.DisplayText.ObjToString().ObjToDateTime();
                e.DisplayText = date.ToString("yyyy-MM-dd HH:mm:ss");
            }
            else if (e.Column.FieldName.ToUpper().IndexOf("CONTRACTTYPE") >= 0 && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                string data = e.DisplayText;
                string[] Lines = data.Split('(');
                if (Lines.Length > 1)
                {
                    data = Lines[1];
                    data = data.Replace("(", "");
                    data = data.Replace(")", "");
                    e.DisplayText = data;
                }
            }
            else if (e.Column.FieldName.ToUpper() == "FUNERAL ARRANGER" && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                string str = e.DisplayText.Trim();
                if (!String.IsNullOrWhiteSpace(str))
                {
                    string[] Lines = str.Split('[');
                    if (Lines.Length > 0)
                    {
                        str = Lines[0].Trim();
                        e.DisplayText = str;
                    }
                }
            }
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
        private double originalSize = 0D;
        private Font mainFont = null;
        private Font newFont = null;
        private void ScaleCells()
        {
            if (originalSize == 0D)
            {
                //                originalSize = gridMain.Columns["address1"].AppearanceCell.FontSizeDelta.ObjToDouble();
                originalSize = gridMain.Columns["name"].AppearanceCell.Font.Size;
                mainFont = gridMain.Columns["name"].AppearanceCell.Font;
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
        /***********************************************************************************************/
        private void gridMain_CustomDrawFooterCell(object sender, FooterCellCustomDrawEventArgs e)
        {
            if (newFont != null)
                e.Appearance.Font = newFont;
        }
        /***********************************************************************************************/
        private void btnLeft_Click(object sender, EventArgs e)
        {
            DateTime now = this.dateTimePicker2.Value;
            DateTime date = new DateTime(now.Year, now.Month, 1);
            date = date.AddMonths(-1);
            int days = DateTime.DaysInMonth(date.Year, date.Month);
            date = new DateTime(date.Year, date.Month, days);
            this.dateTimePicker2.Value = date;
            this.dateTimePicker1.Value = new DateTime(date.Year, date.Month, 1 );
        }
        /***********************************************************************************************/
        private void btnRight_Click(object sender, EventArgs e)
        {
            DateTime now = this.dateTimePicker2.Value;
            DateTime date = new DateTime(now.Year, now.Month, 1);
            date = date.AddMonths(1);
            int days = DateTime.DaysInMonth(date.Year, date.Month);
            date = new DateTime(date.Year, date.Month, days);
            this.dateTimePicker2.Value = date;
            this.dateTimePicker1.Value = new DateTime(date.Year, date.Month, 1);
        }
        /***********************************************************************************************/
        private void btnRunCommission_Click(object sender, EventArgs e)
        {
            historicBonus = false;
            if ( groupDt != null )
                groupDt.Rows.Clear();

            if (workDt != null)
            {
                btnSaveBonus.Hide();
                btnSaveBonus.Refresh();

                btnCompare.Hide();
                btnCompare.Refresh();
                btnZeros.Hide();
                btnZeros.Refresh();


                FormWindowState state = this.WindowState;

                this.Cursor = Cursors.WaitCursor;

                runBatch();

                this.Cursor = Cursors.WaitCursor;
                bool savePrint = chkPreparePrint.Checked;

                DataTable dx = ShowDetail(null, null);

                this.Cursor = Cursors.WaitCursor;
                if (savePrint)
                    chkPreparePrint_CheckedChanged(null, null);

                DataView tempview = dx.DefaultView;
                tempview.Sort = "SRVLOC,exclude,serviceId";
                dx = tempview.ToTable();

                if ( G1.get_column_number(dx, "XserviceId") < 0)
                    dx.Columns.Add("XserviceId");
                for ( int i=0; i<dx.Rows.Count; i++)
                {
                    dx.Rows[i]["XserviceId"] = dx.Rows[i]["exclude"].ObjToString() + dx.Rows[i]["serviceId"].ObjToString();
                    dx.Rows[i]["serviceId"] = dx.Rows[i]["exclude"].ObjToString() + dx.Rows[i]["serviceId"].ObjToString();
                }

                //allowSummary = true;

                dgv.DataSource = dx;

                allowSummary = false;

                //calcSummary = true;

                //gridMain.UpdateGroupSummary();

                //calcSummary = false;


                this.Text = "Funerals for Manager " + workManager;

                string title = "Funeral Details for ";
                if (workWho.ToUpper() == "M")
                    title += "All Managers";
                else if (workWho.ToUpper() == "A")
                {
                    if (!String.IsNullOrWhiteSpace(workArranger))
                        title += "All Arrangers";
                    else
                        title += "All Arrangers";
                }
                this.Text = title;

                this.Cursor = Cursors.WaitCursor;
                chkSort.Checked = true;
                //chkSort_CheckedChanged(null, null);

                gridMain.Columns["serviceId"].SortOrder = ColumnSortOrder.Ascending;
                this.Cursor = Cursors.WaitCursor;
                gridMain.ExpandAllGroups();

                this.Cursor = Cursors.WaitCursor;
                gridMain.RefreshData();
                gridMain.RefreshEditor(true);

                this.Cursor = Cursors.WaitCursor;
                allowSummary = false;
                gridMain.UpdateSummary();
                allowSummary = false;

                this.Cursor = Cursors.WaitCursor;
                CheckSaveButton();

                menuStrip1.BackColor = menuBackColor;

                if (this.WindowState != state)
                {
                    this.WindowState = state;
                    this.Visible = true;
                    this.BringToFront();
                }

                btnCompare.Show();
                btnCompare.Refresh();
                btnZeros.Show();
                btnZeros.Refresh();

                this.Cursor = Cursors.Default;
                return;
            }
            DataTable dt = (DataTable)dgv.DataSource;
            DataRow dr = gridMain.GetFocusedDataRow();

            string who = dr["funeralType"].ObjToString();

            //string cmd = "Select * from `funcommissiondata` where `name` = '" + workManager + "' OR `name` = '" + workArranger + "';";
            //DataTable dx = G1.get_db_data(cmd);

            this.Cursor = Cursors.WaitCursor;
            FunCommissions funForm = null;
            if (String.IsNullOrWhiteSpace(workManager) && String.IsNullOrWhiteSpace(workArranger))
            {
                funForm = new FunCommissions(dt, workManager, workLocation, workArranger, workWho, false);
                funForm.Show();
            }
            else
            {
                funForm = new FunCommissions(dt, workManager, workLocation, workArranger, workWho, false);
                funForm.funCommissionClosing += FunForm_funCommissionClosing;
                funForm.ShowDialog();
            }

            menuStrip1.BackColor = menuBackColor;

            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void ShowZeros ( DataTable dt )
        {
            DataTable dx = dt.Clone();
            //DataRow[] dRows = dt.Select("casketCost='0.00' AND exclude<>'Y'");
            if ( dt.Rows.Count > 0 )
            {
                string urnDesc = "";
                double casketCost = 0D;
                double urnCost = 0D;
                double casketAmount = 0D;
                double urnAmount = 0D;
                string casketCode = "";
                string casketDesc = "";
                string funeral_class = "";
                string fpc = "";

                for ( int i=0; i<dt.Rows.Count; i++)
                {
                    funeral_class = dt.Rows[i]["funeral_classification"].ObjToString();
                    if (funeral_class.ToUpper().IndexOf("SHIP") >= 0)
                        continue;
                    else if (funeral_class.ToUpper().IndexOf("TENT AND CHAIR") >= 0)
                        continue;
                    else if (funeral_class.ToUpper().IndexOf("PICKUP") >= 0)
                        continue;
                    fpc = dt.Rows[i]["fpc"].ObjToString().ToUpper();
                    if (fpc == "Y")
                        continue;
                    casketDesc = dt.Rows[i]["casketdesc"].ObjToString().ToUpper();
                    casketCode = dt.Rows[i]["casket"].ObjToString().ToUpper();
                    casketCost = dt.Rows[i]["casketCost"].ObjToDouble();
                    urnCost = dt.Rows[i]["urnCost"].ObjToDouble();
                    if (casketCode == "ALTC")
                        continue;
                    else if (casketCode == "INFANT")
                        continue;
                    else if (casketDesc.IndexOf("CREMATION") >= 0)
                        continue;
                    else if (casketDesc.IndexOf("RENTAL CASKET") >= 0)
                        continue;
                    else if ( casketCost == 0D )
                    {
                        if (urnCost != 0D)
                            continue;
                    }
                    else if ( urnCost == 0D )
                    {
                        if (casketCost != 0D)
                            continue;
                    }
                    dx.ImportRow(dt.Rows[i]);
                }

                if (dx.Rows.Count > 0)
                {
                    ShowZeros zeroForm = new ShowZeros(dx);
                    zeroForm.editDone += ZeroForm_editDone;
                    zeroForm.Show();
                }
            }
        }
        /***********************************************************************************************/
        private void ZeroForm_editDone(DataTable dt)
        {
            string serviceId = "";
            double casketCost = 0D;
            string casketDesc = "";
            string oldDesc = "";
            string record = "";
            string mod = "";
            DataRow[] dRows = null;
            DataTable dx = (DataTable)dgv.DataSource;

            for ( int i=0; i<dt.Rows.Count; i++)
            {
                mod = dt.Rows[i]["mod"].ObjToString();
                if (mod != "Y")
                    continue;
                serviceId = dt.Rows[i]["serviceId"].ObjToString();
                casketCost = dt.Rows[i]["casketCost"].ObjToDouble();
                casketDesc = dt.Rows[i]["casketdesc"].ObjToString();
                if ( !String.IsNullOrWhiteSpace ( serviceId ))
                {
                    dRows = dx.Select("serviceId='" + serviceId + "'");
                    if ( dRows.Length > 0 )
                    {
                        record = dRows[0]["record"].ObjToString();
                        dRows[0]["casketCost"] = casketCost;
                        oldDesc = dRows[0]["casketdesc"].ObjToString();
                        if (String.IsNullOrWhiteSpace(oldDesc))
                            dRows[0]["casketdesc"] = casketDesc;
                        else
                            casketDesc = oldDesc;

                        G1.update_db_table("fcust_extended", "record", record, new string[] {"casketCost", casketCost.ToString(), "casketdesc", casketDesc });
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void CheckSaveButton ()
        {
            DateTime date = this.dateTimePicker2.Value;
            DateTime startDate = new DateTime(date.Year, date.Month, 1);
            date = this.dateTimePicker1.Value;
            if ( date == startDate )
            {
                btnSaveBonus.Show();
                btnSaveBonus.Refresh();
            }
        }
        /***********************************************************************************************/
        private void runBatch ()
        {
            workPDF = true;

            string prefix = "";
            string suffix = "";
            string firstName = "";
            string lastName = "";
            string mi = "";

            DataTable dt = null;
            myDataTable = null;
            workCommission = 0D;
            workFirst = true;
            int lastRow = workDt.Rows.Count;

            barImport.Show();
            barImport.Minimum = 0;
            barImport.Maximum = lastRow;
            barImport.Refresh();

            //iTextSharp.text.Document sourceDocument = null;
            //PdfCopy pdfCopyProvider = null;
            //PdfImportedPage importedPage;
            //string outputPdfPath = @"C:/rag/pdfX.pdf";

            //if (File.Exists(outputPdfPath))
            //{
            //    File.SetAttributes(outputPdfPath, FileAttributes.Normal);
            //    File.Delete(outputPdfPath);
            //}

            //sourceDocument = new iTextSharp.text.Document();
            //pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

            ////output file Open  
            //sourceDocument.Open();

            string historyFile = @"C:/rag/pdfDaily.pdf";
            string manualFile = @"c:/rag/Manual.pdf";
            FunCommissions funForm = null;

            dbCount = 0;
            bool gotLocation = true;
            if (G1.get_column_number(workDt, "location") < 0)
                gotLocation = false;

            //lastRow = 1;

            this.Cursor = Cursors.WaitCursor;

            for (int i = 0; i < lastRow; i++)
            {
                barImport.Value = i + 1;
                barImport.Refresh();

                if (workDt.Rows[i]["atneedcode"].ObjToString().ToUpper() == "MC")
                    continue;
                workManager = workDt.Rows[i]["name"].ObjToString();
                if (String.IsNullOrWhiteSpace(workManager))
                    continue;
                workWho = workDt.Rows[i]["ma"].ObjToString();
                workLocation = "";
                if ( gotLocation )
                    workLocation = workDt.Rows[i]["location"].ObjToString();

                workArrangerLastName = "";
                workArrangerFirstName = "";

                G1.ParseOutName(workManager, ref prefix, ref firstName, ref lastName, ref mi, ref suffix);

                if (!String.IsNullOrWhiteSpace(firstName))
                    workArrangerFirstName = firstName;
                if (!String.IsNullOrWhiteSpace(lastName))
                    workArrangerLastName = lastName;

                workArranger = firstName + " " + lastName;

                SetupServiceLocs();

                historyFile = @"C:/rag/pdfDaily.pdf";

                dt = LoadData();


                Alldbs[dbCount] = dt.Copy();
                dbCount++;

                //dgv.DataSource = dt;
                //gridMain.RefreshEditor(true);

                //printPreviewToolStripMenuItem_Click(null, null);

                //printPreview();

                //this.Cursor = Cursors.WaitCursor;

                //dt = (DataTable)dgv.DataSource;

                //manualFile = @"c:/rag/Manual.pdf";
                ////string payOffFile = @"c:/rag/ForcePayoff.pdf";

                //funForm = new FunCommissions(dt, workManager, workLocation, workArranger, workWho, true);
                //funForm.funCommissionClosing += FunForm_funCommissionClosing;
                //funForm.Show();

                this.Cursor = Cursors.Default;
                //DailyHistory histForm = new DailyHistory(contract, historyFile, true);

                //string manualFile = @"c:/rag/Manual.pdf";
                //string payOffFile = @"c:/rag/ForcePayoff.pdf";


                //ForcedPayoffs forceForm = new ForcedPayoffs(true, record, this.dateTimePicker1.Value, this.dateTimePicker2.Value);

                //MergeAllPDF(pdfCopyProvider, historyFile, manualFile );

                //if (File.Exists(payOffFile))
                //{
                //    File.SetAttributes(payOffFile, FileAttributes.Normal);
                //    File.Delete(payOffFile);
                //}

                //if (File.Exists(historyFile))
                //{
                //    File.SetAttributes(historyFile, FileAttributes.Normal);
                //    File.Delete(historyFile);
                //}

                //if (File.Exists(manualFile))
                //{
                //    File.SetAttributes(manualFile, FileAttributes.Normal);
                //    File.Delete(manualFile);
                //}
            }

            //funForm = new FunCommissions(myDataTable, workWho, workCommission );
            //funForm.Show();

            //MergeAllPDF(pdfCopyProvider, null, manualFile);

            barImport.Value = lastRow;
            barImport.Refresh();

            //save the output file  
            //sourceDocument.Close();

            //ViewPDF myView = new ViewPDF("TEST", outputPdfPath);
            //myView.ShowDialog();

            //if (File.Exists(outputPdfPath))
            //{
            //    File.SetAttributes(outputPdfPath, FileAttributes.Normal);
            //    File.Delete(outputPdfPath);
            //}

            barImport.Hide();

            workManager = "";
            workArranger = "";
            myDataTable = null;
            workFirst = true;
        }
        /***********************************************************************************************/
        private void CalcSummaryAverages ()
        {
            funeralHomes = G1.get_db_data("Select * from `funeralhomes`;");

            if (summaryDt == null)
            {
                summaryDt = new DataTable();

                summaryDt.Columns.Add("serviceLoc");
                summaryDt.Columns.Add("location");
                summaryDt.Columns.Add("manager");
                summaryDt.Columns.Add("totalfuneral", Type.GetType("System.Double"));
                summaryDt.Columns.Add("netfuneral", Type.GetType("System.Double"));
                summaryDt.Columns.Add("average", Type.GetType("System.Double"));
                summaryDt.Columns.Add("met", Type.GetType("System.Double"));
                summaryDt.Columns.Add("calcavg", Type.GetType("System.Double"));
                summaryDt.Columns.Add("other", Type.GetType("System.Double"));
                summaryDt.Columns.Add("detail");
            }
            else
                summaryDt.Rows.Clear();


            string cmd = "Select * from `funcommissiondata`;";
            DataTable manDt = G1.get_db_data(cmd);

            DataTable dt = null;
            double totalFunerals = 0D;
            double average = 0D;
            double met = 0D;
            double calcavg = 0D;

            double netFuneral = 0D;
            double totalNet = 0D;
            double totalOther = 0D;
            double other = 0D;
            double dValue = 0D;
            int count = 0;
            string str = "";
            double minimumFuneral = 0D;
            double funeralAverage = 0D;

            DataRow dRow = null;
            DataRow [] manRow = null;
            DataTable manDtt = null;
            DataRow[] ddRows = null;

            string location = "";
            string manager = "";
            string serviceLoc = "";
            DateTime bonusDate = DateTime.Now;
            int bonusCount = 0;
            string exclude = "";
            int idx = 0;

            for ( int i=0; i<dbCount; i++)
            {
                dt = Alldbs[i];

                netFuneral = 0D;
                totalFunerals = 0D;
                serviceLoc = "";
                location = "";
                manager = "";
                minimumFuneral = 0D;
                funeralAverage = 0D;
                totalNet = 0D;
                met = 0;
                calcavg = 0D;
                totalOther = 0D;
                count = 0;
                bonusCount = 0;
                exclude = "";

                for ( int k=0; k<dt.Rows.Count; k++)
                {
                    if (String.IsNullOrWhiteSpace(serviceLoc))
                    {
                        serviceLoc = dt.Rows[k]["serviceLoc"].ObjToString();
                        manager = dt.Rows[k]["manager"].ObjToString();
                        idx = serviceLoc.IndexOf("-");
                        if (idx > 0)
                        {
                            location = serviceLoc;
                            str = serviceLoc.Substring(0, idx);
                            ddRows = funeralHomes.Select("atneedcode='" + str + "'");
                            if (ddRows.Length > 0)
                                location = ddRows[0]["LocationCode"].ObjToString();
                            manRow = manDt.Select("name='" + manager + "' AND location='" + location + "'");
                            if ( manRow.Length <= 0 )
                            {
                                serviceLoc = dt.Rows[k]["serviceLoc"].ObjToString();
                                manager = dt.Rows[k]["manager"].ObjToString();
                                idx = serviceLoc.IndexOf("-");
                                str = serviceLoc.Substring(idx + 1);
                                ddRows = funeralHomes.Select("atneedcode='" + str + "'");
                                if (ddRows.Length > 0)
                                    location = ddRows[0]["LocationCode"].ObjToString();
                                manRow = manDt.Select("name='" + manager + "' AND location='" + location + "'");
                            }
                        }
                        else
                        {
                            if (!String.IsNullOrWhiteSpace(serviceLoc))
                            {
                                ddRows = funeralHomes.Select("atneedcode='" + serviceLoc + "'");
                                if (ddRows.Length > 0)
                                    location = ddRows[0]["LocationCode"].ObjToString();
                                manRow = manDt.Select("name='" + manager + "' AND location='" + location + "'");
                            }
                        }

                        if ( manRow.Length > 0 )
                        {
                            manDtt = manRow.CopyToDataTable();
                            ddRows = manDtt.Select("option='Minimum Funerals'");
                            if (ddRows.Length > 0)
                                minimumFuneral = ddRows[0]["answer"].ObjToDouble();
                            ddRows = manDtt.Select("option='Funeral Average'");
                            if (ddRows.Length > 0)
                                funeralAverage = ddRows[0]["answer"].ObjToDouble();
                        }
                    }

                    if (chkHonorFilter.Checked)
                    {
                        str = dt.Rows[k]["bonusDate"].ObjToString();
                        if (str != "00/00/0000")
                        {
                            bonusDate = str.ObjToDateTime();
                            if (bonusDate.Year > 1000)
                            {
                                if (bonusDate < this.dateTimePicker1.Value || bonusDate > this.dateTimePicker2.Value)
                                    continue;
                            }
                        }
                    }

                    exclude = dt.Rows[k]["exclude"].ObjToString();
                    if (exclude == "Y")
                        continue;

                    totalFunerals += dt.Rows[k]["currentPrice"].ObjToDouble();
                    other = dt.Rows[k]["otherBonuses"].ObjToDouble();
                    totalOther += other;

                    str = dt.Rows[k]["gotPackage"].ObjToString();
                    if (String.IsNullOrWhiteSpace(str))
                    {
                        str = dt.Rows[k]["urn"].ObjToString();
                        dValue = str.ObjToDouble();
                        if (dValue == 0D)
                        {
                            count++;
                            bonusCount++;
                            netFuneral = dt.Rows[k]["netFuneral"].ObjToDouble();

                            if ( netFuneral >= funeralAverage )
                            {
                                met++;
                                totalNet += netFuneral;
                            }
                        }
                    }
                }

                if ( serviceLoc == "CW")
                {
                }

                dRow = summaryDt.NewRow();

                dRow["serviceLoc"] = serviceLoc;
                dRow["location"] = location;
                dRow["manager"] = manager;
                dRow["totalFuneral"] = totalFunerals;
                dRow["netFuneral"] = totalNet;
                dRow["average"] = funeralAverage;
                calcavg = 0D;
                if ( bonusCount > 0 )
                    calcavg = totalNet / bonusCount;
                dRow["calcAvg"] = calcavg;
                dRow["met"] = met;
                dRow["other"] = totalOther;

                //string manager = dRows[0]["manager"].ObjToString();
                //string location = dRows[0]["location"].ObjToString();
                //double totalFunerals = dRows[0]["totalFuneral"].ObjToDouble();
                //double netFunerals = dRows[0]["netFuneral"].ObjToDouble();
                //double met = dRows[0]["met"].ObjToDouble();
                //double average = dRows[0]["average"].ObjToDouble();
                //double calcavg = dRows[0]["calcavg"].ObjToDouble();
                //double other = dRows[0]["other"].ObjToDouble();
                //string str = $"{caption} : {info.GroupValueText} ({view.GetChildRowCount(e.RowHandle)})";

                str = serviceLoc + " (" + bonusCount + ") F=" + location + " M=" + manager + " Met = " + met.ToString() + " Avg Needed= " + G1.ReformatMoney(funeralAverage) + " CalcAvg " + G1.ReformatMoney(calcavg) + " TotFun=" + G1.ReformatMoney(totalFunerals) + " NetFun=" + G1.ReformatMoney(totalNet) + " Other=" + G1.ReformatMoney(totalOther);

                dRow["detail"] = str;

                summaryDt.Rows.Add(dRow);
            }
        }
        /***********************************************************************************************/
        private static void MergeAllPDF(PdfCopy pdfCopyProvider, string File2, string File3)
        {
            string[] fileArray = new string[3];
            fileArray[0] = File2;
            fileArray[1] = File3;

            PdfReader reader = null;
            PdfImportedPage importedPage;


            //files list wise Loop  
            for (int f = 0; f < fileArray.Length - 1; f++)
            {
                if (fileArray[f] == null)
                    continue;

                int pages = TotalPageCount(fileArray[f]);

                reader = new PdfReader(fileArray[f]);
                //Add pages in new file  
                for (int i = 1; i <= pages; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }

                reader.Close();
            }
        }
        /***********************************************************************************************/
        private static int TotalPageCount(string file)
        {
            using (StreamReader sr = new StreamReader(System.IO.File.OpenRead(file)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());

                return matches.Count;
            }
        }
        /***********************************************************************************************/
        private DataTable myDataTable = null;
        private bool workFirst = true;
        private double workCommission = 0D;
        private void FunForm_funCommissionClosing(DataTable dt, string manager, string arranger, string who, double commission )
        {
            if (myDataTable == null)
                myDataTable = dt.Clone();

            DataRow dR = null;
            dR = myDataTable.NewRow();
            if (who.ToUpper() == "M")
                dR["name"] = manager;
            else
                dR["name"] = arranger;
            dR["commission"] = commission;

            myDataTable.Rows.Add(dR);

            workCommission += commission;
        }
        /***********************************************************************************************/
        private void chkSingle_CheckedChanged(object sender, EventArgs e)
        {
            SetAllClaims();
        }
        /***********************************************************************************************/
        private DataTable ShowDetail (object sender, EventArgs e)
        {
            //DataTable dt = (DataTable)dgv.DataSource;
            if (dbCount <= 0)
                return null;

            DataTable dt = Alldbs[0].Clone();

            dt.Rows.Clear();

            DataTable dx = null;

            for ( int i=0; i<dbCount; i++)
            {
                dt.Merge(Alldbs[i]);
            }

            //G1.NumberDataTable(dt);
            //dgv.DataSource = dt;

            dt = G1.RemoveDuplicates(dt, "serviceId");

            originalDt = dt.Copy();

            allowSummary = false;

            LoadDBs(dt);

            CalcSummaryAverages();

            gridMain.RefreshData();
            gridMain.RefreshEditor(true);
            dgv.Refresh();

            return dt;
        }
        /***********************************************************************************************/
        private void LoadDBs ( DataTable dt )
        {
            DataTable locDt = (DataTable) chkComboLocation.Properties.DataSource;

            string atNeedCode = "";

            DataRow[] dRows = null;
            DataTable tempDt = null;

            locDt = G1.GetGroupBy(dt, "serviceLoc");

            dbCount = 0;

            for ( int i=0; i<locDt.Rows.Count; i++)
            {
                atNeedCode = locDt.Rows[i]["serviceLoc"].ObjToString();
                if ( !String.IsNullOrWhiteSpace ( atNeedCode ))
                {
                    dRows = dt.Select("serviceLoc='" + atNeedCode + "'");
                    if ( dRows.Length > 0 )
                    {
                        tempDt = dRows.CopyToDataTable();
                        Alldbs[dbCount] = tempDt;
                        dbCount++;
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void runCommissionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            DataRow dr = gridMain.GetFocusedDataRow();

            DataTable dx = dt.Clone();
            dx.ImportRow(dr);

            string who = dr["funeralType"].ObjToString();

            this.Cursor = Cursors.WaitCursor;
            FunCommissions funForm = null;

            string tempManager = "";
            string tempArranger = "";

            if ( who.ToUpper() == "MA" )
            {
                tempArranger = dr["Funeral Arranger"].ObjToString();
            }
            else if ( who.ToUpper() == "A" )
            {
                tempArranger = dr["Funeral Arranger"].ObjToString();
            }
            else if ( who.ToUpper() == "M" )
            {
                tempArranger = dr["Funeral Arranger"].ObjToString();
            }

            string str = tempArranger;
            if (!String.IsNullOrWhiteSpace(str))
            {
                string[] Lines = str.Split('[');
                if (Lines.Length > 0)
                {
                    str = Lines[0].Trim();
                    tempArranger = str;
                }
            }

            if (who.ToUpper() == "MA")
                tempManager = tempArranger;

            if (String.IsNullOrWhiteSpace(workManager) && String.IsNullOrWhiteSpace(workArranger))
            {
                funForm = new FunCommissions(dx, tempManager, "", tempArranger, who, false);
                funForm.Show();
            }
            else
            {
                funForm = new FunCommissions(dx, tempManager, "", tempArranger, who, false);
                //funForm.funCommissionClosing += FunForm_funCommissionClosing;
                funForm.ShowDialog();
            }
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void chkArranger_CheckedChanged(object sender, EventArgs e)
        {
            if (dgv.DataSource == null)
                return;
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
                return;
            try
            {
                if (byPass)
                    return;
                if (chkSort.Checked || chkManager.Checked || chkDirector.Checked )
                {
                    byPass = true;
                    chkSort.Checked = false;
                    chkManager.Checked = false;
                    chkDirector.Checked = false;
                    gridMain.Columns["manager"].GroupIndex = -1;
                    gridMain.Columns["Funeral Director"].GroupIndex = -1;
                    gridMain.Columns["serviceLoc"].GroupIndex = -1;
                    gridMain.Columns["location"].GroupIndex = -1;
                    byPass = false;
                }
                if (chkArranger.Checked)
                {
                    DataView tempview = dt.DefaultView;
                    tempview.Sort = "Funeral Arranger, lastName, firstName";
                    dt = tempview.ToTable();
                    G1.NumberDataTable(dt);
                    dgv.DataSource = dt;

                    gridMain.Columns["Funeral Arranger"].GroupIndex = 0;
                    gridMain.OptionsView.ShowFooter = true;
                    if (!gridMain.Columns["Funeral Arranger"].Visible)
                        gridMain.Columns["Funeral Arranger"].Visible = true;

                    this.gridMain.ExpandAllGroups();
                    SetupTotalsSummary();
                }
                else
                {
                    gridMain.Columns["Funeral Arranger"].GroupIndex = -1;
                    gridMain.CollapseAllGroups();
                }
                gridMain.RefreshData();
                gridMain.RefreshEditor(true);
            }
            catch (Exception ex)
            {
            }
            dgv.Refresh();
        }
        /***********************************************************************************************/
        private void chkManager_CheckedChanged(object sender, EventArgs e)
        {
            if (dgv.DataSource == null)
                return;
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
                return;
            try
            {
                if (byPass)
                    return;
                if (chkSort.Checked || chkArranger.Checked || chkDirector.Checked )
                {
                    byPass = true;
                    chkSort.Checked = false;
                    chkArranger.Checked = false;
                    chkDirector.Checked = false;
                    gridMain.Columns["Funeral Arranger"].GroupIndex = -1;
                    gridMain.Columns["Funeral Director"].GroupIndex = -1;
                    gridMain.Columns["serviceLoc"].GroupIndex = -1;
                    gridMain.Columns["location"].GroupIndex = -1;
                    byPass = false;
                }
                if (chkManager.Checked)
                {
                    DataView tempview = dt.DefaultView;
                    tempview.Sort = "manager, lastName, firstName";
                    dt = tempview.ToTable();
                    G1.NumberDataTable(dt);
                    dgv.DataSource = dt;

                    gridMain.Columns["manager"].GroupIndex = 0;
                    gridMain.OptionsView.ShowFooter = true;
                    if (!gridMain.Columns["manager"].Visible)
                        gridMain.Columns["manager"].Visible = true;

                    this.gridMain.ExpandAllGroups();
                    SetupTotalsSummary();
                }
                else
                {
                    gridMain.Columns["manager"].GroupIndex = -1;
                    gridMain.CollapseAllGroups();
                }
                gridMain.RefreshData();
                gridMain.RefreshEditor(true);
            }
            catch (Exception ex)
            {
            }
            dgv.Refresh();
        }
        /***********************************************************************************************/
        private void gridMain_CustomRowFilter(object sender, RowFilterEventArgs e)
        {
            int row = e.ListSourceRow;
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0)
                return;

            if (!chkHonorFilter.Checked )
                return;

            DateTime date = dt.Rows[row]["bonusDate"].ObjToDateTime();
            if ( date.Year > 1000 )
            {
                if (date < this.dateTimePicker1.Value || date > this.dateTimePicker2.Value)
                {
                    e.Visible = false;
                    e.Handled = true;
                }
                return;
            }
        }
        /***********************************************************************************************/
        private void chkDirector_CheckedChanged(object sender, EventArgs e)
        {
            if (dgv.DataSource == null)
                return;
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
                return;
            try
            {
                if (byPass)
                    return;
                if (chkSort.Checked || chkArranger.Checked || chkManager.Checked )
                {
                    byPass = true;
                    chkSort.Checked = false;
                    chkArranger.Checked = false;
                    chkManager.Checked = false;
                    gridMain.Columns["Funeral Arranger"].GroupIndex = -1;
                    gridMain.Columns["Funeral Director"].GroupIndex = -1;
                    gridMain.Columns["serviceLoc"].GroupIndex = -1;
                    gridMain.Columns["location"].GroupIndex = -1;
                    byPass = false;
                }
                if (chkDirector.Checked)
                {
                    DataView tempview = dt.DefaultView;
                    tempview.Sort = "Funeral Director, lastName, firstName";
                    dt = tempview.ToTable();
                    G1.NumberDataTable(dt);
                    dgv.DataSource = dt;

                    gridMain.Columns["Funeral Director"].GroupIndex = 0;
                    gridMain.OptionsView.ShowFooter = true;
                    if (!gridMain.Columns["Funeral Director"].Visible)
                        gridMain.Columns["Funeral Director"].Visible = true;

                    this.gridMain.ExpandAllGroups();
                    SetupTotalsSummary();
                }
                else
                {
                    gridMain.Columns["Funeral Director"].GroupIndex = -1;
                    gridMain.CollapseAllGroups();
                }
                gridMain.RefreshData();
                gridMain.RefreshEditor(true);
            }
            catch (Exception ex)
            {
            }
            dgv.Refresh();
        }
        /***********************************************************************************************/
        private int footerCount = 0;
        private string majorLastLocation = "";
        private string lastLocation = "";
        private string majorLastDetail = "";
        private bool firstPrint = true;
        private bool gotFooter = true;
        /***********************************************************************************************/
        private void gridMain_BeforePrintRow(object sender, DevExpress.XtraGrid.Views.Printing.CancelPrintRowEventArgs e)
        {

            if (gotFooter)
            {
                //string detail = FindLastLocation(e);
                //AddHeading(e, detail);

                gotFooter = false;
            }

            if (e.HasFooter)
            {
                footerCount++;
            }
        }
        /***********************************************************************************************/
        private void gridMain_AfterPrintRow(object sender, DevExpress.XtraGrid.Views.Printing.PrintRowEventArgs e)
        {
            if (e.HasFooter)
            {
                gotFooter = true;

                footerCount++;
                if (footerCount >= 1)
                {
                    if (!historicBonus)
                    {
                        string detail = FindLastLocation(e);
                        AddHeading((DevExpress.XtraGrid.Views.Printing.CancelPrintRowEventArgs)e, detail);
                    }

                    footerCount = 0;
                    e.PS.InsertPageBreak(e.Y);
                    //printRow = 0;
                }
            }
        }
        /***********************************************************************************************/
        //private void AddHeading(DevExpress.XtraGrid.Views.Printing.PrintRowEventArgs e, string detail)
        //{
        //    TextBrick tb = e.PS.CreateTextBrick() as TextBrick;
        //    //tb.Text = majorLastDetail;
        //    tb.Text = detail;
        //    tb.Font = new Font(tb.Font, FontStyle.Bold);
        //    tb.HorzAlignment = DevExpress.Utils.HorzAlignment.Near;
        //    //tb.Padding = new PaddingInfo(5, 0, 0, 0);
        //    tb.BackColor = Color.LightGray;
        //    tb.ForeColor = Color.Black;
        //    // Get the client page width. 
        //    SizeF clientPageSize = (e.BrickGraphics as BrickGraphics).ClientPageSize;
        //    float textBrickHeight = e.Graphics.MeasureString(tb.Text, tb.Font).Height + 4;
        //    // Calculate a rectangle for the brick and draw the brick. 
        //    tb.Padding = new PaddingInfo(0, 0, 0, 0);

        //    int y = e.Y;

        //    RectangleF textBrickRect = new RectangleF(0, y, (int)clientPageSize.Width, textBrickHeight);
        //    e.BrickGraphics.DrawBrick(tb, textBrickRect);
        //    // Adjust the current Y position to print the following row below the brick. 
        //    //e.Y += (int)textBrickHeight;

        //    //if (printRow > 0)
        //    //    e.Y += (int)textBrickHeight * printRow;
        //    printRow++;
        //}
        /***********************************************************************************************/
        private void AddHeading(DevExpress.XtraGrid.Views.Printing.CancelPrintRowEventArgs e, string detail )
        {
            TextBrick tb = e.PS.CreateTextBrick() as TextBrick;
            //tb.Text = majorLastDetail;
            tb.Text = detail;
            //Font font = tb.Font;
            //font = new Font(font.Name, 16F, FontStyle.Bold);

            tb.Font = new Font(tb.Font.Name, 16F, FontStyle.Regular);
            tb.HorzAlignment = DevExpress.Utils.HorzAlignment.Near;
            //tb.Padding = new PaddingInfo(5, 0, 0, 0);
            tb.BackColor = Color.LightGray;
            tb.ForeColor = Color.Black;
            // Get the client page width. 
            SizeF clientPageSize = (e.BrickGraphics as BrickGraphics).ClientPageSize;
            float textBrickHeight = e.Graphics.MeasureString(tb.Text, tb.Font).Height + 4;
            float textBrickWidth = e.Graphics.MeasureString(tb.Text, tb.Font).Width + 4;
            textBrickHeight = 30f;
            // Calculate a rectangle for the brick and draw the brick. 
            tb.Padding = new PaddingInfo(0, 0, 0, 0);

            int y = e.Y;
            //if (printRow >= 2)
            //    y += 5;

            //RectangleF textBrickRect = new RectangleF(0, y, (int)clientPageSize.Width, textBrickHeight);
            RectangleF textBrickRect = new RectangleF(0, y, textBrickWidth, textBrickHeight);
            e.BrickGraphics.DrawBrick(tb, textBrickRect);
            // Adjust the current Y position to print the following row below the brick. 
            // e.Y += (int)textBrickHeight;

            //if (printRow > 0)
            e.Y += (int) textBrickHeight;
            printRow++;
        }
        /***********************************************************************************************/
        private string FindLastLocation(DevExpress.XtraGrid.Views.Printing.PrintRowEventArgs e)
        {
            majorLastLocation = majorLastDetail;
            lastLocation = "";

            try
            {
                DevExpress.XtraGrid.Views.BandedGrid.AdvBandedGridView gMain = gridMain;
                DataTable dt = (DataTable)dgv.DataSource;
                int rowHandle = e.RowHandle;
                int row = gMain.GetDataSourceRowIndex(rowHandle);
                lastLocation = dt.Rows[row]["serviceLoc"].ObjToString();
                if ( lastLocation == "BK")
                {
                }

                string manager = dt.Rows[row]["manager"].ObjToString();

                //row = printRow;
                //lastLocation = summaryDt.Rows[row]["serviceLoc"].ObjToString();
                //manager = summaryDt.Rows[row]["manager"].ObjToString();

                DataRow[] dRows = summaryDt.Select("serviceLoc='" + lastLocation + "'");
                if (dRows.Length > 0)
                    majorLastDetail = dRows[0]["detail"].ObjToString();
                //DataRow[] dRows = pre2002Dt.Select("locind='" + lastLocation + "'");
                //if (dRows.Length > 0)
                //{
                //    lastLocation = dRows[0]["name"].ObjToString();
                //    majorLastDetail = lastLocation;
                //}
                //}
            }
            catch ( Exception ex)
            {
            }
            return majorLastDetail;
        }
        /***********************************************************************************************/
        private DataTable LoadArrangerData( DataTable localDt )
        {
            string cmd = "Select * from `funcommissiondata` where `name` = '" + workManager + "' OR `name` = '" + workArranger + "';";
            if ( !String.IsNullOrWhiteSpace ( workLocation ))
                cmd = "Select * from `funcommissiondata` where (`name` = '" + workManager + "' OR `name` = '" + workArranger + "' ) AND `location` = '" + workLocation + "';";
            DataTable dt = G1.get_db_data(cmd);

            if (dt.Rows.Count <= 0)
            {
                string who = "";
                string option = "";
                string data = "";
                DataRow dRow = null;

                DataTable funDt = G1.get_db_data("Select * from `funcommoptions` ORDER by `order`;");
                for (int i = 0; i < funDt.Rows.Count; i++)
                {
                    who = funDt.Rows[i]["who"].ObjToString();
                    option = funDt.Rows[i]["option"].ObjToString();
                    data = funDt.Rows[i]["defaults"].ObjToString();

                    dRow = dt.NewRow();
                    dRow["name"] = workManager;
                    dRow["ma"] = who;
                    dRow["option"] = option;
                    dRow["answer"] = data;
                    dt.Rows.Add(dRow);
                }
            }

            dt.Columns.Add("num");
            dt.Columns.Add("mod");
            G1.NumberDataTable(dt);

            string what = "";

            for (int i = (dt.Rows.Count - 1); i >= 0; i--)
            {
                what = dt.Rows[i]["ma"].ObjToString();
                if (what.ToUpper() != "A" )
                    dt.Rows.RemoveAt(i);
            }

            LoadFuneralDetails(dt, localDt );

            return dt;
        }
        /***********************************************************************************************/
        private void LoadFuneralDetails(DataTable dt, DataTable localDt )
        {
            string option = "";
            string answer = "";
            string ma = "";
            double count = 0D;
            double detail = 0D;
            bool processOption = false;

            if (G1.get_column_number(dt, "count") < 0)
                dt.Columns.Add("count", Type.GetType("System.Double"));
            if (G1.get_column_number(dt, "detail") < 0)
                dt.Columns.Add("detail", Type.GetType("System.Double"));
            if (G1.get_column_number(dt, "commission") < 0)
                dt.Columns.Add("commission", Type.GetType("System.Double"));

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    option = dt.Rows[i]["option"].ObjToString();
                    answer = dt.Rows[i]["answer"].ObjToString();
                    ma = dt.Rows[i]["ma"].ObjToString().ToUpper();

                    processOption = true;
                    //if (workWho.ToUpper() == "MA")
                    //    processOption = true;
                    //else if (workWho.ToUpper() == "M" && ma == "M")
                    //    processOption = true;
                    //else if (workWho.ToUpper() == "A" && ma == "A")
                    //    processOption = true;

                    if (!processOption)
                        continue;

                    ParseOutOption(localDt, option, answer, ma, ref count, ref detail);

                    dt.Rows[i]["count"] = count;
                    dt.Rows[i]["detail"] = detail;
                    if (detail > 0D)
                    {
                        if (option != "Funeral Average")
                            dt.Rows[i]["commission"] = detail;
                    }
                }
                catch ( Exception ex)
                {
                }
            }

            count = 0D;
            for (int i = 0; i < localDt.Rows.Count; i++)
            {
                answer = localDt.Rows[i]["funeralType"].ObjToString();
                if (answer.IndexOf("M") >= 0)
                {
                    answer = localDt.Rows[i]["gotPackage"].ObjToString();
                    if (String.IsNullOrWhiteSpace(answer))
                    {
                        answer = localDt.Rows[i]["urn"].ObjToString();
                        detail = answer.ObjToDouble();
                        if (detail == 0D)
                            count++;
                    }
                }
            }

            double dollarsPerFuneral = 0D;
            double minimumFunerals = 0D;
            double funeralAverage = 0D;
            double averageMinimum = 0D;
            DataRow[] dRows = dt.Select("option='Funeral Average'");
            if (dRows.Length > 0)
            {
                funeralAverage = dRows[0]["detail"].ObjToDouble();
                averageMinimum = dRows[0]["answer"].ObjToDouble();
            }

            if (funeralAverage > averageMinimum)
            {
                dRows = dt.Select("option='Minimum Funerals'");
                if (dRows.Length > 0)
                {
                    minimumFunerals = dRows[0]["answer"].ObjToDouble();
                    dRows[0]["count"] = count;
                    dRows[0]["detail"] = count - minimumFunerals;
                }

                dRows = dt.Select("option='Dollars per Funeral'");
                if (dRows.Length > 0)
                    dollarsPerFuneral = dRows[0]["answer"].ObjToDouble();


                dRows = dt.Select("option='Dollars per Funeral'");
                if (dRows.Length > 0)
                {
                    dRows[0]["count"] = (count - minimumFunerals);
                    if (count > minimumFunerals)
                    {
                        dRows[0]["detail"] = (count - minimumFunerals) * dollarsPerFuneral;
                        dRows[0]["commission"] = (count - minimumFunerals) * dollarsPerFuneral;
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void ParseOutOption ( DataTable dt, string option, string answer, string ma, ref double count, ref double detail)
        {
            count = 0D;
            detail = 0D;
            string who = "";
            string str = "";
            double dValue = 0D;
            double gauge = 0D;
            double totalDiscount = 0D;
            string[] Lines = null;
            if (option == "Funeral Average")
            {
                double netFuneral = 0D;
                double totalNet = 0D;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    totalDiscount = dt.Rows[i]["totalDiscount"].ObjToDouble();
                    if (totalDiscount > 0D)
                        continue;
                    who = dt.Rows[i]["funeralType"].ObjToString();
                    who = "A";
                    if (who == ma || who == "MA" && ma == "M")
                    {
                        str = dt.Rows[i]["funeralType"].ObjToString();
                        if (str.IndexOf("M") >= 0)
                        {
                            str = dt.Rows[i]["gotPackage"].ObjToString();
                            if (String.IsNullOrWhiteSpace(str))
                            {
                                str = dt.Rows[i]["urn"].ObjToString();
                                dValue = str.ObjToDouble();
                                if (dValue == 0D)
                                {
                                    count++;
                                    netFuneral = dt.Rows[i]["netFuneral"].ObjToDouble();
                                    totalNet += netFuneral;
                                }
                            }
                        }
                    }
                }
                if (count > 0D)
                {
                    detail = totalNet / count;
                    detail = G1.RoundValue(detail);
                }
            }
            else if (option == "Vault")
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    totalDiscount = dt.Rows[i]["totalDiscount"].ObjToDouble();
                    if (totalDiscount > 0D)
                        continue;
                    who = dt.Rows[i]["funeralType"].ObjToString();
                    who = "A";
                    if (who == ma || who == "MA" && ma == "M")
                    {
                        str = dt.Rows[i]["vault"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(str))
                            count++;
                    }
                    detail = count * answer.ObjToDouble();
                }
            }
            else if (option == "Urn")
            {
                double totalUrn = 0D;
                answer = answer.Replace("%", "");
                double percent = 0D;
                if (G1.validate_numeric(answer))
                    percent = answer.ObjToDouble() / 100D;
                if (percent > 0D)
                {
                    count = 0;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        totalDiscount = dt.Rows[i]["totalDiscount"].ObjToDouble();
                        if (totalDiscount > 0D)
                            continue;
                        who = dt.Rows[i]["funeralType"].ObjToString();
                        who = "A";
                        if (who == ma || who == "MA" && ma == "M")
                        {
                            str = dt.Rows[i]["urn"].ObjToString();
                            if (!String.IsNullOrWhiteSpace(str))
                            {
                                dValue = str.ObjToDouble();
                                if (dValue > 0D)
                                {
                                    dValue = dValue * percent;
                                    totalUrn += dValue;
                                    count++;
                                }
                            }
                        }
                        if (count > 0D)
                            detail = totalUrn;
                    }
                }
            }
            else if (option.ToUpper().IndexOf("CASKET GAUGE") == 0)
            {
                option = ParseOutEqual(option);
                option = option.ToUpper().Replace("GAUGE", "").Trim();
                gauge = option.ObjToDouble();
                count = 0;
                if (gauge > 0D)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        totalDiscount = dt.Rows[i]["totalDiscount"].ObjToDouble();
                        if (totalDiscount > 0D)
                            continue;
                        who = dt.Rows[i]["funeralType"].ObjToString();
                        who = "A";
                        if (who == ma || who == "MA" && ma == "A")
                        {
                            str = dt.Rows[i]["casketgauge"].ObjToString();
                            if (!String.IsNullOrWhiteSpace(str))
                            {
                                dValue = str.ObjToDouble();
                                if (dValue == gauge)
                                    count++;
                            }
                        }
                    }
                }
                detail = count * answer.ObjToDouble();
            }
            else if (option.ToUpper().IndexOf("CASKET TYPE") == 0)
            {
                count = 0;
                option = ParseOutEqual(option);
                if (!String.IsNullOrWhiteSpace(option))
                {
                    if (option.IndexOf("+") < 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            totalDiscount = dt.Rows[i]["totalDiscount"].ObjToDouble();
                            if (totalDiscount > 0D)
                                continue;
                            who = dt.Rows[i]["funeralType"].ObjToString();
                            who = "A";
                            if (who == ma || who == "MA" && ma == "A")
                            {
                                str = dt.Rows[i]["caskettype"].ObjToString();
                                if (str.ToUpper().Trim() == option.ToUpper().Trim())
                                    count++;
                            }
                        }
                    }
                    else
                    {
                        Lines = option.Split('+');
                        if (Lines.Length > 1)
                        {
                            option = Lines[0].ToUpper().Trim();
                            Lines = Lines[1].Trim().Split(',');
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                totalDiscount = dt.Rows[i]["totalDiscount"].ObjToDouble();
                                if (totalDiscount > 0D)
                                    continue;
                                who = dt.Rows[i]["funeralType"].ObjToString();
                                who = "A";
                                if (who == ma || who == "MA" && ma == "A")
                                {
                                    str = dt.Rows[i]["caskettype"].ObjToString();
                                    if (str.ToUpper().Trim() == option.ToUpper().Trim())
                                    {
                                        str = dt.Rows[i]["casketdesc"].ObjToString().Trim().ToUpper();
                                        for (int j = 0; j < Lines.Length; j++)
                                        {
                                            if (!String.IsNullOrWhiteSpace(Lines[j]))
                                            {
                                                if (str.IndexOf(Lines[j].Trim().ToUpper()) >= 0)
                                                    count++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                detail = count * answer.ObjToDouble();
            }
        }
        /***********************************************************************************************/
        private string ParseOutEqual(string option)
        {
            string[] Lines = option.Split('=');
            if (Lines.Length <= 1)
                return option;
            option = Lines[1].Trim();
            return option;
        }
        /***********************************************************************************************/
        private void showAgentDetailsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);
            string contractNumber = dr["contractNumber"].ObjToString();
            string ma = workDt.Rows[0]["ma"].ObjToString();

            DataTable dt = (DataTable)dgv.DataSource;
            DataTable localDt = dt.Clone();
            localDt.ImportRow(dr);

            if (workingManagers)
                workArranger = dr["manager"].ObjToString();
            else if (workingArrangers)
            {
                workArranger = dr["Funeral Arranger"].ObjToString();
                workArranger = CleanupArranger(workArranger);
            }

            DataTable tempDt = LoadArrangerData(localDt);

            tempDt = LoadExceptions(contractNumber, tempDt);

            FunCommissions funCommForm = new FunCommissions(tempDt, workArranger );
            funCommForm.Show();
        }
        /***********************************************************************************************/
        private DataTable LoadExceptions ( string contractNumber, DataTable tempDt )
        {
            FunServices funForm = new FunServices(contractNumber);
            DataTable funDt = funForm.funServicesDT;

            double price = 0D;
            double currentPrice = 0D;
            string type = "";
            string service = "";
            DataRow[] dRows = null;

            DataTable exceptionDt = G1.get_db_data("Select * from `funeral_master`;");

            for (int i = 0; i < funDt.Rows.Count; i++)
            {
                if (funDt.Rows[i]["ignore"].ObjToString().ToUpper() == "Y")
                    continue;
                price = funDt.Rows[i]["price"].ObjToDouble();

                if (price == 0D)
                    continue;

                currentPrice = funDt.Rows[i]["currentPrice"].ObjToDouble();

                type = funDt.Rows[i]["type"].ObjToString().ToUpper();
                service = funDt.Rows[i]["service"].ObjToString().ToUpper();
                dRows = exceptionDt.Select("service='" + service + "'");
                if ( dRows.Length > 0 )
                {
                    if (dRows.Length > 0 && type.ToUpper() != "CASH ADVANCE")
                    {
                        if (dRows[0]["asService"].ObjToString() == "1")
                            tempDt = AddException ( service, currentPrice, "asService", tempDt );
                        if (dRows[0]["fromService"].ObjToString() == "1")
                        {
                            if (type == "SERVICE")
                                tempDt = AddException(service, currentPrice, "fromService", tempDt);
                        }
                        if (dRows[0]["fromMerc"].ObjToString() == "1")
                        {
                            if (type == "MERCHANDISE")
                                tempDt = AddException(service, currentPrice, "fromMerc", tempDt);
                        }
                        if (dRows[0]["asCash"].ObjToString() == "1")
                        {
                            if (type.ToUpper() == "MERCHANDISE")
                            {
                                tempDt = AddException(service, currentPrice, "asCash", tempDt);
                                tempDt = AddException(service, currentPrice, "asService", tempDt);
                            }
                            else
                                tempDt = AddException(service, currentPrice, "asCash", tempDt);
                        }
                        if (dRows[0]["asNothing"].ObjToString() == "1")
                            tempDt = AddException(service, currentPrice, "asNothing", tempDt);
                        if (dRows[0]["asMerc"].ObjToString() == "1")
                        {
                            if (type.ToUpper() != "MERCHANDISE")
                                tempDt = AddException(service, currentPrice, "asMerc", tempDt);
                        }
                    }
                    else
                    {
                        if (service.ToUpper().IndexOf("MILEAGE") >= 0)
                        {
                            if (type.ToUpper() != "CASH ADVANCE")
                            {
                                if (type.ToUpper() == "MERCHANDISE")
                                    tempDt = AddException(service, currentPrice, "fromMerc", tempDt);
                                else
                                    tempDt = AddException(service, currentPrice, "asCash", tempDt);
                            }
                        }
                        else if (service.ToUpper().IndexOf("ENGRAV") >= 0)
                        {
                            if (type.ToUpper() != "CASH ADVANCE")
                            {
                                if (type.ToUpper() == "MERCHANDISE")
                                {
                                    tempDt = AddException(service, currentPrice, "asCash", tempDt);
                                    tempDt = AddException(service, currentPrice, "asService", tempDt);
                                }
                                else
                                    tempDt = AddException(service, currentPrice, "asCash", tempDt);
                            }
                        }
                        else if (service.ToUpper().IndexOf("BOOKMARK") >= 0)
                        {
                            if (type.ToUpper() != "CASH ADVANCE")
                            {
                                if (type.ToUpper() == "MERCHANDISE")
                                {
                                    tempDt = AddException(service, currentPrice, "asCash", tempDt);
                                    tempDt = AddException(service, currentPrice, "asService", tempDt);
                                }
                                else
                                    tempDt = AddException(service, currentPrice, "asCash", tempDt);
                            }
                        }
                        else if (service.ToUpper().IndexOf("SHIPPING") >= 0)
                        {
                            if (type.ToUpper() != "CASH ADVANCE")
                            {
                                if (type.ToUpper() == "MERCHANDISE")
                                    tempDt = AddException(service, currentPrice, "fromMerc", tempDt);
                                else
                                    tempDt = AddException(service, currentPrice, "asCash", tempDt);
                            }
                        }
                        else if (type.ToUpper() == "SERVICE")
                        {
                            if (service.ToUpper().IndexOf("D-") == 0)
                            {
                                if (service.ToUpper().IndexOf("INFANT CASKET") < 0)
                                    tempDt = AddException(service, currentPrice, "asCash", tempDt);
                            }
                        }
                        else
                        {
                        }
                    }
                }
            }

            return tempDt;
        }
        /***********************************************************************************************/
        private DataTable AddException ( string service, double currentPrice, string what, DataTable tempDt)
        {
            DataRow dRow = tempDt.NewRow();
            dRow["option"] = what + " " + service;
            dRow["detail"] = currentPrice.ObjToDouble();
            tempDt.Rows.Add(dRow);
            return tempDt;
        }
        /***********************************************************************************************/
        private string CleanupArranger ( string arranger )
        {
            int idx = arranger.IndexOf("[");
            if (idx > 0)
                arranger = arranger.Substring(0, idx);
            return arranger;
        }
        /***********************************************************************************************/
        private void chkPreparePrint_CheckedChanged(object sender, EventArgs e)
        {
            if ( chkPreparePrint.Checked )
            {
                gridMain.Columns["funeralType"].Visible = false;
                gridMain.Columns["urn"].Visible = false;
                gridMain.Columns["casketgauge"].Visible = false;
                gridMain.Columns["caskettype"].Visible = false;

                gridMain.Columns["casketdesc"].Visible = false;

                gridMain.Columns["upgrade"].Visible = false;
                gridMain.Columns["premiumType"].Visible = false;
                gridMain.Columns["gotPackage"].Visible = false;
            }
            else
            {
                gridMain.Columns["funeralType"].Visible = true;
                gridMain.Columns["urn"].Visible = true;
                gridMain.Columns["casketgauge"].Visible = true;
                gridMain.Columns["casketdesc"].Visible = true;
                gridMain.Columns["upgrade"].Visible = true;
                gridMain.Columns["premiumType"].Visible = true;
                gridMain.Columns["gotPackage"].Visible = true;
            }

            gridMain.RefreshEditor(true);
            dgv.Refresh();
        }
        /***********************************************************************************************/
        private void gridMain_CustomDrawGroupRow(object sender, RowObjectCustomDrawEventArgs e)
        {
            var view = (GridView)sender;
            var info = (GridGroupRowInfo)e.Info;
            var caption = info.Column.Caption;
            if (info.Column.Caption == string.Empty)
            {
                caption = info.Column.ToString();
            }

            if ( summaryDt == null )
            {

            }

            string serviceLoc = info.GroupValueText.Trim();
            if (!String.IsNullOrWhiteSpace(serviceLoc) && !historicBonus )
            {
                DataRow [] dRows = summaryDt.Select("serviceLoc='" + serviceLoc + "'");
                if (dRows.Length > 0)
                {
                    string manager = dRows[0]["manager"].ObjToString();
                    string location = dRows[0]["location"].ObjToString();
                    double totalFunerals = dRows[0]["totalFuneral"].ObjToDouble();
                    double netFunerals = dRows[0]["netFuneral"].ObjToDouble();
                    double met = dRows[0]["met"].ObjToDouble();
                    double average = dRows[0]["average"].ObjToDouble();
                    double calcavg = dRows[0]["calcavg"].ObjToDouble();
                    double other = dRows[0]["other"].ObjToDouble();
                    string str = $"{caption} : {info.GroupValueText} ({view.GetChildRowCount(e.RowHandle)})";
                    str += " F=" + location + " M=" + manager + " Met = " + met.ToString() + " Avg Needed= " + G1.ReformatMoney (average) + " CalcAvg " + G1.ReformatMoney(calcavg) + " TotFun=" + G1.ReformatMoney(totalFunerals) + " NetFun=" + G1.ReformatMoney(netFunerals) + " Other=" + G1.ReformatMoney(other);

                    //dRows[0]["detail"] = str;
                    info.GroupText = str;
                }
                else
                    info.GroupText = $"{caption} : {info.GroupValueText} ({view.GetChildRowCount(e.RowHandle)})";
            }
            else
                info.GroupText = $"{caption} : {info.GroupValueText} ({view.GetChildRowCount(e.RowHandle)})";
        }
        /***********************************************************************************************/
        private void gridMain_DoubleClick(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            if (!String.IsNullOrWhiteSpace(contract))
            {
                this.Cursor = Cursors.WaitCursor;

                Form form = G1.IsFormOpen("EditCust", contract);
                if (form != null)
                {
                    form.Show();
                    form.WindowState = FormWindowState.Maximized;
                    form.Visible = true;
                    form.BringToFront();
                }
                else
                {
                    EditCust custForm = new EditCust(contract);
                    custForm.Tag = contract;
                    //custForm.custClosing += CustForm_custClosing;
                    custForm.Show();
                }
                //FunPayments editFunPayments = new FunPayments(null, contract, "", false, true );
                //editFunPayments.Show();

                //CustomerDetails clientForm = new CustomerDetails(contract);
                //clientForm.Show();
                this.Cursor = Cursors.Default;
            }
        }
        /***********************************************************************************************/
        private void excludeFuneralMenu_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);
            string contractNumber = dr["contractNumber"].ObjToString();
            string serviceId = dr["serviceId"].ObjToString();

            string record = dr["record"].ObjToString();
            string exclude = dr["exclude"].ObjToString();
            if (exclude == "Y")
                exclude = "";
            else
                exclude = "Y";
            dr["exclude"] = exclude;
            G1.update_db_table("fcust_extended", "record", record, new string[] {"exclude", exclude });

            gridMain.RefreshEditor(true);


            if (groupDt != null)
                groupDt.Rows.Clear();

            allowSummary = true;
            gridMain.UpdateSummary();
            allowSummary = false;
            //gridMain.UpdateGroupSummary();

            dgv.Refresh();
        }
        /***********************************************************************************************/
        private void gridMain_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView View = sender as GridView;
            if (e.Column.FieldName.ToUpper().IndexOf("NAME") >= 0)
            {
                Font font = e.Appearance.Font;
                string str = View.GetRowCellValue(e.RowHandle, "exclude").ObjToString();
                if ( String.IsNullOrWhiteSpace ( str ))
                {
                    e.Appearance.ForeColor = Color.Black;
                    font = new Font(font.Name, font.Size, FontStyle.Regular);
                    e.Appearance.Font = font;

                }
                else if (str == "Y")
                {
                    e.Appearance.ForeColor = Color.Red;
                    font = new Font(font.Name, font.Size, FontStyle.Italic );
                    e.Appearance.Font = font;
                }
            }
        }
        /***********************************************************************************************/
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);
            string contractNumber = dr["contractNumber"].ObjToString();
            string serviceId = dr["serviceId"].ObjToString();

            string record = dr["record"].ObjToString();
            string exclude = dr["exclude"].ObjToString();
            DateTime date = dr["bonusDate"].ObjToDateTime();
            using (GetDate dateForm = new GetDate(date, "Enter Bonus Date"))
            {
                dateForm.TopMost = true;
                dateForm.ShowDialog();
                if (dateForm.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    date = dateForm.myDateAnswer;
                    dr["bonusDate"] = G1.DTtoMySQLDT(date);
                    if (!String.IsNullOrWhiteSpace(record))
                    {
                        G1.update_db_table("fcust_extended", "record", record, new string[] { "bonusDate", date.ToString("yyyy-MM-dd") });
                    }
                    //DataChanged();
                    gridMain.ClearSelection();
                    gridMain.FocusedRowHandle = rowHandle;

                    gridMain.RefreshData();
                    gridMain.RefreshEditor(true);
                    gridMain.SelectRow(rowHandle);

                    //if (date <= this.dateTimePicker1.Value || date >= this.dateTimePicker2.Value )
                    //{
                    //    dt.Rows.Remove(dr);

                    //    gridMain.DeleteRow( rowHandle);

                    //    gridMain.RefreshData();
                    //    gridMain.RefreshEditor(true);
                    //}
                }
                else if (dateForm.DialogResult == System.Windows.Forms.DialogResult.Cancel)
                {
                }
            }
        }
        /***********************************************************************************************/
        private void chkHonorFilter_CheckedChanged(object sender, EventArgs e)
        {
            gridMain.RefreshData();
            gridMain.RefreshEditor(true);

            gridMain.UpdateSummary();

            CalcSummaryAverages();
        }
        /***********************************************************************************************/
        private void btnSaveBonus_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to save this data to the database?", "Save Bonus Data Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
                return;

            DataTable dt = (DataTable)dgv.DataSource;

            this.Cursor = Cursors.WaitCursor;

            if (!SaveData(dt))
                return;

            btnSaveBonus.Hide();
            btnSaveBonus.Refresh();
        }
        /***********************************************************************************************/
        private bool DeletePreviousData()
        {
            bool success = true;

            DateTime date = this.dateTimePicker2.Value;

            string date1 = date.ToString("yyyyMMdd");

            string cmd = "DELETE from `funeralBonuses` where `runDate` = '" + date1 + "';";

            try
            {
                G1.get_db_data(cmd);
            }
            catch (Exception ex)
            {
                MessageBox.Show("***ERROR*** Delete Previous Data " + ex.Message.ToString());
                success = false;
            }
            return success;
        }
        /***********************************************************************************************/
        private bool SaveData(DataTable dt)
        {
            this.Cursor = Cursors.WaitCursor;

            if (!DeletePreviousData())
            {
                this.Cursor = Cursors.Default;
                return false;
            }

            DataTable saveDt = new DataTable();

            G1.duplicate_dt_column(dt, "tmstamp", saveDt);
            G1.duplicate_dt_column(dt, "record", saveDt);
            G1.duplicate_dt_column(dt, "runDate", saveDt);
            G1.duplicate_dt_column(dt, "serviceId", saveDt);
            G1.duplicate_dt_column(dt, "contractNumber", saveDt);
            G1.duplicate_dt_column(dt, "serviceDate", saveDt);
            G1.duplicate_dt_column(dt, "serviceLoc", saveDt);
            G1.duplicate_dt_column(dt, "name", saveDt);
            G1.duplicate_dt_column(dt, "deceasedDate", saveDt);
            G1.duplicate_dt_column(dt, "funeralType", saveDt);
            G1.duplicate_dt_column(dt, "gotPackage", saveDt);
            G1.duplicate_dt_column(dt, "casket", saveDt);
            G1.duplicate_dt_column(dt, "vault", saveDt);
            G1.duplicate_dt_column(dt, "casketgauge", saveDt);
            G1.duplicate_dt_column(dt, "caskettype", saveDt);
            G1.duplicate_dt_column(dt, "premiumType", saveDt);
            G1.duplicate_dt_column(dt, "casketdesc", saveDt);

            G1.duplicate_dt_column(dt, "amountReceived", saveDt );
            G1.duplicate_dt_column(dt, "amountFiled", saveDt);
            G1.duplicate_dt_column(dt, "custPrice", saveDt);
            G1.duplicate_dt_column(dt, "cust_Merchandise", saveDt);
            G1.duplicate_dt_column(dt, "custServices", saveDt);
            G1.duplicate_dt_column(dt, "totalDiscount", saveDt);
            G1.duplicate_dt_column(dt, "currentPrice", saveDt);
            G1.duplicate_dt_column(dt, "currentMerchandise", saveDt);
            G1.duplicate_dt_column(dt, "currentServices", saveDt);
            G1.duplicate_dt_column(dt, "balanceDue", saveDt);
            G1.duplicate_dt_column(dt, "additionalDiscount", saveDt);
            G1.duplicate_dt_column(dt, "classa", saveDt);
            G1.duplicate_dt_column(dt, "grossAmountReceived", saveDt);
            G1.duplicate_dt_column(dt, "amountDiscount", saveDt);
            G1.duplicate_dt_column(dt, "amountGrowth1", saveDt);
            G1.duplicate_dt_column(dt, "cashAdvance", saveDt);
            G1.duplicate_dt_column(dt, "trustAmountFiled", saveDt);
            G1.duplicate_dt_column(dt, "trustAmountReceived", saveDt);
            G1.duplicate_dt_column(dt, "insAmountFiled", saveDt);
            G1.duplicate_dt_column(dt, "insAmountReceived", saveDt);
            G1.duplicate_dt_column(dt, "trustPayment", saveDt);
            G1.duplicate_dt_column(dt, "netFuneral", saveDt);
            G1.duplicate_dt_column(dt, "cashCheck", saveDt);
            G1.duplicate_dt_column(dt, "cc", saveDt);
            G1.duplicate_dt_column(dt, "casketAmount", saveDt);
            G1.duplicate_dt_column(dt, "vaultAmount", saveDt);
            G1.duplicate_dt_column(dt, "casketCost", saveDt);
            G1.duplicate_dt_column(dt, "vaultCost", saveDt);
            G1.duplicate_dt_column(dt, "endingBalance", saveDt);
            G1.duplicate_dt_column(dt, "upgrade", saveDt);
            G1.duplicate_dt_column(dt, "otherBonuses", saveDt);
            G1.duplicate_dt_column(dt, "urn", saveDt);
            G1.duplicate_dt_column(dt, "SRVLOC", saveDt);
            G1.duplicate_dt_column(dt, "Funeral Director", saveDt);
            G1.duplicate_dt_column(dt, "Funeral Arranger", saveDt);
            G1.duplicate_dt_column(dt, "Funeral Creator", saveDt);
            G1.duplicate_dt_column(dt, "pendingComment", saveDt);
            G1.duplicate_dt_column(dt, "contractType", saveDt);
            G1.duplicate_dt_column(dt, "funeral_classification", saveDt);
            G1.duplicate_dt_column(dt, "manager", saveDt);
            G1.duplicate_dt_column(dt, "asService", saveDt);
            G1.duplicate_dt_column(dt, "asCash", saveDt);
            G1.duplicate_dt_column(dt, "asNothing", saveDt);
            G1.duplicate_dt_column(dt, "asMerc", saveDt);
            G1.duplicate_dt_column(dt, "exclude", saveDt);
            G1.duplicate_dt_column(dt, "bonusDate", saveDt);
            G1.duplicate_dt_column(dt, "fromService", saveDt);
            G1.duplicate_dt_column(dt, "fromMerc", saveDt);
            G1.duplicate_dt_column(dt, "urnCost", saveDt);
            G1.duplicate_dt_column(dt, "urnDesc", saveDt);

            saveDt = CleanupTable(saveDt);

            Structures.TieDbTable("funeralBonuses", saveDt);

            string record = "";
            string runDate = "";
            string serviceId = "";
            string contractNumber = "";

            string updateLine = "";
            string field = "";
            string data = "";
            string cmd = "";
            string type = "";

            int lastRow = saveDt.Rows.Count;
            barImport.Show();
            barImport.Minimum = 0;
            barImport.Maximum = lastRow;
            barImport.Refresh();

            for ( int i=0; i<saveDt.Rows.Count; i++ )
            {

                barImport.Value = i + 1;
                barImport.Refresh();

                runDate = saveDt.Rows[i]["runDate"].ObjToString();
                serviceId = saveDt.Rows[i]["serviceId"].ObjToString();


                record = G1.create_record("funeralBonuses", "serviceId", serviceId);
                if (G1.BadRecord("funeralBonuses", record))
                    break;

                updateLine = "";
                for ( int j=0; j<saveDt.Columns.Count; j++)
                {
                    field = saveDt.Columns[j].ColumnName.ObjToString();
                    if (field == "record")
                        continue;
                    if (field == "tmstamp")
                        continue;

                    type = saveDt.Columns[j].DataType.ObjToString().ToUpper();

                    if (type == "SYSTEM.DATE")
                        data = saveDt.Rows[i][j].ObjToDateTime().ToString("yyyyMMdd");
                    else if (type == "SYSTEM.DATETIME")
                        data = saveDt.Rows[i][j].ObjToDateTime().ToString("yyyyMMddHHmmss");
                    else if (type == "MYSQL.DATA.TYPES.MYSQLDATETIME")
                    {
                        data = saveDt.Rows[i][j].ObjToDateTime().ToString("yyyyMMddHHmmss");
                    }
                    else
                        data = saveDt.Rows[i][j].ObjToString();

                    updateLine = updateLine += "`" + field + "`='" + data + "',";
                }

                updateLine = updateLine.TrimEnd(',');

                cmd = "UPDATE `funeralBonuses` SET " + updateLine + " WHERE `record` = '" + record + "';";
                G1.update_db_data(cmd);
            }

            barImport.Value = lastRow;
            barImport.Refresh();

            barImport.Hide();
            barImport.Refresh();

            saveDt.Dispose();
            saveDt = null;

            this.Cursor = Cursors.Default;
            return true;
        }
        /***********************************************************************************************/
        private DataTable CleanupTable ( DataTable dt )
        {
            string str = "";
            for ( int i=0; i<dt.Columns.Count; i++)
            {
                if ( dt.Columns[i].DataType.ToString() == "System.String" )
                {
                    for ( int j=0; j<dt.Rows.Count; j++)
                    {
                        try
                        {
                            str = dt.Rows[j][i].ObjToString();
                            if (!String.IsNullOrWhiteSpace(str))
                            {
                                if (str.IndexOf("\n") >= 0)
                                    str = str.Replace("\n", "");
                                dt.Rows[j][i] = G1.try_protect_data(str);
                                if (str.Length > 79)
                                    dt.Rows[j][i] = G1.Truncate(dt.Rows[j][i].ObjToString(), 80);
                            }
                        }
                        catch ( Exception ex)
                        {
                        }

                    }
                }
            }
            return dt;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            historicBonus = true;

            DateTime date = dateTimePicker2.Value;

            DateTime firstDate = new DateTime(date.Year, date.Month, 1);

            string date1 = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string date2 = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DialogResult result = MessageBox.Show("Are you sure you want to READ OLD Bonus Data for " + date.ToString("MM/dd/yyyy") + "?", "Read Bonus Report Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
                return;

            this.Cursor = Cursors.WaitCursor;

            string cmd = "Select * from `funeralBonuses` WHERE `runDate` >= '" + date1.ObjToDateTime().ToString("yyyyMMdd") + "' AND `runDate` <= '" + date2.ObjToDateTime().ToString("yyyyMMdd") + "';";
            DataTable dx = G1.get_db_data(cmd);


            cmd = "SELECT * FROM `fcust_extended` e LEFT JOIN `fcontracts` p ON p.`contractNumber` = e.`contractNumber` left join `fcustomers` d ON d.`contractNumber` = e.`contractNumber` LEFT JOIN `icontracts` i ON i.`contractNumber` = e.`contractNumber` LEFT JOIN `icustomers` j ON j.`contractNumber` = e.`contractNumber` ";
            cmd += " WHERE e.`ServiceID` <> '' ";

            if (!String.IsNullOrWhiteSpace(serviceLocs))
                cmd += " AND ( e.`serviceLoc` IN (" + serviceLocs + ") OR e.`Funeral Arranger` LIKE '" + workArrangerFirstName + "%' AND e.`Funeral Arranger` LIKE  '%" + workArrangerLastName + "%' )";
            else
                cmd += " AND e.`Funeral Arranger` LIKE '" + workArrangerFirstName + "%' AND e.`Funeral Arranger` LIKE  '%" + workArrangerLastName + "%' ";

            if (chkUseDates.Checked || chkDeceasedDate.Checked)
            {
                if (chkDeceasedDate.Checked)
                    cmd += " AND ( (p.`deceasedDate` >= '" + date1 + "' AND p.`deceasedDate` <= '" + date2 + "' ) OR ( e.`bonusDate` >= '" + date1 + "' AND e.`bonusDate` <= '" + date2 + "' ) )";
                else
                    cmd += " AND ( (e.`serviceDate` >= '" + date1 + "' AND e.`serviceDate` <= '" + date2 + "' ) OR ( e.`bonusDate` >= '" + date1 + "' AND e.`bonusDate` <= '" + date2 + "' ) )";
            }

            cmd += " ORDER BY e.`serviceDate` DESC ";
            cmd += " LIMIT 1 ";
            cmd += ";";

            DataTable dt = G1.get_db_data(cmd);
            dt.Rows.Clear(); // Just needed the format

            this.Cursor = Cursors.WaitCursor;
            string name = "";
            for (int i = 0; i < dx.Columns.Count; i++)
            {
                name = dx.Columns[i].ColumnName.Trim();
                G1.duplicate_dt_column(dx, name, dt);
            }

            G1.NumberDataTable(dt);

            this.Cursor = Cursors.WaitCursor;
            LoadDBs(dt);

            originalDt = dt.Copy();
            dgv.DataSource = dt;
            this.Cursor = Cursors.WaitCursor;
            ScaleCells();

            if (summaryDt != null)
                summaryDt.Rows.Clear();

            this.Cursor = Cursors.WaitCursor;
            if ( this.dateTimePicker1.Value == firstDate )
                CalcSummaryAverages();

            chkSort.Checked = true;
            //chkSort_CheckedChanged(null, null);

            gridMain.Columns["serviceId"].SortOrder = ColumnSortOrder.Ascending;

            gridMain.RefreshData();
            gridMain.RefreshEditor(true);

            if (this.dateTimePicker1.Value == firstDate)
                gridMain.UpdateSummary();

            gridMain.ExpandAllGroups();

            menuStrip1.BackColor = Color.Pink;

            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private string CleanServiceId ( string serviceId )
        {
            if (!String.IsNullOrWhiteSpace(serviceId ))
            {
                string c = serviceId.Substring(0, 1);
                if (c == "Y")
                    serviceId = serviceId.Substring(1);
            }
            return serviceId;
        }
        /***********************************************************************************************/
        private void btnCompare_Click(object sender, EventArgs e)
        {
            CompareTimJones(true);
        }
        /***********************************************************************************************/
        private void CompareTimJones ( bool pullFile = false )
        {
            double cost = 0D;
            string str = "";
            if (timJonesDt == null || pullFile )
            {

                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string file = ofd.FileName;

                        this.Cursor = Cursors.WaitCursor;
                        timJonesDt = ExcelWriter.ReadFile2(file, 0, "Sheet1");
                    }
                }
            }

            this.Cursor = Cursors.Default;

            if (timJonesDt.Rows.Count <= 0)
                return;

            this.Cursor = Cursors.WaitCursor;

            DataTable dx = (DataTable)dgv.DataSource;

            string serviceId = "";
            double casketCost = 0D;
            double vaultCost = 0D;
            double netFuneral = 0D;
            double cashAdvance = 0D;
            double other = 0D;
            double totalFuneral = 0D;
            double services = 0D;

            string timServiceId = "";
            double timCasketCost = 0D;
            double timVaultCost = 0D;
            double timTotalNet = 0D;
            double timCashAdvance = 0D;
            double timOther = 0D;
            double timTotalFuneral = 0D;
            double timServices = 0D;

            DataRow[] dRows = null;
            DataRow dRow = null;

            string status = "";

            if (badDt != null)
                badDt.Rows.Clear();

            double urnCost = 0D;

            for (int i = 1; i < dx.Rows.Count; i++)
            {
                status = "";
                serviceId = dx.Rows[i]["serviceId"].ObjToString();
                serviceId = CleanServiceId(serviceId);
                if (serviceId == "HA25008")
                {
                }
                dRows = timJonesDt.Select("Column1='" + serviceId + "'");
                if (dRows.Length > 0)
                {
                    try
                    {
                        //casketCost = dx.Rows[i]["casketCost"].ObjToDouble();
                        //urnCost = dx.Rows[i]["urn"].ObjToDouble();
                        //if ( urnCost > 0D)
                        //    dx.Rows[i]["casketCost"] = casketCost + urnCost;
                        status = CompareValues(dRows, "Column11", dx, i, "classa", status, "CL");
                        status = CompareValues(dRows, "Column9", dx, i, "cashAdvance", status, "CA");
                        status = CompareValues(dRows, "Column18", dx, i, "vaultCost", status, "VC");
                        status = CompareValues(dRows, "Column19", dx, i, "otherBonuses", status, "OB");
                        status = CompareValues(dRows, "Column17", dx, i, "casketCost", status, "CC");
                        status = CompareValues(dRows, "Column12", dx, i, "totalDiscount", status, "TD");
                        status = CompareValues(dRows, "Column10", dx, i, "currentPrice", status, "TF");
                        status = CompareValues(dRows, "Column20", dx, i, "netFuneral", status, "NF");
                        status = CompareValues(dRows, "Column6", dx, i, "currentServices", status, "SA");

                        if (!String.IsNullOrWhiteSpace(status))
                        {
                            if (badDt == null)
                            {
                                badDt = new DataTable();
                                badDt.Columns.Add("serviceId");
                                badDt.Columns.Add("BAD");
                                badDt.Columns.Add("casketdesc");
                                badDt.Columns.Add("cashAdvance", Type.GetType("System.Double"));
                                badDt.Columns.Add("timCashAdvance", Type.GetType("System.Double"));
                                badDt.Columns.Add("casketCost", Type.GetType("System.Double"));
                                badDt.Columns.Add("timCasketCost", Type.GetType("System.Double"));
                                badDt.Columns.Add("vaultCost", Type.GetType("System.Double"));
                                badDt.Columns.Add("timVaultCost", Type.GetType("System.Double"));
                                badDt.Columns.Add("otherBonus", Type.GetType("System.Double"));
                                badDt.Columns.Add("timOtherBonus", Type.GetType("System.Double"));
                                badDt.Columns.Add("totalDiscount", Type.GetType("System.Double"));
                                badDt.Columns.Add("timTotalDiscount", Type.GetType("System.Double"));
                                badDt.Columns.Add("timServiceAmount", Type.GetType("System.Double"));
                                badDt.Columns.Add("serviceAmount", Type.GetType("System.Double"));
                            }

                            dRow = badDt.NewRow();
                            dRow["BAD"] = status;
                            dRow["serviceId"] = dx.Rows[i]["serviceId"].ObjToString();
                            dRow["cashAdvance"] = dx.Rows[i]["cashAdvance"].ObjToDouble();
                            dRow["timCashAdvance"] = dRows[0]["column9"].ObjToDouble();
                            dRow["casketCost"] = dx.Rows[i]["casketCost"].ObjToDouble();
                            dRow["timCasketCost"] = dRows[0]["column17"].ObjToDouble();
                            dRow["vaultCost"] = dx.Rows[i]["vaultCost"].ObjToDouble();
                            dRow["timVaultCost"] = dRows[0]["column18"].ObjToDouble();
                            dRow["otherBonus"] = dx.Rows[i]["otherBonuses"].ObjToDouble();
                            dRow["timOtherBonus"] = dRows[0]["column19"].ObjToDouble();
                            dRow["totalDiscount"] = dx.Rows[i]["totalDiscount"].ObjToDouble();
                            dRow["timTotalDiscount"] = dRows[0]["column12"].ObjToDouble();
                            dRow["casketdesc"] = dx.Rows[i]["casketdesc"].ObjToString();
                            dRow["timServiceAmount"] = dRows[0]["column6"].ObjToDouble();
                            dRow["serviceAmount"] = dx.Rows[i]["currentServices"].ObjToString();
                            badDt.Rows.Add(dRow);
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    dx.Rows[i]["bad"] = status;
                }
                else
                    dx.Rows[i]["bad"] = "BAD ID";
            }

            G1.NumberDataTable(badDt);

            gridMain.Columns["bad"].Visible = true;
            gridMain.RefreshEditor(true);
            gridMain.PostEditor();

            dgv.DataSource = dx;
            dgv.Refresh();

            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private string CompareValues(DataRow[] dRows, string timCol, DataTable dx, int i, string myCol, string status, string abrev )
        {
            double timValue = dRows[0][timCol].ObjToDouble();
            double myValue = dx.Rows[i][myCol].ObjToDouble();
            myValue = G1.RoundValue(myValue);
            if ( myValue != timValue )
            {
                //if ( String.IsNullOrWhiteSpace ( status ))
                //    status = "BAD ";
                status += abrev + ",";
            }

            return status;
        }
        /***********************************************************************************************/
        private void gridMain_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            DataTable dt = (DataTable)dgv.DataSource;
            if (e.Column.FieldName.ToUpper() == "casketCost")
            {
                string record = dr["record"].ObjToString();
                string casketCost = dr["casketCost"].ObjToString();
                casketCost = casketCost.Replace(",", "");
                G1.update_db_table("fcust_extended", "record", record, new string[] { "casketCost", casketCost });
            }
            else if (e.Column.FieldName.ToUpper() == "vaultCost")
            {
                string record = dr["record"].ObjToString();
                string vaultCost = dr["vaultCost"].ObjToString();
                vaultCost = vaultCost.Replace(",", "");
                G1.update_db_table("fcust_extended", "record", record, new string[] { "vaultCost", vaultCost });
            }
        }
        /***********************************************************************************************/
        private void btnZeros_Click(object sender, EventArgs e)
        {
            DataTable dx = (DataTable)dgv.DataSource;
            ShowZeros(dx);
        }
        /***********************************************************************************************/
        private DataTable LoadFuneralGroups ( DataTable dt )
        {
            string cmd = "Select * from `funeral_groups` ORDER by `order`;";
            DataTable gDt = G1.get_db_data(cmd);
            if (gDt.Rows.Count <= 0)
                return dt;

            string groupName = "";
            DataTable dx = null;
            string serviceRecord = "";
            string service1 = "";
            DataRow[] dRows = null;
            double price = 0D;

            for ( int i=0; i<gDt.Rows.Count; i++)
            {
                groupName = gDt.Rows[i]["shortname"].ObjToString();
                if (String.IsNullOrWhiteSpace(groupName))
                    continue;
                cmd = "Select * from `funeral_gplgroups` g LEFT JOIN `services` s on g.`service` = s.`service` where g.`groupname` = '" + groupName + "' order by g.`record`;";

                //cmd = "Select * from `packages` where `groupname` = '" + groupName + "' AND `PackageName` = 'master';";
                dx = G1.get_db_data(cmd);
                if (dx.Rows.Count <= 0)
                    continue;
                G1.AddNewColumn(gridMain2, groupName, groupName, "N2", FormatType.Numeric, 80, true);
                gridMain2.Columns[groupName].OptionsColumn.FixedWidth = true;
                dt.Columns.Add(groupName, Type.GetType("System.Double"));

                for ( int j=0; j<dx.Rows.Count; j++)
                {
                    service1 = dx.Rows[j]["service1"].ObjToString();
                    if (String.IsNullOrWhiteSpace(service1))
                        continue;
                    dRows = dt.Select("service='" + service1 + "'");
                    if ( dRows.Length > 0 )
                    {
                        price = dx.Rows[j]["price"].ObjToDouble();
                        dRows[0][groupName] = price;
                    }
                }
            }

            G1.ClearAllPositions(gridMain2);
            string name = "";
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                name = dt.Columns[i].ColumnName.Trim();
                if (name.ToUpper().IndexOf("AS") == 0)
                    gridMain2.Columns[i].Visible = false;
                else
                    G1.SetColumnPosition(gridMain2, dt.Columns[i].ColumnName, i);
            }

            return dt;
        }
        /***********************************************************************************************/
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            TabPage current = (sender as TabControl).SelectedTab;
            if (current.Name.Trim().ToUpper() == "TABMAIN")
                return;

            if (current.Name.Trim().ToUpper() == "TABSERVICES")
            {
                if ( dgv2.DataSource == null )
                {
                    this.Cursor = Cursors.WaitCursor;
                    string cmd = "Select * from `funeral_master`;";
                    DataTable dt = G1.get_db_data(cmd);

                    //LoadFuneralGroups(dt);

                    G1.NumberDataTable(dt);
                    dgv2.DataSource = dt;
                    this.Cursor = Cursors.Default;
                }
            }
            else if (current.Name.Trim().ToUpper() == "TABINVENTORY")
            {
                if (dgv3.DataSource == null)
                {
                    this.Cursor = Cursors.WaitCursor;
                    string cmd = "Select * from `inventorylist` ORDER by `order`;";
                    DataTable dt = G1.get_db_data(cmd);
                    dt.Columns.Add("num");
                    dt.Columns.Add("masterCost", Type.GetType("System.Double"));

                    cmd = "Select * from `casket_master`;";
                    DataTable dx = G1.get_db_data(cmd);
                    DataRow[] dRows = null;
                    string casketCode = "";
                    double cost = 0D;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dt.Rows[i]["casketguage"] = dt.Rows[i]["casketguage"].ObjToString().Trim();
                        casketCode = dt.Rows[i]["casketcode"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(casketCode))
                        {
                            dRows = dx.Select("casketcode='" + casketCode + "'");
                            if ( dRows.Length > 0 )
                            {
                                cost = dRows[0]["casketcost"].ObjToDouble();
                                dt.Rows[i]["masterCost"] = cost;
                            }
                        }
                    }


                    G1.NumberDataTable(dt);
                    dgv3.DataSource = dt;
                    this.Cursor = Cursors.Default;
                }
            }
            else if (current.Name.Trim().ToUpper() == "TABDIFFERENCE")
            {
                if ( badDt != null )
                {
                    dgv4.DataSource = badDt;
                    dgv4.Refresh();
                }
            }
        }
        /***********************************************************************************************/
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            G1.SpyGlass(gridMain3);
        }
        /***********************************************************************************************/
        private void btnGetTim_Click(object sender, EventArgs e)
        {
            double cost = 0D;
            string str = "";
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string file = ofd.FileName;
                    DataTable dt = null;

                    this.Cursor = Cursors.WaitCursor;
                    dt = ExcelWriter.ReadFile2(file, 0, "Sheet1");
                    if (dt.Rows.Count > 0)
                    {
                        try
                        {
                            dt.Columns["Column26"].ColumnName = "casketCode";
                            dt.Columns["Column27"].ColumnName = "casketDescription";
                            dt.Columns.Add("casketCost", Type.GetType("System.Double"));
                            dt.Columns.Add("Status");

                            dt.Rows.RemoveAt(0);

                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                str = dt.Rows[i]["Column28"].ObjToString();
                                cost = str.ObjToDouble();
                                dt.Rows[i]["casketCost"] = cost;
                            }

                            for (int i = (dt.Rows.Count - 1); i >= 0; i--)
                            {
                                str = dt.Rows[i]["casketCode"].ObjToString();
                                if (String.IsNullOrWhiteSpace(str))
                                {
                                    str = dt.Rows[i]["casketDescription"].ObjToString();
                                    if (String.IsNullOrWhiteSpace(str))
                                        dt.Rows.RemoveAt(i);
                                }
                            }
                        }
                        catch ( Exception ex)
                        {
                        }

                        try
                        {
                            string casketCode = "";
                            DataRow[] dRows = null;

                            DataTable dx = (DataTable)dgv3.DataSource;
                            if (G1.get_column_number(dx, "timCost") < 0)
                                dx.Columns.Add("timCost", Type.GetType("System.Double"));

                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                cost = dt.Rows[i]["casketCost"].ObjToDouble();
                                casketCode = dt.Rows[i]["casketCode"].ObjToString();
                                if ( casketCode == "O39")
                                {
                                }
                                if (!String.IsNullOrWhiteSpace(casketCode))
                                {
                                    dRows = dx.Select("casketcode='" + casketCode + "'");
                                    if (dRows.Length > 0)
                                        dRows[0]["timCost"] = cost;
                                }
                            }
                            dgv3.DataSource = dx;
                            dgv3.RefreshDataSource();
                            gridMain3.RefreshEditor(true);
                            dgv3.Refresh();
                        }
                        catch ( Exception ex)
                        {
                        }

                    }
                    this.Cursor = Cursors.Default;
                }
            }
        }
        /***********************************************************************************************/
        private void gridMain4_DoubleClick(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            DataRow dr = gridMain4.GetFocusedDataRow();

            string serviceId = dr["serviceId"].ObjToString();
            DataRow[] dRows = dt.Select("serviceId='" + serviceId + "'");
            if (dRows.Length <= 0)
                return;

            string contract = dRows[0]["contractNumber"].ObjToString();

            if (!String.IsNullOrWhiteSpace(contract))
            {
                this.Cursor = Cursors.WaitCursor;

                Form form = G1.IsFormOpen("EditCust", contract);
                if (form != null)
                {
                    form.Show();
                    form.WindowState = FormWindowState.Maximized;
                    form.Visible = true;
                    form.BringToFront();
                }
                else
                {
                    EditCust custForm = new EditCust(contract);
                    custForm.Tag = contract;
                    //custForm.custClosing += CustForm_custClosing;
                    custForm.Show();
                }
                //FunPayments editFunPayments = new FunPayments(null, contract, "", false, true );
                //editFunPayments.Show();

                //CustomerDetails clientForm = new CustomerDetails(contract);
                //clientForm.Show();
                this.Cursor = Cursors.Default;
            }
        }
        /***********************************************************************************************/
        private void gridMain_CustomColumnDisplayText_1(object sender, CustomColumnDisplayTextEventArgs e)
        {
            ColumnView view = sender as ColumnView;

            string columnName = e.Column.FieldName.ToUpper();

            if (e.Column.FieldName.ToUpper() == "SERVICEID" && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                string str = e.DisplayText.Trim();
                if (!String.IsNullOrWhiteSpace(str))
                    e.DisplayText = CleanServiceId(str);
            }
        }
        /***********************************************************************************************/
    }
}
