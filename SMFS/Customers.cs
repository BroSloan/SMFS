using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

using GeneralLib;
using MySql.Data.MySqlClient;
using MySql.Data.Types;

using DevExpress.XtraPrinting;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.Utils.DPI;
using DevExpress.CodeParser;
using System.Web.UI.WebControls;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.Utils.Menu;
using DevExpress.XtraGrid.Localization;
/***********************************************************************************************/

// Adam Sloan

/***********************************************************************************************/
namespace SMFS
{
    /***********************************************************************************************/
    public partial class Customers : Form
    {
        /***********************************************************************************************/
        private bool selecting = false;
        private bool loading = true;
        private string workType = "";
        private string LapsedDate = "";
        private bool modTrusts = false;
        private bool modInsurance = false;
        private bool testing = false;
        private string oldWhat = "";
        /***********************************************************************************************/
        public Customers(bool select = false)
        {
            workType = "";
            selecting = select;
            InitializeComponent();
            btnDOLP.Hide();
        }
        /***********************************************************************************************/
        public Customers(string type, bool orphans = false )
        {
            workType = type;
            InitializeComponent();
            btnDOLP.Hide();
            if ( orphans )
            {
                cmbType.Text = "Insurance";
                cmbQualify.Text = "Orphans";
            }
        }
        /***********************************************************************************************/
        private string customersFile = "customers";
        private string contractsFile = "contracts";
        /***********************************************************************************************/
        private void LoadData()
        {
            string what = cmbType.Text.Trim().ToUpper();
            if ( what.Trim().ToUpper() == "PAYERS")
            {
                G1.CleanupDataGrid(ref dgv2);
                GC.Collect();
                dgv2.Show();
                dgv.Hide();
                LoadPayers();
                return;
            }
            G1.CleanupDataGrid(ref dgv);
            GC.Collect();
            dgv2.Hide();
            dgv.Show();
            loading = true;
            G1.loadGroupCombo(cmbSelectColumns, "Customers", "Primary");
            txtThreshold.Text = "0.00";
            chkMismatches.Visible = false;
            chkMismatchDates.Visible = false;
            gridMain.Columns["cbal"].Visible = false;
            gridMain.Columns["newduedate"].Visible = false;
            gridMain.Columns["days"].Visible = false;
            gridMain.Columns["creditBalance"].Visible = false;
            gridMain.Columns["totalInterest"].Visible = false;
            gridMain.Columns["cint"].Visible = false;
            gridMain.Columns["lapseDate8"].Visible = false;
            gridMain.Columns["ServiceId1"].Visible = true;
            gridMain.Columns["deceasedDate"].Visible = false;
            gridMain.Columns["payer"].Visible = false;
            barImport.Visible = false;
            barImport.Hide();
            if (!LoginForm.administrator)
                btnRecalc.Hide();
            labelValue.Hide();
            labValue.Hide();
            this.Cursor = Cursors.WaitCursor;
            SetupVisibleColumns();

            string qualify = cmbQualify.Text.Trim().ToUpper();

            what = cmbType.Text.Trim().ToUpper();
            if ( !String.IsNullOrWhiteSpace (workType))
            {
                if (workType.ToUpper() == "INSURANCE")
                {
                    what = "INSURANCE";
                    cmbType.Text = what;
                }
                workType = "";
            }

            customersFile = "customers";
            contractsFile = "contracts";
            if ( what == "INSURANCE")
            {
                customersFile = "icustomers";
                contractsFile = "icontracts";
                gridMain.Columns["payer"].Visible = true;
                gridMain.Columns["amtOfMonthlyPayt"].Visible = true;
            }
            else if ( what == "FUNERALS")
            {
                customersFile = "fcustomers";
                contractsFile = "fcontracts";
                gridMain.Columns["payer"].Visible = false;
                gridMain.Columns["amtOfMonthlyPayt"].Visible = false;
                gridMain.Columns["ServiceId1"].Visible = true;
            }

            string cmd = "Select * from `" + customersFile + "` p JOIN `" + contractsFile + "` d ON p.`contractNumber` = d.`contractNumber` ";
            if (what == "INSURANCE" || what == "PAYERS" )
            {
                cmd += " JOIN `payers` j on p.`payer` = j.`payer` ";
            }

            cmd += " WHERE ";
            if (what == "TRUSTS" || what == "FUNERALS")
            {
                this.Text = "Trust Customers";
                if (what == "FUNERALS")
                    this.Text = "Funeral Customers";

                cmd += " p.`coverageType` <> 'ZZ' ";
                if (qualify == "ACTIVE")
                {
                    cmd += " AND p.`lapsed` <> 'Y' AND d.`lapsed` <> 'Y' ";
                    if ( what == "FUNERALS")
                        cmd += " AND d.`deceasedDate` >= '18001231' ";
                    else
                        cmd += " AND d.`deceasedDate` < '19101231' ";
                    this.Text = "Trust Customers (Active)";
                    if (what == "FUNERALS")
                        this.Text = "Funeral Customers";
                }
                else if (qualify == "LAPSED")
                {
                    cmd += " AND (p.`lapsed` = 'Y' or d.`lapsed` = 'Y') ";
                    this.Text = "Trust Customers (Lapsed)";
                    if ( what == "FUNERALS")
                        this.Text = "Funeral Customers (Lapsed)";
                }
                else if (qualify == "DECEASED")
                {
//                    cmd += " AND d.`deceasedDate` > '19101231' ";
                    cmd += " AND d.`deceasedDate` > '18001231' ";
                    this.Text = "Trust Customers (Deceased)";
                    if (what == "FUNERALS")
                        this.Text = "Funeral Customers (Deceased)";
                }

                if (SMFS.activeSystem.ToUpper() == "RILES")
                {
                    cmd += " AND p.`contractNumber` LIKE 'RF%' ";
                    this.Text = "Riles " + this.Text;
                }
                else 
                {
                    cmd += " AND p.`contractNumber` NOT LIKE 'RF%' ";
                }

            }
            else if (what == "SINGLE PREMIUM")
            {
                cmd += " p.`coverageType` <> 'ZZ' ";
                cmd += " AND d.`downpayment` >= (`serviceTotal`+`merchandiseTotal`-`allowInsurance` - `allowMerchandise`) AND `serviceTotal` > '0' AND p.`coverageType` <> 'ZZ' ";
                if (qualify == "ACTIVE")
                {
                    cmd += " AND p.`lapsed` <> 'Y' AND d.`lapsed` <> 'Y' ";
                    cmd += " AND d.`deceasedDate` < '19101231' ";
                }
                else if (qualify == "LAPSED")
                    cmd += " AND (p.`lapsed` = 'Y' or d.`lapsed` = 'Y') ";
                else if (qualify == "DECEASED")
                    cmd += " AND d.`deceasedDate` > '18101231' ";
            }
            else if (what == "INSURANCE")
            {
                this.Text = "Insurance Customers";
                if (qualify == "ACTIVE")
                {
                    if ( G1.isField () )
                        cmd += " p.`contractNumber` LIKE 'ZZ%' ";
                    else
                        cmd += " (p.`contractNumber` LIKE 'ZZ%' OR p.`contractNumber` LIKE 'OO%' OR p.`contractNumber` LIKE 'MM%') ";
//                    cmd += " p.`coverageType` = 'ZZ' ";
                    cmd += " AND p.`lapsed` <> 'Y' AND d.`lapsed` <> 'Y' ";
                    cmd += " AND d.`deceasedDate` < '18101231' ";
                    this.Text = "Insurance Customers (Active)";
                }
                else if (qualify == "LAPSED")
                {
                    cmd += " (p.`contractNumber` LIKE 'ZZ%' OR p.`contractNumber` LIKE 'OO%' OR p.`contractNumber` LIKE 'MM%') ";
//                    cmd += " p.`coverageType` = 'ZZ' ";
                    cmd += " AND (p.`lapsed` = 'Y' or d.`lapsed` = 'Y') ";
                    this.Text = "Insurance Customers (Lapsed)";
                }
                else if (qualify == "DECEASED")
                {
                    cmd += " (p.`contractNumber` LIKE 'ZZ%' OR p.`contractNumber` LIKE 'OO%' OR p.`contractNumber` LIKE 'MM%') ";
//                    cmd += " p.`coverageType` = 'ZZ' ";
                    cmd += " AND d.`deceasedDate` > '18101231' ";
                    this.Text = "Insurance Customers (Deceased)";
                }
                else if (qualify == "ORPHANS")
                {
                    cmd += " (p.`contractNumber` LIKE 'OO%' or p.`contractNumber` LIKE 'MM%') ";
                    this.Text = "Insurance Customers (Orphans)";
                }
            }

            if (qualify == "LAPSED")
                gridMain.Columns["lapseDate8"].Visible = true;
            //if (qualify == "DECEASED")
            //    gridMain.Columns["deceasedDate"].Visible = true;
            
            if (G1.RobbyServer)
            {
                if (testing)
                    cmd += " LIMIT 300 ";
            }

            cmd += ";";
            cmd = cmd.Replace("WHERE ;", ";");


            DataTable dt = G1.get_db_data(cmd);

            //            DataRow[] dRow = dt.Select("contractNumber='P16050UI'");
            //            int len = dRow.Length;
            dt.Columns.Add("num");
            dt.Columns.Add("bDate");
            dt.Columns.Add("ssno");
            dt.Columns.Add("agreement");
            dt.Columns.Add("select");
            dt.Columns.Add("fullname");
            dt.Columns.Add("paid", Type.GetType("System.Double"));
            dt.Columns.Add("purchase", Type.GetType("System.Double"));
            dt.Columns.Add("dueDate");
            dt.Columns.Add("DOLP");
            dt.Columns.Add("cbal", Type.GetType("System.Double"));
            dt.Columns.Add("newduedate");
            dt.Columns.Add("days", Type.GetType("System.Int32"));
            dt.Columns.Add("cint", Type.GetType("System.Double"));
            dt.Columns.Add("financed", Type.GetType("System.Double"));
            dt.Columns.Add("trust85", Type.GetType("System.Double"));
            dt.Columns.Add("contractValue", Type.GetType("System.Double"));
            dt.Columns.Add("percentPaid", Type.GetType("System.Double"));
            dt.Columns.Add("idate");
            dt.Columns.Add("realDOLP");
            dt.Columns.Add("DDATE");
            dt.Columns.Add("mod");

            DateTime ddate = DateTime.Now;
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    ddate = dt.Rows[i]["dueDate8"].ObjToDateTime();
            //    dt.Rows[i]["dueDate"] = ddate.ToString("yyyy-MM-dd");
            //    ddate = dt.Rows[i]["issueDate8"].ObjToDateTime();
            //    dt.Rows[i]["idate"] = ddate.ToString("yyyy-MM-dd");
            //    ddate = dt.Rows[i]["lastDatePaid8"].ObjToDateTime();
            //    dt.Rows[i]["realDOLP"] = ddate.ToString("yyyy-MM-dd");
            //}


            labelValue.Hide();
            labValue.Hide();
            if (qualify == "LAPSED")
            {
                gridMain.Columns["lapseDate8"].Visible = true;
                gridMain.Columns["DOLP"].Visible = true;
            }
            else
            {
                gridMain.Columns["lapseDate8"].Visible = false;
                gridMain.Columns["DOLP"].Visible = false;
            }

            if (chkIncludePaid.Checked)
            {
                gridMain.Columns["paid"].Visible = true;
                gridMain.Columns["purchase"].Visible = true;
            }
            else
            {
                gridMain.Columns["paid"].Visible = false;
                gridMain.Columns["purchase"].Visible = false;
            }

            bool showOrNot = true;
            if (what == "INSURANCE")
                showOrNot = false;

            gridMain.Columns["trust85"].Visible = showOrNot;
            gridMain.Columns["cbal"].Visible = showOrNot;
            gridMain.Columns["purchase"].Visible = showOrNot;
            gridMain.Columns["paid"].Visible = showOrNot;
            gridMain.Columns["financed"].Visible = showOrNot;
            gridMain.Columns["newduedate"].Visible = showOrNot;
            gridMain.Columns["days"].Visible = showOrNot;
            gridMain.Columns["apr"].Visible = showOrNot;
            gridMain.Columns["percentPaid"].Visible = showOrNot;
            gridMain.Columns["contractValue"].Visible = showOrNot;
            gridMain.Columns["creditBalance"].Visible = showOrNot;
            gridMain.Columns["extraItemAmtMI1"].Visible = showOrNot;
            gridMain.Columns["extraItemAmtMI2"].Visible = showOrNot;
            gridMain.Columns["agreement"].Visible = showOrNot;
            gridMain.Columns["trustRemoved"].Visible = showOrNot;

            gridMain.Columns["cbal"].Visible = false;
            gridMain.Columns["newduedate"].Visible = false;
            gridMain.Columns["days"].Visible = false;
            gridMain.Columns["creditBalance"].Visible = false;
            gridMain.Columns["totalInterest"].Visible = false;
            gridMain.Columns["cint"].Visible = false;
            if (barImport.Visible)
            {
                gridMain.Columns["cbal"].Visible = true;
                gridMain.Columns["newduedate"].Visible = true;
                gridMain.Columns["days"].Visible = true;
                gridMain.Columns["creditBalance"].Visible = true;
                gridMain.Columns["totalInterest"].Visible = true;
                gridMain.Columns["cint"].Visible = true;
            }


            G1.NumberDataTable(dt);
            //FixDates(dt, "birthDate", "bDate");
            FormatSSN(dt, "ssn", "ssno");
            SetupFullNames(dt);
            SetupAgreementIcon(dt);
            if (chkIncludePaid.Checked)
            {
                if ( qualify != "ALL")
                {
                    labelValue.Show();
                    labValue.Show();
                    CalcPaid(dt);
                }
            }

            if ( what == "TRUSTS" && qualify == "ACTIVE")
            {
                gridMain.Columns["trustRemoved"].Visible = false;
            }

            if (what == "TRUSTS")
            {
                GetFinancedAmount(dt);
                CalcPaid(dt);
            }
            if (gridMain.Columns["DOLP"].Visible == true)
                DetermineLapsed(dt);
            FixDeceasedDate(dt);
            if (selecting)
                gridMain.Columns["select"].Visible = true;
            dgv.DataSource = dt;
            gridBand2.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gridMain.Columns["num"].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gridMain.Columns["contractNumber"].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            LoadPreviousContracts();

            SetFieldUserFormat();

            string type = cmbType.Text.ToUpper();
            if (type == "TRUSTS" || type == "FUNERALS" || type == "SINGLE PREMIUM")
            {
                dgv.DataSource = dt;
                dgv2.Hide();
                dgv.Show();
            }
            else
            {
                LoadPremiums(dt);
                dgv2.DataSource = dt;
                dgv.Hide();
                dgv2.Show();
            }

            this.Cursor = Cursors.Default;
            loading = false;
        }
        /***********************************************************************************************/
        private void LoadPremiums ( DataTable dt )
        {
            string payer = "";
            double premium = 0D;
            double premium2 = 0D;
            double monthlyPremium = 0D;
            double historicPremium = 0D;
            double monthlySecNat = 0D;
            double monthly3rdParty = 0D;
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                payer = dt.Rows[i]["payer"].ObjToString();
                if ( !String.IsNullOrWhiteSpace ( payer ))
                {
                    if ( payer == "BB-0472T")
                    {
                    }
                    premium = dt.Rows[i]["amtOfMonthlyPayt"].ObjToDouble();
                    premium2 = dt.Rows[i]["amtOfMonthlyPayt1"].ObjToDouble();
                    if (premium != premium2)
                    {
                        premium = Policies.CalcMonthlyPremium(payer);
                        dt.Rows[i]["amtOfMonthlyPayt"] = premium;
                    }
                    else
                    {
                        premium = Policies.CalcMonthlyPremium(payer, DateTime.Now);

                        //CustomerDetails.CalcMonthlyPremium ( payer, ref monthlyPremium, ref historicPremium, ref monthlySecNat, ref monthly3rdParty );

                        dt.Rows[i]["amtOfMonthlyPayt"] = premium;
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void LoadPayers ()
        {
            this.Cursor = Cursors.WaitCursor;
            string cmd = "Select * from `payers` p LEFT JOIN `icustomers` x ON p.`contractNumber` = x.`contractNumber` ";
            string type = cmbQualify.Text.ToUpper();
            if (type == "ACTIVE")
                cmd += " WHERE p.`deceasedDate` < '100-01-01' AND `lapseDate8` < '100-01-01' AND p.`lapsed` <> 'Y' ";
            else if ( type == "LAPSED")
                cmd += " WHERE `lapseDate8` > '100-01-01' OR p.`lapsed` = 'Y' ";
            else if (type == "DECEASED")
                cmd += " WHERE p.`deceasedDate` > '100-01-01' ";
            cmd += ";";

            DataTable dx = G1.get_db_data(cmd);
            G1.NumberDataTable(dx);
            dgv2.DataSource = dx;
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void LoadPreviousContracts ()
        {
            string contract = "";
            string cmd = "Select * from `previouscustomers` where `user` = '" + LoginForm.username + "' order by `tmstamp` DESC;";
            DataTable dt = G1.get_db_data(cmd);
            ToolStripMenuItem menu = this.previousContractsToolStripMenuItem;
            for (int i = (menu.DropDownItems.Count-1); i >= 0; i--)
            {
                ToolStripMenuItem sMenu = (ToolStripMenuItem) (menu.DropDownItems[i]);
                menu.DropDownItems.RemoveAt(i);
                sMenu.Dispose();
            }
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                contract = dt.Rows[i]["contractNumber"].ObjToString();
                ToolStripMenuItem sMenu = new ToolStripMenuItem();
                sMenu.Name = contract;
                sMenu.Text = contract;
                sMenu.Click += SMenu_Click;
                menu.DropDownItems.Add(sMenu);
            }
        }
        /***********************************************************************************************/
        public static void DetermineIfInsurance ( string contractNumber, ref string contractsFile, ref string customersFile )
        {
            contractsFile = "contracts";
            customersFile = "customers";
            if ( contractNumber.Trim().ToUpper().IndexOf( "ZZ") == 0 )
            {
                contractsFile = "icontracts";
                customersFile = "icustomers";
            }
        }
        /***********************************************************************************************/
        private void SMenu_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem sMenu = (ToolStripMenuItem)sender;
            string contract = sMenu.Name.ObjToString();
            string contractFile = "";
            string customerFile = "";
            DetermineIfInsurance(contract, ref contractFile, ref customerFile);
            string cmd = "Select * from `" + contractFile + "` where `contractNumber` = '" + contract + "';";
            DataTable dt = G1.get_db_data(cmd);
            if ( dt.Rows.Count > 0 )
            {
                this.Cursor = Cursors.WaitCursor;
                CustomerDetails clientForm = new CustomerDetails(contract);
                clientForm.Show();
                this.Cursor = Cursors.Default;
            }
        }
        /***********************************************************************************************/
        private void GetFinancedAmount( DataTable dt)
        {
            double financed = 0D;
            double contractValue = 0D;
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                financed = DailyHistory.GetFinanceValue(dt.Rows[i]);
                dt.Rows[i]["financed"] = financed;
                contractValue = DailyHistory.GetContractValue(dt.Rows[i]);
                dt.Rows[i]["contractValue"] = contractValue;
            }
        }
        /***********************************************************************************************/
        private void FixDeceasedDate(DataTable dt)
        {
            string date1 = "";
            string date2 = "";
            if (G1.get_column_number(dt, "DDATE") < 0)
                dt.Columns.Add("DDATE");
            if (G1.get_column_number(dt, "deceasedDate") < 0)
                return;
            if (G1.get_column_number(dt, "deceasedDate1") < 0)
                return;
            DateTime date = DateTime.Now;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                date1 = dt.Rows[i]["deceasedDate"].ObjToString();
                date = date1.ObjToDateTime();
                if (date.Year > 100)
                    dt.Rows[i]["DDATE"] = date.ToString("yyyy-MM-dd");
                //if (date1.IndexOf("0000") >= 0)
                //{
                //    date2 = dt.Rows[i]["deceasedDate1"].ObjToString();
                //    if (date2.IndexOf("0000") < 0)
                //        dt.Rows[i]["deceasedDate"] = dt.Rows[i]["deceasedDate1"];
                //}
            }
        }
        /***********************************************************************************************/
        private void DetermineLapsed(DataTable dt)
        {
            if (G1.get_column_number(dt, "lapsed1") < 0)
                return;
            string lapse = "";
            string lapse1 = "";
            string cnum = "";
            string cmd = "";
            DataTable dx = null;
            DateTime dolp = DateTime.Now;
            string paymentsFile = "payments";
            if ( contractsFile == "INSURANCE")
                paymentsFile = "ipayments";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                lapse = dt.Rows[i]["lapsed"].ObjToString();
                lapse1 = dt.Rows[i]["lapsed1"].ObjToString();
                if (String.IsNullOrWhiteSpace(lapse))
                    lapse = " ";
                lapse += lapse1;
                dt.Rows[i]["lapsed"] = lapse;
                //if (!String.IsNullOrWhiteSpace(lapse))
                //{
                //    cnum = dt.Rows[i]["contractNumber"].ObjToString();
                //    if (!String.IsNullOrWhiteSpace(cnum))
                //    {
                //        cmd = "Select * from `" + paymentsFile + "` where `contractNumber` = '" + cnum + "' order by `payDate8` DESC LIMIT 1";
                //        dx = G1.get_db_data(cmd);
                //        if (dx.Rows.Count > 0)
                //        {
                //            dolp = dx.Rows[0]["payDate8"].ObjToDateTime();
                //            dt.Rows[i]["DOLP"] = dolp.ToString("MM/dd/yyyy");
                //        }
                //    }
                //}
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
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                serviceTotal = dt.Rows[i]["serviceTotal"].ObjToDouble();
                merchandiseTotal = dt.Rows[i]["merchandiseTotal"].ObjToDouble();
                balanceDue = dt.Rows[i]["balanceDue"].ObjToDouble();
                totalPurchase = serviceTotal + merchandiseTotal;
                contractValue = dt.Rows[i]["contractValue"].ObjToDouble();
//                paid = totalPurchase - balanceDue;
                paid = contractValue - balanceDue;
                dt.Rows[i]["paid"] = paid;
                dt.Rows[i]["purchase"] = totalPurchase;
                totalPaid += paid;
            }
            labValue.Text = "$" + G1.ReformatMoney(totalPaid);
        }
        /***********************************************************************************************/
        private void SetupFullNames(DataTable dt)
        {
            string fullname = "";
            string fname = "";
            string lname = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                fname = dt.Rows[i]["firstName"].ObjToString();
                lname = dt.Rows[i]["lastName"].ObjToString();
                fullname = fname + " " + lname;
                dt.Rows[i]["fullname"] = fullname;
            }
        }
        /***********************************************************************************************/
        private void SetupVisibleColumns()
        {
            ToolStripMenuItem menu = this.columnsToolStripMenuItem;
            for (int i = 0; i < gridMain.Columns.Count; i++)
            {
                string name = gridMain.Columns[i].Name;
                string caption = gridMain.Columns[i].Caption;
                ToolStripMenuItem nmenu = new ToolStripMenuItem();
                nmenu.Name = name;
                nmenu.Text = caption;
                nmenu.Checked = true;
                nmenu.Click += new EventHandler(nmenu_Click);
                menu.DropDownItems.Add(nmenu);
            }
        }
        /***********************************************************************************************/
        void nmenu_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem menu = (ToolStripMenuItem)sender;
            string name = menu.Name;
            int index = getGridColumnIndex(name);
            if (index < 0)
                return;
            if (menu.Checked)
            {
                menu.Checked = false;
                gridMain.Columns[index].Visible = false;
            }
            else
            {
                menu.Checked = true;
                gridMain.Columns[index].Visible = true;
            }
            gridMain.RefreshData();
            dgv.Refresh();
            ToolStripMenuItem xmenu = this.columnsToolStripMenuItem;
            xmenu.ShowDropDown();
        }
        /***********************************************************************************************/
        private int getGridColumnIndex(string columnName)
        {
            int index = -1;
            for (int i = 0; i < gridMain.Columns.Count; i++)
            {
                string name = gridMain.Columns[i].Name;
                if (name == columnName)
                {
                    index = i;
                    break;
                }
            }
            return index;
        }
        /***********************************************************************************************/
        private void FormatSSN(DataTable dt, string columnName, string newColumn)
        {
            string ssn = "";
            string ssno = "";
            int row = 0;
            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    row = i;
                    ssn = dt.Rows[i][columnName].ObjToString().Trim();
                    ssn = ssn.Replace("-", "");
                    ssno = ssn;
                    if (ssn.Trim().Length >= 8)
                        try { ssno = "XXX-XX-" + ssn.Substring(5, 4); }
                        catch { }
                    dt.Rows[i][newColumn] = ssno;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("***ERROR*** Fixing Date Field " + columnName + " Row= " + row + "SSN= " + ssn + " " + ex.Message.ObjToString());
            }
        }
        /***********************************************************************************************/
        private void FixDates(DataTable dt, string columnName, string newColumn)
        {
            string date = "";
            long ldate = 0L;
            int row = 0;
            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    row = i;
                    date = dt.Rows[i][columnName].ObjToString();
                    if (String.IsNullOrWhiteSpace(date))
                        continue;
                    if (date == "0000-00-00")
                    {
                        date = "";
                        dt.Rows[i][columnName] = date;
                    }
                    else
                    {
                        ldate = G1.date_to_days(date);
                        date = G1.days_to_date(ldate);
                        dt.Rows[i][newColumn] = date;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("***ERROR*** Fixing Date Field " + columnName + " Row= " + row + "Date= " + date + " " + ex.Message.ObjToString());
            }
        }
        /***********************************************************************************************/
        private void Customers_Load(object sender, EventArgs e)
        {
            btnSave.Hide ();

            SetupEdit(false);

            if (!G1.isAdminOrSuper())
                chkFullEdit.Hide();

            labelValue.Hide();
            labValue.Hide();
            btnFixInt.Hide();

            importToolStripMenuItem.HideDropDown();
            importToolStripMenuItem.Enabled = false;
            importToolStripMenuItem.Visible = false;

            if (!G1.isAdmin())
                chkDownPayment.Hide();


            SetupMenuPreferences();

            string type = cmbType.Text.ToUpper();
            CleanUpMenu(type, false );
            LoadPreviousContracts();
            dgv2.Hide();
            dgv2.Dock = DockStyle.Fill;
            dgv.Show();
            dgv.Dock = DockStyle.Fill;
            //            LoadData();
        }
        /***********************************************************************************************/
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (dgv.Visible)
            {
                G1.ShowHideFindPanel(gridMain);
                //if (gridMain.OptionsFind.AlwaysVisible == true)
                //    gridMain.OptionsFind.AlwaysVisible = false;
                //else
                //    gridMain.OptionsFind.AlwaysVisible = true;
            }
            else if ( dgv2.Visible )
            {
                G1.ShowHideFindPanel(gridMain2);
                //if (gridMain2.OptionsFind.AlwaysVisible == true)
                //    gridMain2.OptionsFind.AlwaysVisible = false;
                //else
                //{
                //    gridMain2.OptionsFind.AlwaysVisible = true;
                //}
            }
        }
        /***********************************************************************************************/
        private void gridMain_DoubleClick(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            if (dr == null)
                return;
            string contract = dr["contractNumber"].ObjToString();
            if (!String.IsNullOrWhiteSpace(contract))
            {
                this.Cursor = Cursors.WaitCursor;
                DataTable dt = (DataTable)dgv.DataSource;
                string what = cmbType.Text.Trim().ToUpper();
                G1.UpdatePreviousCustomer(contract, LoginForm.username);
                bool insurance = false;
                if (contract.ToUpper().IndexOf("ZZ") == 0)
                    insurance = true;
                if (contract.ToUpper().IndexOf("MM") == 0)
                    insurance = true;
                if (contract.ToUpper().IndexOf("OO") == 0)
                    insurance = true;
                if (cmbQualify.Text.ToUpper() == "ORPHANS")
                    insurance = true;
                if ( insurance )
                {
                    string cmd = "Select * from `policies` p JOIN `icustomers` d ON p.`contractNumber` = d.`contractNumber` JOIN `icontracts` x ON p.`contractNumber` = x.`contractNumber` ";
                    cmd += " WHERE p.`contractNumber` = '" + contract + "' ";

                    DataTable dx = G1.get_db_data(cmd);
                    if (dx.Rows.Count > 0)
                    {
                        Policies policyForm = new Policies(contract);
                        policyForm.Show();
                    }
                    else
                    {
                        CustomerDetails clientForm = new CustomerDetails(contract);
                        clientForm.Show();
                    }
                }
                else
                {
                    if (what == "FUNERALS")
                    {
                        EditCust custForm = new EditCust(contract);
                        custForm.custClosing += CustForm_custClosing;
                        custForm.Show();
                    }
                    else
                    {
                        CustomerDetails clientForm = new CustomerDetails(contract);
                        clientForm.Show();
                    }
                }
                LoadPreviousContracts();
                this.Cursor = Cursors.Default;
                //DailyHistory dailyForm = new DailyHistory(contract);
                //dailyForm.Show();
            }
        }
        /***********************************************************************************************/
        private void CustForm_custClosing(string record, double amountFiled, double amountReceived)
        {
            if ( !String.IsNullOrWhiteSpace (record))
            {
                string cmd = "Select * from `fcust_extended` where `record` = '" + record + "';";
                DataTable dt = G1.get_db_data(cmd);
                if ( dt.Rows.Count > 0 )
                {
                    string contractNumber = dt.Rows[0]["contractNumber"].ObjToString();
                }
            }
        }
        /***********************************************************************************************/
        private void SetupAgreementIcon(DataTable dt)
        {
            DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit selectnew = this.repositoryItemCheckEdit1;
            selectnew.NullText = "";
            selectnew.ValueChecked = "1";
            selectnew.ValueUnchecked = "0";
            selectnew.ValueGrayed = "";
            string filename = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                filename = dt.Rows[i]["agreementFile"].ObjToString();
                if (!String.IsNullOrWhiteSpace(filename))
                    dt.Rows[i]["agreement"] = "1";
                else
                    dt.Rows[i]["agreement"] = "0";
                dt.Rows[i]["select"] = "0";
            }
        }
        /***********************************************************************************************/
        private void repositoryItemCheckEdit1_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)(dgv.DataSource);
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            string value = dr["agreement"].ObjToString();
            if (value == "1")
            {
                string filename = "";
                string title = "Agreement for (" + contract + ") ";
                string cmd = "Select * from `" + customersFile + "` where `contractNumber` = '" + contract + "';";
                DataTable dx = G1.get_db_data(cmd);
                if (dx.Rows.Count > 0)
                {
                    filename = dx.Rows[0]["agreementFile"].ObjToString();
                    string firstName = dx.Rows[0]["firstName"].ObjToString();
                    string lastName = dx.Rows[0]["lastName"].ObjToString();
                    title = "Agreement for (" + contract + ") " + firstName + " " + lastName;
                    if (!String.IsNullOrWhiteSpace(filename))
                    {
                        string record = dr["!imagesRecord"].ObjToString();
                        if (record != "-1")
                            ShowPDfImage(record, title, filename);
                    }
                }
            }
            else
            {
                string text = btnRefresh.Text.ToUpper();
                if ( text == "FIX")
                {
                    string cmd = "Select * from `pdfimages` where `contractNumber` = '" + contract + "';";
                    DataTable picDt = G1.get_db_data(cmd);
                    if (picDt.Rows.Count > 0)
                    {
                        string record = picDt.Rows[0]["record"].ObjToString();
                        if ( !String.IsNullOrWhiteSpace ( record ))
                        {
                            if ( record != "-1")
                            {
                                string title = "Agreement for (" + contract + ") ";
                                ShowPDfImage(record, title, "");
                            }
                        }
                    }
                }
            }
        }
        /***********************************************************************************************/
        public static void ShowPDfImage(string record, string title, string filename)
        {
            string command = "Select `image` from `pdfimages` where `Record` = '" + record + "';";
            MySqlCommand cmd1 = new MySqlCommand(command, G1.conn1);
            cmd1.Connection.Open();
            try
            {
                using (MySqlDataReader dr = cmd1.ExecuteReader(System.Data.CommandBehavior.Default))
                {
                    if (dr.Read())
                    {
                        byte[] fileData = (byte[])dr.GetValue(0);
                        string str = G1.ConvertToString(fileData);
                        if (str.IndexOf("rtf1") > 0)
                        {
                            ViewRTF aForm = new ViewRTF(str);
                            aForm.Show();
                            //ArrangementForms aForm = new ArrangementForms("Agreement", record, "", fileData);
                            //aForm.Show();
                        }
                        else
                        {
                            ViewPDF pdfForm1 = new ViewPDF(title, "", "", fileData);
                            pdfForm1.Show();
                        }
                    }
                    dr.Close();
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (cmd1.Connection.State == ConnectionState.Open)
                    cmd1.Connection.Close();
            }
        }
        /***********************************************************************************************/
        private void attachAgreementToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)(dgv.DataSource);
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            string record = dr["record"].ObjToString();
            using (OpenFileDialog ofdImage = new OpenFileDialog())
            {
                ofdImage.Multiselect = false;

                if (ofdImage.ShowDialog() == DialogResult.OK)
                {
                    string filename = ofdImage.FileName;
                    filename = filename.Replace('\\', '/');
                    if (!String.IsNullOrWhiteSpace(filename))
                    {
                        G1.update_db_table(customersFile, "record", record, new string[] { "agreementFile", filename });
                        string record1 = G1.create_record("pdfimages", "filename", "-1");
                        G1.update_db_table("pdfimages", "record", record1, new string[] { "filename", filename });
                        G1.ReadAndStorePDF("pdfimages", record1, filename);
                        G1.update_db_table(customersFile, "record", record, new string[] { "!imagesRecord", record1 });
                        dr["!imagesRecord"] = record1;
                        dr["agreement"] = "1";
                    }
                }
                dgv.Refresh();
                this.Refresh();
            }
        }
        /***********************************************************************************************/
        private void detachAgreementToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            string record = dr["record"].ObjToString();
            G1.update_db_table(customersFile, "record", record, new string[] { "agreementFile", "", "!imagesRecord", "-1" });
            dr["agreementFile"] = "";
            dr["!imagesRecord"] = "-1";
            dr["agreement"] = "0";
            dgv.RefreshDataSource();
        }
        /***********************************************************************************************/
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            string text = btnRefresh.Text;
            if (text.ToUpper() == "FIX")
                FixData();
            else if (text.ToUpper() == "CREATE CUSTOMERS")
                CreateTheData();
            else
                LoadData();
        }
        /***********************************************************************************************/
        private void gridMain_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName.ToUpper() == "NUM")
            {
                //string filter = gridMain.ActiveFilterString.ObjToString();
                //if ( !String.IsNullOrWhiteSpace ( filter ))
                //{
                //}
                if (e.RowHandle >= 0)
                {
                    string num = (e.RowHandle + 1).ToString();
                    e.DisplayText = num;
                }
            }
            else if (e.Column.FieldName.ToUpper() == "CBAL" && !chkShowTEB.Checked )
            {
                if (e.RowHandle >= 0)
                {
                    int row = gridMain.GetDataSourceRowIndex(e.RowHandle);
                    DataTable dt = (DataTable)dgv.DataSource;
                    double balance = dt.Rows[row]["balanceDue"].ObjToDouble();
                    //if (balance <= 0D)
                    //    return;
                    balance = G1.RoundValue(balance);
                    double cbal = dt.Rows[row]["cbal"].ObjToDouble();
                    cbal = G1.RoundValue(cbal);
                    if ( cbal != balance)
                    {
                        e.Appearance.BackColor = System.Drawing.Color.Red;
                        e.Appearance.ForeColor = System.Drawing.Color.Yellow;
                    }
                    else
                    {
                        e.Appearance.BackColor = System.Drawing.Color.Green;
                        e.Appearance.ForeColor = System.Drawing.Color.Yellow;
                    }
                }
            }
            else if (e.Column.FieldName.ToUpper() == "NEWDUEDATE")
            {
                if (!this.barImport.Visible)
                    return;
                int row = gridMain.GetDataSourceRowIndex(e.RowHandle);
                DataTable dt = (DataTable)dgv.DataSource;
                DateTime newDueDate = dt.Rows[row]["newduedate"].ObjToDateTime();
                DateTime dueDate = dt.Rows[row]["duedate8"].ObjToDateTime();
                if ( dueDate != newDueDate && dueDate.Year != 2039 )
                {
                    e.Appearance.BackColor = System.Drawing.Color.Red;
                    e.Appearance.ForeColor = System.Drawing.Color.Yellow;
                }
                else
                {
                    e.Appearance.BackColor = System.Drawing.Color.Green;
                    e.Appearance.ForeColor = System.Drawing.Color.Yellow;
                }
            }
            else if (e.Column.FieldName.ToUpper() == "LAPSED")
            {
                if (e.RowHandle >= 0)
                {
                    if (!String.IsNullOrWhiteSpace(e.DisplayText))
                    {
                        DataTable dt = (DataTable)dgv.DataSource;
                        int row = gridMain.GetDataSourceRowIndex(e.RowHandle);

                        DateTime lapseDate = dt.Rows[row]["lapseDate8"].ObjToDateTime();
                        DateTime datePaid = dt.Rows[row]["DOLP"].ObjToDateTime();
                        TimeSpan ts = datePaid - lapseDate;
                        if (ts.TotalDays > 0)
                        {
                            e.Appearance.BackColor = System.Drawing.Color.Red;
                            e.Appearance.ForeColor = System.Drawing.Color.Yellow;
                        }
                    }
                }
            }
            else if (e.Column.FieldName.ToUpper() == "DECEASEDDATE")
            {
                if (e.RowHandle >= 0)
                {
                    if (!String.IsNullOrWhiteSpace(e.DisplayText))
                    {
                        if (e.DisplayText.IndexOf("0001") >= 0)
                            e.DisplayText = "00/00/0000";
                    }
                }
            }
            else if (e.Column.FieldName.ToUpper() == "TRUSTREMOVED")
            {
                if (e.RowHandle >= 0)
                {
                    DataTable dt = (DataTable)dgv.DataSource;
                    int row = gridMain.GetDataSourceRowIndex(e.RowHandle);

                    string trustRemoved = dt.Rows[row]["trustRemoved"].ObjToString().Trim().ToUpper();
                    if (trustRemoved == "YES" )
                    {
                        e.Appearance.BackColor = System.Drawing.Color.Red;
                        e.Appearance.ForeColor = System.Drawing.Color.Yellow;
                    }
                }
            }
            else if (e.Column.FieldName.ToUpper() == "NEWDUEDATE")
            {
                if (e.RowHandle >= 0)
                {
                    if (!String.IsNullOrWhiteSpace(e.DisplayText))
                    {
                        if (e.DisplayText.IndexOf("0001") >= 0)
                            e.DisplayText = "00/00/0000";
                    }
                }
            }
            else if (e.Column.FieldName.ToUpper() == "DOLP")
            {
                if (gridMain.Columns["DOLP"].Visible)
                {
                    if (e.RowHandle >= 0)
                    {
                        if (!String.IsNullOrWhiteSpace(e.DisplayText))
                        {
                            if (e.DisplayText.IndexOf("0001") >= 0)
                                e.DisplayText = "00/00/0000";
                            else
                            {
                                DataTable dt = (DataTable)dgv.DataSource;
                                int row = gridMain.GetDataSourceRowIndex(e.RowHandle);

                                DateTime lapseDate = dt.Rows[row]["lastDatePaid8"].ObjToDateTime();
                                DateTime datePaid = dt.Rows[row]["DOLP"].ObjToDateTime();
                                TimeSpan ts = datePaid - lapseDate;
                                if (ts.TotalDays != 0)
                                {
                                    e.Appearance.BackColor = System.Drawing.Color.Red;
                                    e.Appearance.ForeColor = System.Drawing.Color.Yellow;
                                }
                            }
                        }
                    }
                }
            }
        }
        /***********************************************************************************************/
        public delegate void d_void_eventdone_string(string s);
        public event d_void_eventdone_string ModuleDone;
        /***********************************************************************************************/
        private void repositoryItemCheckEdit2_Click(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            string record = dr["record"].ObjToString();
            if (ModuleDone != null)
                ModuleDone.Invoke(record);
            this.Close();
        }
        /***********************************************************************************************/
        private void trust85ReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            Form form = G1.IsFormOpen("Trust85");
            if (form != null)
            {
                form.Show();
                form.WindowState = FormWindowState.Normal;
                form.Visible = true;
                form.Refresh();
                form.BringToFront();
            }
            else
            {
                Trust85 trustForm = new Trust85(dt);
                trustForm.Show();
            }
        }
        /***********************************************************************************************/
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        { // Reinstate Customer
            DataTable dt = null;
            DataRow dr = null;
            string record = "";
            string contractNumber = "";

            if (dgv.Visible)
            {
                dt = (DataTable)dgv.DataSource;
                dr = gridMain.GetFocusedDataRow();
                record = dr["record"].ObjToString();
                contractNumber = dr["contractNumber"].ObjToString();
            }
            else
            {
                dt = (DataTable)dgv2.DataSource;
                dr = gridMain2.GetFocusedDataRow();
                record = dr["record"].ObjToString();
                contractNumber = dr["contractNumber"].ObjToString();
            }
            DialogResult result = MessageBox.Show("Are you sure you want to REINSTATE customer (" + contractNumber + ") ?", "Reinstate Contract Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (result == DialogResult.No)
                return;
            DateTime lapseDate = dr["lapseDate8"].ObjToDateTime();
            string date1 = dr["lapseDate8"].ObjToString();
            if ( date1.IndexOf ( "0000") >= 0 )
            {
                MessageBox.Show("***ERROR*** Contract ( " + contractNumber + ") is not Lapsed! Therefore, it cannot be Reinstated!");
                return;
            }

            DateTime today = DateTime.Now;
            string reinstateDate = today.ToString("yyyy-MM-dd");

            DailyHistory.SetReinstate(contractNumber, today.ToString("yyyy-MM-dd"));

            //record = dr["record"].ObjToString();
            //G1.update_db_table(customersFile, "record", record, new string[] { "lapsed", "" });
            //record = dr["record1"].ObjToString();
            //G1.update_db_table(customersFile, "record", record, new string[] { "lapsed", "", "lapseDate8", "0000-00-00", "reinstateDate8", reinstateDate});

            if ( DailyHistory.isInsurance ( contractNumber))
                CustomerDetails.UpdatePayersDetail(contractNumber);

            G1.AddToAudit(LoginForm.username, "Customers", "Reinstate", "Set", contractNumber);
        }
        /***********************************************************************************************/
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            DataTable dt = null;
            DataRow dr = null;
            string record = "";
            string contractNumber = "";
            int rowHandle = 0;
            int row = 0;

            if (dgv.Visible)
            {
                dt = (DataTable)dgv.DataSource;
                dr = gridMain.GetFocusedDataRow();
                rowHandle = gridMain.FocusedRowHandle;
                row = gridMain.GetDataSourceRowIndex(rowHandle);
                record = dr["record"].ObjToString();
                contractNumber = dr["contractNumber"].ObjToString();
            }
            else
            {
                dt = (DataTable)dgv2.DataSource;
                dr = gridMain2.GetFocusedDataRow();
                rowHandle = gridMain2.FocusedRowHandle;
                row = gridMain2.GetDataSourceRowIndex(rowHandle);
                record = dr["record"].ObjToString();
                contractNumber = dr["contractNumber"].ObjToString();
            }

            DialogResult result = MessageBox.Show("Are you sure you want to Clear Lapsed for customer (" + contractNumber + ") ?", "Clear Lapsed Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (result == DialogResult.No)
                return;

            if ( DailyHistory.ClearLapsed(contractNumber) )
            {
                DateTime myDate = "0000-00-00".ObjToDateTime();
                dt.Rows[row]["lapseDate8"] = G1.DTtoMySQLDT(myDate);
                dt.Rows[row]["lapsed"] = "";
                if ( dgv.Visible )
                    dt.Rows[row]["lapsed1"] = "";
                dt.AcceptChanges();
                dr["lapsed"] = "";
                dr["lapseDate8"] = G1.DTtoMySQLDT(myDate);
                if (dgv.Visible)
                {
                    gridMain.RefreshData();
                    dgv.Refresh();
                }
                else
                {
                    gridMain2.RefreshData();
                    dgv2.Refresh();
                }
            }
        }
        /***********************************************************************************************/
        private void weeklyCloseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WeeklyClose weekForm = new WeeklyClose();
            weekForm.Show();
        }
        /***********************************************************************************************/
        private void btnRecalc_Click(object sender, EventArgs e)
        {
            if (customersFile.ToUpper() == "ICUSTOMERS")
                return;
            chkMismatches.Visible = true;
            chkMismatchDates.Visible = true;
            gridMain.Columns["cbal"].Visible = true;
            gridMain.Columns["newduedate"].Visible = true;
            gridMain.Columns["days"].Visible = true;
            gridMain.Columns["creditBalance"].Visible = true;

            //gridMain.Columns["totalInterest"].Visible = true;
            //gridMain.Columns["cint"].Visible = true;

            barImport.Visible = true;
            barImport.Show();
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt == null)
                return;

            bool doMarked = false;

            int[] rows = gridMain.GetSelectedRows();
            if (rows.Length > 0)
            {
                DialogResult result = MessageBox.Show("Run on Selected Marked Rows ONLY?", "Selected Rows Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Yes)
                    doMarked = true;
            }
            int lastRow = dt.Rows.Count;
            if (doMarked)
                lastRow = rows.Length;

            string str = "";
            double trust85 = 0D;
            double balanceDue = 0D;
            double newBalance = 0D;
            double totalInterest = 0D;
            double startBalance = 0D;
            string contract = "";
            int numPayments = 0;
            double payment = 0D;
            double percentPaid = 0D;
            double contractValue = 0D;
            double dAPR = 0D;
            DateTime lastDate = DateTime.Now;
            DataTable dx = new DataTable();
            string cmd = "";
            DateTime testDate = new DateTime(2018, 7, 1);
            DateTime myTestDate = new DateTime(2018, 12, 31);
            DateTime maxDate = new DateTime(2019, 12, 31);
            DateTime docp = DateTime.Now;
            DateTime oldDueDate = DateTime.Now;
            string myFields = "";
            string pastRecord = "";
            string edited = "";
            double pastInterest = 0D;
            double interest = 0D;
            if (G1.get_column_number(dt, "fix") < 0)
                dt.Columns.Add("fix");

            TimeSpan ts;


            barImport.Minimum = 0;
            barImport.Maximum = lastRow;
            for (int i = 0; i < lastRow; i++)
            {
                barImport.Value = i;
                barImport.Refresh();
                try
                {
                    int row = i;
                    if (doMarked)
                    {
                        row = rows[i];
                        row = gridMain.GetDataSourceRowIndex(row);
                    }
                    DateTime date = dt.Rows[row]["issueDate8"].ObjToDateTime();

                    balanceDue = dt.Rows[row]["balanceDue"].ObjToDouble();
                    contract = dt.Rows[row]["contractNumber"].ObjToString();
                    payment = dt.Rows[row]["amtOfMonthlyPayt"].ObjToString().ObjToDouble();
                    numPayments = dt.Rows[row]["numberOfPayments"].ObjToString().ObjToInt32();
                    contractValue = DailyHistory.GetContractValue(dt.Rows[row]);
                    startBalance = DailyHistory.GetFinanceValue(dt.Rows[row]);
                    oldDueDate = dt.Rows[row]["dueDate"].ObjToDateTime();
                    if (oldDueDate.Year <= 1)
                    {
                        if (numPayments == 0 && contractValue == 0D && startBalance == 0D)
                        {
                            dt.Rows[row]["newDueDate"] = "12/31/2039";
                            dt.Rows[row]["days"] = 0;
                            continue;
                        }
                    }
                    if (oldDueDate.Year >= 2039)
                    {
                        dt.Rows[row]["newDueDate"] = dt.Rows[row]["dueDate"].ObjToDateTime();
                        dt.Rows[row]["days"] = 0;
                        continue;
                    }
                    if (startBalance <= 0D)
                    {
                        dt.Rows[row]["cbal"] = 0D;
                        if (numPayments <= 0)
                            dt.Rows[row]["cbal"] = startBalance;
                        dt.Rows[row]["newDueDate"] = dt.Rows[row]["dueDate"].ObjToDateTime();
                        dt.Rows[row]["days"] = 0;
                        continue;
                    }
                    else if (numPayments <= 0)
                    {
                        dt.Rows[row]["cbal"] = 0D;
                        dt.Rows[row]["newDueDate"] = dt.Rows[row]["dueDate"].ObjToDateTime();
                        dt.Rows[row]["days"] = 0;
                        continue;
                    }

                    DateTime iDate = DailyHistory.GetIssueDate(dt.Rows[row]["issueDate8"].ObjToDateTime(), contract, null);
                    string issueDate = iDate.ToString("MM/dd/yyyy");
                    lastDate = issueDate.ObjToDateTime();
                    string apr = dt.Rows[row]["APR"].ObjToString();
                    dAPR = apr.ObjToDouble() / 100.0D;

                    cmd = "Select * from `payments` where `contractNumber` = '" + contract + "' order by `paydate8` DESC, `tmstamp` DESC, `record` DESC;";
                    dx = G1.get_db_data(cmd);
                    DailyHistory.CalculateNewStuff(dx, dAPR, numPayments, startBalance, lastDate);
                    if (dx.Rows.Count <= 0)
                    {
                        dt.Rows[row]["cbal"] = startBalance;
                        continue;
                    }
                    double newbalance = 0D;
                    iDate = DailyHistory.getNextDueDate(dx, payment, ref newbalance);

                    //for (int j = 0; j < dx.Rows.Count; j++)
                    //{
                    //    if (dx.Rows[j]["fill"].ObjToString().ToUpper() == "D")
                    //        continue;

                    //    edited = dx.Rows[j]["edited"].ObjToString().Trim().ToUpper();
                    //    if (edited == "TRUSTADJ")
                    //        continue;

                    //    docp = dx.Rows[j]["payDate8"].ObjToDateTime();
                    //    if (docp > testDate && docp < maxDate)
                    //    {
                    //        //if (edited.ToUpper() == "MANUAL")
                    //        //    continue;
                    //        pastInterest = dx.Rows[j]["interestPaid"].ObjToDouble();
                    //        interest = dx.Rows[j]["int"].ObjToDouble();
                    //        if (interest != pastInterest)
                    //        {
                    //            if (interest == 0D || pastInterest == 0D)
                    //            {
                    //                gridMain.Columns["fix"].Visible = true;
                    //                pastRecord = dx.Rows[j]["record"].ObjToString();
                    //                myFields = "interestPaid," + interest.ToString();
                    //                dt.Rows[row]["fix"] = "FIX";
                    //                //                            G1.update_db_table("payments", "record", pastRecord, myFields);
                    //            }
                    //        }
                    //    }
                    //}

                    oldDueDate = dt.Rows[row]["dueDate"].ObjToDateTime();
                    for (int j = 0; j < dx.Rows.Count; j++)
                    {
                        if (dx.Rows[j]["fill"].ObjToString().ToUpper() == "D")
                            continue;
                        edited = dx.Rows[j]["edited"].ObjToString().Trim().ToUpper();
                        if (edited == "TRUSTADJ" || edited == "CEMETERY" )
                            continue;

                        newBalance = dx.Rows[j]["balance"].ObjToDouble();
                        iDate = dx.Rows[j]["nextDueDate"].ObjToDateTime();
                        if (newBalance <= 0D && oldDueDate.Year >= 2039)
                            iDate = oldDueDate;
                        ts = iDate - oldDueDate;
                        dt.Rows[row]["days"] = ts.Days;
                        dt.Rows[row]["newduedate"] = iDate.ToString("MM/dd/yyyy");
                        dt.Rows[row]["creditBalance"] = dx.Rows[j]["runningCB"].ObjToDouble();
                        dt.Rows[row]["cbal"] = newBalance;
//                        dt.Rows[row]["cint"] = totalInterest;
                        break;
                    }
                    //trust85 = DailyHistory.CalculatePayout(contract, dx);
                    //dt.Rows[row]["trust85"] = trust85;
                    //contractValue = dt.Rows[row]["contractValue"].ObjToDouble();
                    //if (contractValue > 0D)
                    //{
                    //    percentPaid = trust85 / contractValue;
                    //    //percentPaid = percentPaid * 100D;
                    //    //str = percentPaid.ToString("###.00%");
                    //    dt.Rows[row]["percentPaid"] = percentPaid;
                    //}
                }
                catch (Exception ex)
                {
                    MessageBox.Show("***ERROR*** Customer i=" + i.ToString() + " " + ex.Message.ToString());
                }
            }
            barImport.Value = lastRow;
            barImport.Refresh();
            dgv.DataSource = dt;
            dgv.Refresh();
            this.Refresh();
            this.Cursor = Cursors.Default;
            chkMismatches.Refresh();
            chkMismatchDates.Refresh();
            if (gridMain.Columns["fix"].Visible)
                btnFixInt.Show();
        }
        /***********************************************************************************************/
        private bool alreadyAddedMenu = false;
        private void btnRecalc_Click1(object sender, EventArgs e)
        {
            if (customersFile.ToUpper() == "ICUSTOMERS")
                return;
            if (!alreadyAddedMenu)
            {
                ToolStripMenuItem fixDueDate = new ToolStripMenuItem();
                fixDueDate.Text = "Fix Due Dates";
                fixDueDate.Click += FixDueDate_Click;
                miscToolStripMenuItem.DropDownItems.Add(fixDueDate);
            }
            chkMismatches.Visible = true;
            chkMismatchDates.Visible = true;
            gridMain.Columns["cbal"].Visible = true;
            gridMain.Columns["newduedate"].Visible = true;
            gridMain.Columns["days"].Visible = true;
            gridMain.Columns["creditBalance"].Visible = true;

            //gridMain.Columns["totalInterest"].Visible = true;
            //gridMain.Columns["cint"].Visible = true;

            barImport.Visible = true;
            barImport.Show();
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = (DataTable)dgv.DataSource;

            bool doMarked = false;

            int[] rows = gridMain.GetSelectedRows();
            if ( rows.Length > 0 )
            {
                DialogResult result = MessageBox.Show("Run on Selected Marked Rows ONLY?", "Selected Rows Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Yes)
                    doMarked = true;
            }
            int lastRow = dt.Rows.Count;
            if (doMarked)
                lastRow = rows.Length;

            string str = "";
            double trust85 = 0D;
            double balanceDue = 0D;
            double newBalance = 0D;
            double totalInterest = 0D;
            double startBalance = 0D;
            string contract = "";
            int numPayments = 0;
            double payment = 0D;
            double percentPaid = 0D;
            double contractValue = 0D;
            double dAPR = 0D;
            DateTime lastDate = DateTime.Now;
            DataTable dx = new DataTable();
            string cmd = "";
            DateTime testDate = new DateTime(2018, 7, 1);
            DateTime myTestDate = new DateTime(2018, 12, 31);
            DateTime maxDate = new DateTime(2019, 12, 31);
            DateTime docp = DateTime.Now;
            DateTime oldDueDate = DateTime.Now;
            string myFields = "";
            string pastRecord = "";
            string edited = "";
            double pastInterest = 0D;
            double interest = 0D;
            if (G1.get_column_number(dt, "fix") < 0)
                dt.Columns.Add("fix");

            TimeSpan ts;


            barImport.Minimum = 0;
            barImport.Maximum = lastRow;
            for (int i = 0; i < lastRow; i++)
            {
                barImport.Value = i;
                barImport.Refresh();
                try
                {
                    int row = i;
                    if (doMarked )
                    {
                        row = rows[i];
                        row = gridMain.GetDataSourceRowIndex(row);
                    }
                    DateTime date = dt.Rows[row]["issueDate8"].ObjToDateTime();
                    //if (date < testDate)
                    //    continue;
                    balanceDue = dt.Rows[row]["balanceDue"].ObjToDouble();
                    contract = dt.Rows[row]["contractNumber"].ObjToString();
                    payment = dt.Rows[row]["amtOfMonthlyPayt"].ObjToString().ObjToDouble();
                    numPayments = dt.Rows[row]["numberOfPayments"].ObjToString().ObjToInt32();
                    contractValue = DailyHistory.GetContractValue(dt.Rows[row]);
                    startBalance = DailyHistory.GetFinanceValue(dt.Rows[row]);
                    oldDueDate = dt.Rows[row]["dueDate"].ObjToDateTime();
                    if ( oldDueDate.Year <= 1 )
                    {
                        if ( numPayments == 0 && contractValue == 0D && startBalance == 0D)
                        {
                            dt.Rows[row]["newDueDate"] = "12/31/2039";
                            dt.Rows[row]["days"] = 0;
                            continue;
                        }
                    }
                    if ( oldDueDate.Year >= 2039)
                    {
                        dt.Rows[row]["newDueDate"] = dt.Rows[row]["dueDate"].ObjToDateTime();
                        dt.Rows[row]["days"] = 0;
                        continue;
                    }
                    if (startBalance <= 0D)
                    {
                        dt.Rows[row]["cbal"] = 0D;
                        if (numPayments <= 0)
                            dt.Rows[row]["cbal"] = startBalance;
                        dt.Rows[row]["newDueDate"] = dt.Rows[row]["dueDate"].ObjToDateTime();
                        dt.Rows[row]["days"] = 0;
                        continue;
                    }
                    else if ( numPayments <= 0 )
                    {
                        dt.Rows[row]["cbal"] = 0D;
                        dt.Rows[row]["newDueDate"] = dt.Rows[row]["dueDate"].ObjToDateTime();
                        dt.Rows[row]["days"] = 0;
                        continue;
                    }
                    DateTime iDate = DailyHistory.GetIssueDate(dt.Rows[row]["issueDate8"].ObjToDateTime(), contract, null);
                    string issueDate = iDate.ToString("MM/dd/yyyy");
                    lastDate = issueDate.ObjToDateTime();
                    string apr = dt.Rows[row]["APR"].ObjToString();
                    dAPR = apr.ObjToDouble() / 100.0D;

                    cmd = "Select * from `payments` where `contractNumber` = '" + contract + "' order by `paydate8` DESC, `record` DESC;";
                    dx = G1.get_db_data(cmd);
                    DailyHistory.CalculateNewStuff(dx, dAPR, numPayments, startBalance, lastDate);
                    if (dx.Rows.Count > 0)
                    {
                        for (int j = 0; j < dx.Rows.Count;j++)
                        {
                            if (dx.Rows[j]["fill"].ObjToString().ToUpper() == "D")
                                continue;
                            docp = dx.Rows[j]["payDate8"].ObjToDateTime();
                            if (docp > testDate && docp < maxDate )
                            {
                                edited = dx.Rows[j]["edited"].ObjToString();
                                if (edited.ToUpper() == "MANUAL")
                                    continue;
                                pastInterest = dx.Rows[j]["interestPaid"].ObjToDouble();
                                interest = dx.Rows[j]["int"].ObjToDouble();
                                if (interest != pastInterest)
                                {
                                    if (interest == 0D || pastInterest == 0D)
                                    {
                                        gridMain.Columns["fix"].Visible = true;
                                        pastRecord = dx.Rows[j]["record"].ObjToString();
                                        myFields = "interestPaid," + interest.ToString();
                                        dt.Rows[row]["fix"] = "FIX";
                                        //                            G1.update_db_table("payments", "record", pastRecord, myFields);
                                    }
                                }
                            }
                        }

                        oldDueDate = dt.Rows[row]["dueDate"].ObjToDateTime();
                        for (int j = 0; j < dx.Rows.Count; j++)
                        {
                            if (dx.Rows[j]["fill"].ObjToString().ToUpper() == "D")
                                continue;
                            newBalance = dx.Rows[j]["newbalance"].ObjToDouble();
                            iDate = dx.Rows[j]["nextDueDate"].ObjToDateTime();
                            if (newBalance <= 0D && oldDueDate.Year >= 2039)
                                iDate = oldDueDate;
                            ts = iDate - oldDueDate;
                            dt.Rows[row]["days"] = ts.Days;
                            dt.Rows[row]["newduedate"] = iDate.ToString("MM/dd/yyyy");
                            dt.Rows[row]["creditBalance"] = dx.Rows[j]["creditBalance"].ObjToDouble();
                            dt.Rows[row]["cbal"] = newBalance;
                            dt.Rows[row]["cint"] = totalInterest;
                            break;
                        }
                        trust85 = DailyHistory.CalculatePayout(contract, dx);
                        dt.Rows[row]["trust85"] = trust85;
                        contractValue = dt.Rows[row]["contractValue"].ObjToDouble();
                        if (contractValue > 0D)
                        {
                            percentPaid = trust85 / contractValue;
                            //percentPaid = percentPaid * 100D;
                            //str = percentPaid.ToString("###.00%");
                            dt.Rows[row]["percentPaid"] = percentPaid;
                        }
                    }
                    else
                    {
                        dt.Rows[row]["cbal"] = startBalance;
                    }
                }
                catch ( Exception ex)
                {
                    MessageBox.Show("***ERROR*** Customer i=" + i.ToString() + " " + ex.Message.ToString());
                }
            }
            barImport.Value = lastRow;
            barImport.Refresh();
            dgv.DataSource = dt;
            dgv.Refresh();
            this.Refresh();
            this.Cursor = Cursors.Default;
            chkMismatches.Refresh();
            chkMismatchDates.Refresh();
            if (gridMain.Columns["fix"].Visible)
                btnFixInt.Show();
        }
        ///***********************************************************************************************/
        private void FixDueDate_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = (DataTable)dgv.DataSource;
            string contractNumber = "";
            string record = "";
            DateTime newDueDate = DateTime.Now;
            string dueDate = "";
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                newDueDate = dt.Rows[i]["newDueDate"].ObjToDateTime();
                if ( newDueDate.Year > 100)
                {
                    record = dt.Rows[i]["record1"].ObjToString();
                    dueDate = G1.DTtoMySQLDT(newDueDate.ToString ("MM/dd/yyyy")).ObjToString();
                    G1.update_db_table("contracts", "record", record, new string[] { "dueDate8", dueDate});
                }
            }
            this.Cursor = Cursors.Default;
        }
        ///***********************************************************************************************/
        //private void btnRecalc_Clickx(object sender, EventArgs e)
        //{
        //    chkMismatches.Visible = true;
        //    gridMain.Columns["cbal"].Visible = true;
        //    gridMain.Columns["newduedate"].Visible = true;
        //    gridMain.Columns["creditBalance"].Visible = true;
        //    //gridMain.Columns["totalInterest"].Visible = true;
        //    //gridMain.Columns["cint"].Visible = true;

        //    barImport.Visible = true;
        //    barImport.Show();
        //    this.Cursor = Cursors.WaitCursor;
        //    DataTable dt = (DataTable)dgv.DataSource;
        //    double balanceDue = 0D;
        //    double newBalance = 0D;
        //    double totalInterest = 0D;
        //    string contract = "";
        //    barImport.Minimum = 0;
        //    barImport.Maximum = dt.Rows.Count;
        //    for (int i = 0; i < dt.Rows.Count; i++)
        //    {
        //        barImport.Value = i;
        //        barImport.Refresh();
        //        balanceDue = dt.Rows[i]["balanceDue"].ObjToDouble();
        //        contract = dt.Rows[i]["contractNumber"].ObjToString();
        //        //if (contract != "WM17018UI")
        //        //    continue;
        //        newBalance = DailyHistory.ReCalculateDetails(contract, ref totalInterest);
        //        dt.Rows[i]["cbal"] = newBalance;
        //        dt.Rows[i]["cint"] = totalInterest;
        //    }
        //    dgv.DataSource = dt;
        //    dgv.Refresh();
        //    this.Cursor = Cursors.Default;
        //    chkMismatches.Refresh();
        //}
        /***********************************************************************************************/
        private void chkMismatches_CheckedChanged(object sender, EventArgs e)
        {
            dgv.Refresh();
            gridMain.RefreshData();
            this.Refresh();
        }
        /***********************************************************************************************/
        private void chkMismatchDates_CheckedChanged(object sender, EventArgs e)
        {
            dgv.Refresh();
            gridMain.RefreshData();
            this.Refresh();
        }
        /***********************************************************************************************/
        private void gridMain_CustomRowFilter(object sender, DevExpress.XtraGrid.Views.Base.RowFilterEventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            if (chkMismatches.Checked)
            {
                int rowHandle = e.ListSourceRow;
                int row = gridMain.GetDataSourceRowIndex(rowHandle);
                if (row < 0)
                    return;
                double balanceDue = dt.Rows[row]["balanceDue"].ObjToDouble();
                balanceDue = G1.RoundValue(balanceDue);
                double cbal = dt.Rows[row]["cbal"].ObjToDouble();
                cbal = G1.RoundValue(cbal);
                double diff = cbal - balanceDue;
                string str = txtThreshold.Text;
                double threshold = 0D;
                if (G1.validate_numeric(str))
                    threshold = str.ObjToDouble();
                if ( diff >= -threshold && diff <= threshold )
                {
                    e.Visible = false;
                    e.Handled = true;
                    return;
                }
            }
            if ( chkLessZero.Checked )
            {
                int rowHandle = e.ListSourceRow;
                int row = gridMain.GetDataSourceRowIndex(rowHandle);
                if (row < 0)
                    return;
                double balanceDue = dt.Rows[row]["balanceDue"].ObjToDouble();
                balanceDue = G1.RoundValue(balanceDue);
                if ( balanceDue > 0D)
                {
                    e.Visible = false;
                    e.Handled = true;
                    return;
                }
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
            if (this.components == null)
                this.components = new System.ComponentModel.Container();
            DevExpress.XtraPrinting.PrintingSystem printingSystem1 = new DevExpress.XtraPrinting.PrintingSystem(this.components);
            DevExpress.XtraPrinting.PrintableComponentLink printableComponentLink1 = new DevExpress.XtraPrinting.PrintableComponentLink(this.components);

            printingSystem1.Links.AddRange(new object[] {
            printableComponentLink1});

            printableComponentLink1.Component = dgv;
            if (dgv2.Visible)
                printableComponentLink1.Component = dgv2;
            printableComponentLink1.PrintingSystemBase = printingSystem1;
            printableComponentLink1.CreateDetailHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateDetailHeaderArea);
            printableComponentLink1.CreateMarginalHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateMarginalHeaderArea);
            printableComponentLink1.BeforeCreateAreas += new System.EventHandler(this.printableComponentLink1_BeforeCreateAreas);
            printableComponentLink1.AfterCreateAreas += new System.EventHandler(this.printableComponentLink1_AfterCreateAreas);
            printableComponentLink1.Landscape = true;

            Printer.setupPrinterMargins(50, 100, 80, 50);

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
            printableComponentLink1.ShowPreview();
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
            if (dgv2.Visible)
                printableComponentLink1.Component = dgv2;
            printableComponentLink1.PrintingSystemBase = printingSystem1;
            printableComponentLink1.CreateDetailHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateDetailHeaderArea);
            printableComponentLink1.CreateMarginalHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateMarginalHeaderArea);
            printableComponentLink1.BeforeCreateAreas += new System.EventHandler(this.printableComponentLink1_BeforeCreateAreas);
            printableComponentLink1.AfterCreateAreas += new System.EventHandler(this.printableComponentLink1_AfterCreateAreas);
            printableComponentLink1.Landscape = true;

            Printer.setupPrinterMargins(50, 100, 80, 50);

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
            if (LoginForm.doLapseReport)
                printableComponentLink1.Print();
            else
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

            font = new Font("Ariel", 6);
            Printer.DrawGridDate(2, 3, 2, 3, Color.Black, BorderSide.None, font);
            Printer.DrawGridPage(11, 3, 2, 3, Color.Black, BorderSide.None, font);

            //            Printer.DrawQuad(1, 9, 2, 3, "User : " + LoginForm.username, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Center);

            font = new Font("Ariel", 10, FontStyle.Bold);
            string title = "Customer Report";
            string qualify = cmbQualify.Text.ToUpper();
            if (qualify == "LAPSED")
                title = "Lapsed Customer Report";
            else if (qualify == "DECEASED")
                title = "Deceased Customer Report";
            Printer.DrawQuad(6, 8, 2, 4, title, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);


            //            Printer.DrawQuadTicks();
            Printer.SetQuadSize(24, 12);
            font = new Font("Ariel", 7, FontStyle.Regular);
            //Printer.DrawQuad(16, 7, 5, 2, lblBalance.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Top);
            //Printer.DrawQuad(16, 10, 5, 2, lblPayment.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);

            font = new Font("Ariel", 8);
//            Printer.DrawQuad(1, 6, 6, 3, search, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Top);
            Printer.DrawQuad(1, 9, 6, 3, "User : " + LoginForm.username, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Center);


            Printer.SetQuadSize(12, 12);
            Printer.DrawQuadBorder(1, 1, 12, 12, BorderSide.All, 1, Color.Black);
            Printer.DrawQuadBorder(12, 1, 1, 12, BorderSide.Right, 1, Color.Black);
        }
        /***********************************************************************************************/
        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            PaymentsReport paymentForm = new PaymentsReport();
            paymentForm.Show();
        }
        /***********************************************************************************************/
        private void showSimpleSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            if (!String.IsNullOrWhiteSpace(contract))
            {
                SimpleSummary simpleForm = new SimpleSummary(contract);
                simpleForm.Show();
            }
        }
        /***********************************************************************************************/
        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            string name = dr["fullname"].ObjToString();
            if (!String.IsNullOrWhiteSpace(contract))
            {
                ManualPayment manualForm = new ManualPayment(contract, name);
                manualForm.Show();
            }
        }
        /***********************************************************************************************/
        private void historicCommissionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //HistoricCommissions histForm = new HistoricCommissions();
            //histForm.Show();

            this.Cursor = Cursors.WaitCursor;
            AgentYearly agentForm = new AgentYearly();
            agentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void changeContractNumberToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ( !LoginForm.administrator )
            {
                MessageBox.Show("***ERROR*** You do not have permission to do this.");
                return;
            }
            string goodContractNumber = "";

            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);
            DataTable dt = (DataTable)dgv.DataSource;

            string badContractNumber = dr["contractNumber"].ObjToString();

            using (Ask askForm = new Ask("Enter Good Contract #?"))
            {
                askForm.Text = "";
                askForm.ShowDialog();
                if (askForm.DialogResult != System.Windows.Forms.DialogResult.OK)
                    return;
                goodContractNumber = askForm.Answer;
                if (String.IsNullOrWhiteSpace(goodContractNumber))
                    return;
            }


            DialogResult result = MessageBox.Show("*** CONFIRM *** Are you SURE you want to change contract number " + badContractNumber + " to " + goodContractNumber + "?", "Change Contract # Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            if (result == DialogResult.No)
                return;
            this.Cursor = Cursors.WaitCursor;
            if (contractsFile.ToUpper() == "ICONTRACTS")
            {
                ChangeContractNumber("icontracts", badContractNumber, goodContractNumber);
                ChangeContractNumber("icustomers", badContractNumber, goodContractNumber);
                ChangeContractNumber("ipayments", badContractNumber, goodContractNumber);
            }
            else
            {
                ChangeContractNumber("contracts", badContractNumber, goodContractNumber);
                ChangeContractNumber("customers", badContractNumber, goodContractNumber);
                ChangeContractNumber("payments", badContractNumber, goodContractNumber);
                ChangeContractNumber("cust_services", badContractNumber, goodContractNumber);
            }
            dr["contractNumber"] = goodContractNumber;
            dt.Rows[row]["contractNumber"] = goodContractNumber;
            this.Cursor = Cursors.Default;
            MessageBox.Show("*** Good News *** Contracts are changes!", "Change Contreact Number Dialog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }
        /***********************************************************************************************/
        private void ChangeContractNumber ( string table, string badContractNumber, string goodContractNumber )
        {
            string record = "";
            string cmd = "Select * from `" + table + "` where `contractNumber` = '" + goodContractNumber + "';";
            DataTable dt = G1.get_db_data(cmd);
            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("***WARNING*** Contract " + goodContractNumber + " Already Exists in " + table + " Table!");
            }
            else
            {
                cmd = "Select * from `" + table + "` where `contractNumber` = '" + badContractNumber + "';";
                dt = G1.get_db_data(cmd);
                if (dt.Rows.Count <= 0)
                {
                    MessageBox.Show("***WARNING*** Contract " + badContractNumber + " DOES NOT Exists in " + table + " Table!");
                }
                else
                {
                    for ( int i=0; i<dt.Rows.Count; i++)
                    {
                        record = dt.Rows[i]["record"].ObjToString();
                        G1.update_db_table(table, "record", record, new string[] { "contractNumber", goodContractNumber });
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void showAutoDraftsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutoDrafts autoForm = new AutoDrafts();
            autoForm.Show();
        }
        /***********************************************************************************************/
        private void btnDOLP_Click(object sender, EventArgs e)
        {
            string cnum = "";
            string cmd = "";
            DateTime dolp = DateTime.Now;
            DataTable dt = (DataTable)dgv.DataSource;
            DataTable dx = null;
            this.Cursor = Cursors.WaitCursor;
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                cnum = dt.Rows[i]["contractNumber"].ObjToString();
                if (!String.IsNullOrWhiteSpace(cnum))
                {
                    string paymentsFile = "payments";
                    if (contractsFile.ToUpper() == "ICONTRACTS")
                        paymentsFile = "ipayments";
                    cmd = "Select * from `" + paymentsFile + "` where `contractNumber` = '" + cnum + "' order by `payDate8` DESC LIMIT 1";
                    dx = G1.get_db_data(cmd);
                    if (dx.Rows.Count > 0)
                    {
                        dolp = dx.Rows[0]["payDate8"].ObjToDateTime();
                        dt.Rows[i]["DOLP"] = dolp.ToString("MM/dd/yyyy");
                    }
                }
            }
            gridMain.Columns["DOLP"].Visible = true;
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void gridMain_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (loading)
                return;
            DataRow dr = gridMain.GetFocusedDataRow();
            DataTable dt = (DataTable)dgv.DataSource;
            if (e.Column.FieldName.ToUpper() == "DUEDATE")
            {
                int rowhandle = gridMain.FocusedRowHandle;
                int row = gridMain.GetDataSourceRowIndex(rowhandle);
                string cnum = dt.Rows[row]["contractNumber"].ObjToString();
                if ( !String.IsNullOrWhiteSpace (cnum))
                {
                    string date = e.Value.ObjToString();
                    if (G1.validate_date(date))
                    {
                        MySqlDateTime myDate = (MySqlDateTime)G1.DTtoMySQLDT(date);
                        string dueDate = myDate.Month.ToString("D2") + "/" + myDate.Day.ToString("D2") + "/" + myDate.Year.ToString("D4");
                        string record = dt.Rows[row]["record1"].ObjToString();
                        G1.update_db_table(contractsFile, "record", record, new string[] {"dueDate8", dueDate });
                        BandedGridView view = sender as BandedGridView;
                        loading = true;
                        view.SetRowCellValue(e.RowHandle, view.Columns["dueDate"], dueDate);
                        loading = false;
                    }
                }
            }
            if (!chkFullEdit.Checked)
            {
                int rowhandle = gridMain.FocusedRowHandle;
                int row = gridMain.GetDataSourceRowIndex(rowhandle);
                string column = e.Column.FieldName;
                return;
            }
            string type = cmbType.Text;
            if (type.ToUpper() == "TRUSTS")
                modTrusts = true;
            else
                modInsurance = true;

            dr["mod"] = "Y";
            btnSave.Show();
            btnSave.Refresh();
        }
        /***********************************************************************************************/
        private void txtThreshold_TextChanged(object sender, EventArgs e)
        {
            //if (loading)
            //    return;
            //string str = txtThreshold.Text;
            //if ( !G1.validate_numeric ( str ))
            //    str = "0.00";
            //double value = str.ObjToDouble();
            //str = G1.ReformatMoney(value);
            //loading = true;
            //txtThreshold.Text = str;
            //loading = false;
        }
        /***********************************************************************************************/
        private void txtThreshold_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string str = txtThreshold.Text;
                if (!G1.validate_numeric(str))
                    str = "0.00";
                double value = str.ObjToDouble();
                str = G1.ReformatMoney(value);
                loading = true;
                txtThreshold.Text = str;
                loading = false;
            }
        }
        /***********************************************************************************************/
        private void SetupSelectedColumns()
        {
            string group = cmbSelectColumns.Text.Trim().ToUpper();
            if (group.Trim().Length == 0)
                return;
            string cmd = "Select * from procfiles where name = '" + group + "' and ProcType = 'Customers' order by seq";
            DataTable dt = G1.get_db_data(cmd);
            for (int i = 0; i < gridMain.Columns.Count; i++)
                gridMain.Columns[i].Visible = false;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string name = dt.Rows[i]["Description"].ToString();
                int index = dt.Rows[i]["seq"].ObjToInt32();
                try
                {
                    gridMain.Columns[name].Visible = true;
                }
                catch
                {
                }
            }
        }
        /***********************************************************************************************/
        private void SetupSelectedColumns(string procType, string group, DevExpress.XtraGrid.GridControl dgv)
        {
            if (String.IsNullOrWhiteSpace(group))
                return;
            if (String.IsNullOrWhiteSpace(procType))
                procType = "Customers";
            string cmd = "Select * from procfiles where name = '" + group + "' and ProcType = '" + procType + "' order by seq";
            DataTable dt = G1.get_db_data(cmd);
            DevExpress.XtraGrid.Views.BandedGrid.AdvBandedGridView gridMain = (DevExpress.XtraGrid.Views.BandedGrid.AdvBandedGridView)dgv.MainView;
            for (int i = 0; i < gridMain.Columns.Count; i++)
                gridMain.Columns[i].Visible = false;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string name = dt.Rows[i]["Description"].ToString();
                int index = dt.Rows[i]["seq"].ObjToInt32();
                try
                {
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
            //SetupSelectedColumns("Customers", "Primary", dgv);
            string format = cmbSelectColumns.Text.Trim();
            if ( !String.IsNullOrWhiteSpace ( format))
                SetupSelectedColumns("Customers", format, dgv);
        }
        /***********************************************************************************************/
        private void btnSelectColumns_Click(object sender, EventArgs e)
        {
            string actualName = cmbSelectColumns.Text;
            SelectColumns sform = new SelectColumns(dgv, "Customers", "Primary", actualName);
            sform.Done += new SelectColumns.d_void_eventdone(sform_Done);
            sform.Show();
        }
        /***********************************************************************************************/
        void sform_Done()
        {
            dgv.Refresh();
            this.Refresh();
        }
        /***********************************************************************************************/
        private void feeAllocationReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TrustsFees trustForm = new TrustsFees();
            trustForm.Show();
        }
        /***********************************************************************************************/
        private void genericPaymentsReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport();
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void paidUpContractsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Paid Up Contracts Report", "Trust Paid Off Contracts (2.0)" );
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void reportPaymentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            WeeklyClose weekForm = new WeeklyClose("Trust Payment Weekly Balance Sheet (1.3)");
            weekForm.Show();
            //PaymentsReport paymentForm = new PaymentsReport("Payments Report", "Trust Payment Weekly Balance Sheet (1.3)" );
            //paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void allActivityReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("All Activity Report", "");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void activeOnlyReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Active Only Report", "");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void downPaymentsReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Down Payments Report", "Trust Down Payment Master List Report (6.1)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Payments Report", "Trust Monthly Payment 85% Master Listing (6.2)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
        }
        /***********************************************************************************************/
        private void balancesPaymentsReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Balances Less Than Payments Report", "");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void balancesXReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Balances Less Than X Report", "");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void paymentsAfterDeathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Payments After Death Report", "Payments After Death Report (5.3)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem4_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("New Business Report", "New Business Report (1.2)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void trustEOMReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Trust EOM Report", "Trust EOM Report (1.1)" );
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Reinstatement Report", "Reinstatement Report (5.0)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("DBR Report", "DBR Report (5.2)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            PastDue pastForm = new PastDue( "Potential Lapse Report (3.0)");
            pastForm.Show();
        }
        /***********************************************************************************************/
        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            PastDue pastForm = new PastDue("Trust Lapse List (4.0)");
            pastForm.Show();
        }
        /***********************************************************************************************/
        private void menuRemovalReport_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Removal Report", "Removal Report (5.1)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Cash Remitted Report", "Cash Remitted Report (8.0)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void weekTotalsReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Insurance Week Totals", "" );
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void ChangeComboDropDown( string type )
        {
            cmbQualify.Items.Clear();
            cmbQualify.Items.Add("Active");
            cmbQualify.Items.Add("Lapsed");
            cmbQualify.Items.Add("Deceased");
            if (type == "INSURANCE")
            {
                if (!G1.isField())
                {
                    cmbQualify.Items.Add("Orphans");
                    cmbQualify.Items.Add("All");
                }
            }
            else
                cmbQualify.Items.Add("All");
        }
        /***********************************************************************************************/
        private void CleanUpMenu ( string type, bool loadData = true )
        {
            ChangeComboDropDown(type);
            bool killInsurance = true;
            if (type.ToUpper() == "INSURANCE" || type.ToUpper() == "PAYERS" )
            {
                killInsurance = false;
                reportsToolStripMenuItem.HideDropDown();
                editToolStripMenuItem.HideDropDown();
                editTrustDownPaymentsToolStripMenuItem1.HideDropDown();
                btnRecalc.Hide();
                txtThreshold.Hide();
                label1.Hide();
                chkMismatchDates.Hide();
                chkMismatches.Hide();
                barImport.Hide();
                SetupMenuPreferences();
                string preference = G1.getPreference(LoginForm.username, "Insurance Reports", "Allow Access");
                if (preference != "YES")
                    killInsurance = true;
            }
            else
            {
                insuranceReportsToolStripMenuItem.HideDropDown();
                chkMismatchDates.Show();
                chkMismatches.Show();
                barImport.Show();
                btnRecalc.Show();
                SetupMenuPreferences();
            }

            string name = "";
            for (int i = (menuStrip1.Items.Count - 1); i >= 0; i--)
            {
                name = menuStrip1.Items[i].Name.ToUpper();
                if (killInsurance)
                {
                    if (name == "INSURANCEREPORTSTOOLSTRIPMENUITEM")
                    {
                        menuStrip1.Items[i].Enabled = false;
                        menuStrip1.Items[i].Visible = false;
                    }
                    else if (name == "REPORTSTOOLSTRIPMENUITEM")
                    {
                        if (!G1.isField())
                        {
                            menuStrip1.Items[i].Enabled = true;
                            menuStrip1.Items[i].Visible = true;
                        }
                        else
                        {
                            menuStrip1.Items[i].Enabled = false;
                            menuStrip1.Items[i].Visible = false;
                        }
                    }
                    else if (name == "EDITTOOLSTRIPMENUITEM")
                    {
                        if (!G1.isField())
                        {
                            menuStrip1.Items[i].Enabled = true;
                            menuStrip1.Items[i].Visible = true;
                        }
                        else
                        {
                            menuStrip1.Items[i].Enabled = false;
                            menuStrip1.Items[i].Visible = false;
                        }
                    }
                    else if (name == "COMMONTOOLSTRIPMENUITEM")
                    {
                        if (!G1.isField())
                        {
                            menuStrip1.Items[i].Enabled = true;
                            menuStrip1.Items[i].Visible = true;
                        }
                        else
                        {
                            menuStrip1.Items[i].Enabled = false;
                            menuStrip1.Items[i].Visible = false;
                        }
                    }
                }
                else
                {
                    if (name == "REPORTSTOOLSTRIPMENUITEM")
                    {
                        menuStrip1.Items[i].Enabled = false;
                        menuStrip1.Items[i].Visible = false;
                    }
                    else if (name == "INSURANCEREPORTSTOOLSTRIPMENUITEM")
                    {
                        if (!G1.isField())
                        {
                            menuStrip1.Items[i].Enabled = true;
                            menuStrip1.Items[i].Visible = true;
                        }
                    }
                    else if (name == "EDITTOOLSTRIPMENUITEM")
                    {
                        menuStrip1.Items[i].Enabled = false;
                        menuStrip1.Items[i].Visible = false;
                    }
                }
            }
            if ( type.ToUpper() == "PAYERS")
            {
                dgv.Hide();
                dgv2.Show();
            }
            else
            {
                dgv.Hide();
                dgv2.Show();
            }
            //if ( loadData )
            //    LoadData();
        }
        /***********************************************************************************************/
        private void cmbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            string type = cmbType.Text.ToUpper();
            CleanUpMenu(type);
            if ( type.ToUpper() == "FUNERALS")
                this.Text = "Funeral Customers";
            else
                this.Text = type + " Customers";
        }
        /***********************************************************************************************/
        private void cmbQualify_SelectedIndexChanged(object sender, EventArgs e)
        {
//            LoadData();
        }
        /***********************************************************************************************/
        private void lapseReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PastDueInsurance pastForm = new PastDueInsurance("Insurance Lapse Report");
            pastForm.Show();
        }
        /***********************************************************************************************/
        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Show Payments Due Next Month", "");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void chkSortLastName_CheckedChanged(object sender, EventArgs e)
        {
            bool sort = false;
            if (chkSortLastName.Checked)
                sort = true;
            string type = cmbType.Text.Trim().ToUpper();
            if (dgv.Visible)
            {
                DataTable dt = (DataTable)dgv.DataSource;
                DataView tempview = dt.DefaultView;
                if (sort)
                {
                    tempview.Sort = "lastName,firstName,contractNumber";
                    gridMain.Columns["lastName"].Visible = true;
                    gridMain.Columns["firstName"].Visible = true;
                    gridMain.Columns["fullname"].Visible = false;
                }
                else
                {
                    tempview.Sort = "contractNumber";
                    gridMain.Columns["lastName"].Visible = true;
                    gridMain.Columns["firstName"].Visible = true;
                    gridMain.Columns["fullname"].Visible = false;
                }
                dt = tempview.ToTable();
                dgv.DataSource = dt;
                gridMain.ExpandAllGroups();
            }
            else if (dgv2.Visible)
            {
                DataTable dt = (DataTable)dgv2.DataSource;
                DataView tempview = dt.DefaultView;
                if (sort)
                {
                    tempview.Sort = "lastName,firstName,payer";
                    gridMain2.Columns["lastName"].Visible = true;
                    gridMain2.Columns["firstName"].Visible = true;
                    gridMain2.Columns["fullname"].Visible = false;
                }
                else
                {
                    tempview.Sort = "payer";
                    gridMain2.Columns["lastName"].Visible = true;
                    gridMain2.Columns["firstName"].Visible = true;
                    gridMain2.Columns["fullname"].Visible = false;
                }
                dt = tempview.ToTable();
                dgv2.DataSource = dt;
                gridMain2.ExpandAllGroups();
            }
        }
        /***********************************************************************************************/
        /***********************************************************************************************/
        private void ListForm_ListDone(string s)
        {
            string[] Lines = s.Split('\n');
            int Count = Lines.Length;
            if (Count == 0)
                return;
            string[] lines = Lines[0].Split(',');
            string contractNumber = lines[0].Trim();
            this.Cursor = Cursors.WaitCursor;
            CustomerDetails clientForm = new CustomerDetails(contractNumber);
            clientForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void chkLessZero_CheckedChanged(object sender, EventArgs e)
        {
            dgv.Refresh();
            dgv.RefreshDataSource();
            this.Refresh();
        }
        /***********************************************************************************************/
        private void quickLookupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form form = G1.IsFormOpen("FastLookup");
            if ( form != null )
            {
                form.Show();
                form.WindowState = FormWindowState.Normal;
                form.Visible = true;
                form.Refresh();
                form.BringToFront();
            }
            else
            {
                FastLookup fastForm = new FastLookup();
                fastForm.Show();
            }
        }
        /***********************************************************************************************/
        private void importDailyDepositFileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            this.Cursor = Cursors.WaitCursor;
            ImportDailyDeposits dailyForm = new ImportDailyDeposits(dt);
            dailyForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void importACHFileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            this.Cursor = Cursors.WaitCursor;
            ImportDailyDeposits dailyForm = new ImportDailyDeposits(dt, true);
            dailyForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void importCreditCardPaymentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            this.Cursor = Cursors.WaitCursor;
            ImportDailyDeposits dailyForm = new ImportDailyDeposits(dt, "CC");
            dailyForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void dailyDepositReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DailyReport dailyForm = new DailyReport();
            dailyForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void collectionReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            InsuranceCollectionsReport collectionForm = new InsuranceCollectionsReport();
            collectionForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Odd Payments Report", "" );
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void massLapseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PastDueMass pastForm = new PastDueMass("Mass Lapse/Reinstate Report");
            pastForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void btnFixInt_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            this.Cursor = Cursors.WaitCursor;
            string contractNumber = "";
            string fix = "";
            string cmd = "";
            DataTable dx = null;
            DateTime lastDate = DateTime.Now;
            double dAPR = 0D;
            double startBalance = 0D;
            int numPayments = 0;
            DateTime myTestDate = new DateTime(2018, 12, 31);
            DateTime docp = DateTime.Now;
            double balanceDue = 0D;
            double payment = 0D;
            double pastInterest = 0D;
            double interest = 0D;
            string myFields = "";
            string pastRecord = "";
            double difference = 0D;
            string contractRecord = "";

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                fix = dt.Rows[i]["fix"].ObjToString();
                if (fix.ToUpper() == "FIX")
                {
                    contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                    contractRecord = dt.Rows[i]["record1"].ObjToString();

                    DateTime date = dt.Rows[i]["issueDate8"].ObjToDateTime();
                    balanceDue = dt.Rows[i]["balanceDue"].ObjToDouble();
                    payment = dt.Rows[i]["amtOfMonthlyPayt"].ObjToString().ObjToDouble();
                    numPayments = dt.Rows[i]["numberOfPayments"].ObjToString().ObjToInt32();
                    startBalance = DailyHistory.GetFinanceValue(dt.Rows[i]);
                    if (startBalance <= 0D)
                    {
                        dt.Rows[i]["cbal"] = startBalance;
                        continue;
                    }
                    DateTime iDate = DailyHistory.GetIssueDate(dt.Rows[i]["issueDate8"].ObjToDateTime(), contractNumber, null);
                    string issueDate = iDate.ToString("MM/dd/yyyy");
                    lastDate = issueDate.ObjToDateTime();
                    string apr = dt.Rows[i]["APR"].ObjToString();
                    dAPR = apr.ObjToDouble() / 100.0D;

                    cmd = "Select * from `payments` where `contractNumber` = '" + contractNumber + "' order by `paydate8` DESC, `record` DESC;";
                    dx = G1.get_db_data(cmd);
                    DailyHistory.CalculateNewStuff(dx, dAPR, numPayments, startBalance, lastDate);
                    if (dx.Rows.Count > 0)
                    {
                        for (int j = 0; j < dx.Rows.Count; j++)
                        {
                            if (dx.Rows[j]["fill"].ObjToString().ToUpper() == "D")
                                continue;
                            docp = dx.Rows[j]["payDate8"].ObjToDateTime();
                            if (docp > myTestDate)
                            {
                                pastInterest = dx.Rows[j]["interestPaid"].ObjToDouble();
                                interest = dx.Rows[j]["int"].ObjToDouble();
                                if (interest != pastInterest)
                                {
                                    if (interest == 0D || pastInterest == 0D)
                                    {
                                        difference = interest - pastInterest;
                                        balanceDue += difference;
                                        balanceDue = G1.RoundValue(balanceDue);
                                        gridMain.Columns["fix"].Visible = true;
                                        pastRecord = dx.Rows[j]["record"].ObjToString();
                                        myFields = "interestPaid," + interest.ToString();
                                        dt.Rows[i]["fix"] = "FIXED";
                                        G1.update_db_table("payments", "record", pastRecord, myFields);
                                        myFields = "balanceDue," + balanceDue.ToString();
                                        G1.update_db_table("contracts", "record", contractRecord, myFields);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItemSetLapsed_Click(object sender, EventArgs e)
        {
            DataTable dt = null;
            DataRow dr = null;
            int rowHandle = 0;
            int row = 0;
            string record = "";
            string contractNumber = "";
            if (dgv.Visible)
            {
                dt = (DataTable)dgv.DataSource;
                dr = gridMain.GetFocusedDataRow();
                rowHandle = gridMain.FocusedRowHandle;
                row = gridMain.GetDataSourceRowIndex(rowHandle);
                record = dr["record"].ObjToString();
                contractNumber = dr["contractNumber"].ObjToString();
            }
            else
            {
                dt = (DataTable)dgv2.DataSource;
                dr = gridMain2.GetFocusedDataRow();
                rowHandle = gridMain2.FocusedRowHandle;
                row = gridMain2.GetDataSourceRowIndex(rowHandle);
                record = dr["record"].ObjToString();
                contractNumber = dr["contractNumber"].ObjToString();
            }
            DialogResult result = MessageBox.Show("Are you sure you want to Set Lapsed for customer (" + contractNumber + ") ?", "Set Lapsed Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (result == DialogResult.No)
                return;

            using (Ask askForm = new Ask("Enter Date to Assign for Lapse . . .", LapsedDate))
            {
                askForm.Text = "";
                askForm.ShowDialog();
                if (askForm.DialogResult != System.Windows.Forms.DialogResult.OK)
                    return;
                string date = askForm.Answer;

                if (String.IsNullOrWhiteSpace(date))
                    return;
                if ( !G1.validate_date ( date))
                {
                    MessageBox.Show("***ERROR*** Not a valid Lapse Date!");
                    return;
                }
                LapsedDate = date;
            }

            if ( DailyHistory.SetLapsed(contractNumber, LapsedDate) )
            {
                DateTime myDate = LapsedDate.ObjToDateTime();
                dt.Rows[row]["lapseDate8"] = G1.DTtoMySQLDT(myDate);
                dt.Rows[row]["lapsed"] = "Y";
                if ( dgv.Visible )
                    dt.Rows[row]["lapsed1"] = "Y";
                dt.AcceptChanges();
                dr["lapsed"] = "YY";
                dr["lapseDate8"] = G1.DTtoMySQLDT(myDate);
                if (dgv.Visible)
                {
                    gridMain.RefreshData();
                    dgv.Refresh();
                }
                else
                {
                    gridMain2.RefreshData();
                    dgv2.Refresh();
                }
            }
        }
        /***********************************************************************************************/
        private void toolStripTrustReports_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            TrustReports trustForm = new TrustReports(dt, dt, DateTime.Now);
            trustForm.Show();
        }
        /***********************************************************************************************/
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EditHelp helpForm = new EditHelp("Customers");
            helpForm.Show();
        }
        /***********************************************************************************************/
        private void menuDeceasedReport_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Deceased Report", "Deceased Report");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void lockBoxDepositReportMenu_Click(object sender, EventArgs e)
        {
            LockBoxDeposits lockBoxForm = new LockBoxDeposits();
            lockBoxForm.Show();
        }
        /***********************************************************************************************/
        private void debitsAndCreditsMenu_Click(object sender, EventArgs e)
        {
            DebitsAndCredits dcForm = new DebitsAndCredits();
            dcForm.Show();
        }
        /***********************************************************************************************/
        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            PaymentsReport paymentForm = new PaymentsReport("Lapse Report", "Lapse Report (5.0)");
            paymentForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void stopBreakReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StopBreakReport stopForm = new StopBreakReport();
            stopForm.Show();
        }
        /***********************************************************************************************/
        private void menuStopLoss_Click(object sender, EventArgs e)
        {
        }
        /***********************************************************************************************/
        private void menuUnityReport_Click(object sender, EventArgs e)
        {
            UnityReport unityForm = new UnityReport();
            unityForm.Show();
        }
        /***********************************************************************************************/
        private void trustSummaryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TrustSummary trustForm = new TrustSummary();
            trustForm.Show();
        }
        /***********************************************************************************************/
        private void insuranceSummaryReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsuranceSummary insuranceForm = new InsuranceSummary();
            insuranceForm.Show();
        }
        /***********************************************************************************************/
        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            LossRecovery lossForm = new LossRecovery();
            lossForm.Show();
        }
        /***********************************************************************************************/
        private void contractsLessThan2039ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TrustNonPaidOff nonForm = new TrustNonPaidOff();
            nonForm.Show();
        }
        /***********************************************************************************************/
        private void gridMain2_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            ColumnView view = sender as ColumnView;
            if (e.Column.FieldName.ToUpper().IndexOf("DATE") >= 0 && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                if (e.DisplayText.IndexOf("0000") >= 0 || e.DisplayText.IndexOf("0001") >= 0)
                    e.DisplayText = "";
                else
                {
                    DateTime date = e.DisplayText.ObjToString().ObjToDateTime();
                    if (date.Year > 500)
                        e.DisplayText = date.ToString("MM/dd/yyyy");
                    else
                        e.DisplayText = "";
                }
            }
        }
        /***********************************************************************************************/
        private void gridMain2_DoubleClick(object sender, EventArgs e)
        {
            DataRow dr = gridMain2.GetFocusedDataRow();
            string contract = dr["contractNumber"].ObjToString();
            if (!String.IsNullOrWhiteSpace(contract))
            {
                this.Cursor = Cursors.WaitCursor;
                DataTable dt = (DataTable)dgv2.DataSource;
                G1.UpdatePreviousCustomer(contract, LoginForm.username);
                CustomerDetails clientForm = new CustomerDetails(contract);
                clientForm.Show();
                this.Cursor = Cursors.Default;
            }
        }
        /***********************************************************************************************/
        private void yearendDeathReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TrustDeathReport deathForm = new TrustDeathReport();
            deathForm.Show();
        }
        /***********************************************************************************************/
        private void thirdPartyDumpReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsuranceThirdPartyDumpReport dumpForm = new InsuranceThirdPartyDumpReport();
            dumpForm.Show();
        }
        /***********************************************************************************************/
        private void menuGenerateInsuranceCoupons_Click(object sender, EventArgs e)
        {
            //if (dgv2.DataSource == null)
            //{
            //    MessageBox.Show("***ERROR*** There are no payers here");
            //    return;
            //}
            //DataTable dt = (DataTable)dgv2.DataSource;
            InsuranceCoupons insuranceForm = new InsuranceCoupons();
            insuranceForm.Show();
        }
        /***********************************************************************************************/
        private void mailingLabelsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv.DataSource == null)
            {
                DialogResult result = MessageBox.Show("Table here is Empty!\nYou must first press Refresh!", "Empty Mailing Labels Dialog", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DataTable dt = (DataTable)dgv.DataSource;
            DataView dv = new DataView( dt );
            ColumnView view = gridMain as ColumnView;

            dv.RowFilter = DevExpress.Data.Filtering.CriteriaToWhereClauseHelper.GetDataSetWhere(view.ActiveFilterCriteria);
            if (dv.ToTable().Rows.Count > 0)
                dt = dv.ToTable();

            MailingLabels mailForm = new MailingLabels(dt);
            mailForm.Show();
        }
        /***********************************************************************************************/
        private void mailingLabelsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dgv2.DataSource == null)
            {
                DialogResult result = MessageBox.Show("Table here is Empty!\nYou must first press Refresh!", "Empty Mailing Labels Dialog", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DataTable dt = (DataTable)dgv2.DataSource;
            DataView dv = new DataView(dt);
            ColumnView view = gridMain2 as ColumnView;

            dv.RowFilter = DevExpress.Data.Filtering.CriteriaToWhereClauseHelper.GetDataSetWhere(view.ActiveFilterCriteria);
            if (dv.ToTable().Rows.Count > 0)
                dt = dv.ToTable();

            MailingLabels mailForm = new MailingLabels(dt, true );
            mailForm.Show();
        }
        /***********************************************************************************************/
        private void menuSecurityNational_Click(object sender, EventArgs e)
        {
            SecurityNationalReport secureReport = new SecurityNationalReport();
            secureReport.Show();
        }
        /***********************************************************************************************/
        private void menuEarlySecNatPayments_Click(object sender, EventArgs e)
        {
            SecNatPayments secForm = new SecNatPayments();
            secForm.Show();
        }
        /***********************************************************************************************/
        private void bankPaymentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BankPayments bankForm = new BankPayments();
            bankForm.Show();
        }
        /***********************************************************************************************/
        private void editTrustDownPaymentsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DownPayments downForm = new DownPayments( "Enter Down Payments and/or Payments");
            downForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void verifyLocationPaymentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DownPayments downForm = new DownPayments("Verify Down Payments and/or Payments", true );
            downForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void menuSecNatNotices_Click(object sender, EventArgs e)
        {
            SecurityNationalNotices secForm = new SecurityNationalNotices();
            secForm.Show();
        }
        /***********************************************************************************************/
        private void findConflictingDeceasedPoliciesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PolicyDuplicates dupForm = new PolicyDuplicates();
            dupForm.Show();
        }
        /***********************************************************************************************/
        private void findProblemInsurancePaymentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsurancePaymentProblems insuranceForm = new InsurancePaymentProblems();
            insuranceForm.Show();
        }
        /***********************************************************************************************/
        private void gridMain_CustomColumnDisplayText(object sender, CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.ToUpper().IndexOf("DATE") >= 0 && e.ListSourceRowIndex != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                if (e.DisplayText.IndexOf("0000") >= 0 || e.DisplayText.IndexOf("0001") >= 0)
                    e.DisplayText = "";
                else
                {
                    DateTime date = e.DisplayText.ObjToString().ObjToDateTime();
                    e.DisplayText = date.ToString("yyyy-MM-dd");
                    if (date.Year < 30)
                        e.DisplayText = "";
                }
            }
        }
        /***********************************************************************************************/
        private void Customers_FormClosing(object sender, FormClosingEventArgs e)
        {
            G1.CleanupDataGrid(ref dgv);
            G1.CleanupDataGrid(ref dgv2);
            GC.Collect();
        }
        /***********************************************************************************************/
        public static int GetNumberOfPayments ( string contractNumber, double amtOfMonthlyPayt )
        {
            int numPayments = 0;
            double payment = 0D;
            double credit = 0D;
            double debit = 0D;
            string cmd = "Select * from `payments` where `contractNumber` = '" + contractNumber + "' order by `payDate8` DESC, `tmstamp` DESC;";
            DataTable dt = G1.get_db_data(cmd);
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                payment = dt.Rows[i]["paymentAmount"].ObjToDouble();
                credit = dt.Rows[i]["creditAdjustment"].ObjToDouble();
                debit = dt.Rows[i]["debitAdjustment"].ObjToDouble();
                if (payment > 0D || credit > 0D)
                    numPayments++;
                else if ( debit > 0D )
                {
                    if (debit == amtOfMonthlyPayt)
                        numPayments--;
                }
            }
            return numPayments;
        }
        /***********************************************************************************************/
        private void potentialPaidOffToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string contractNumber = "";
            double endingBalance = 0D;
            double beginningBalance = 0D;
            double contractValue = 0D;
            double amtOfMonthlyPayt = 0D;
            double maxTrust85 = 0D;
            int numPayments = 0;
            int numPaid = 0;
            double totalTrust85 = 0D;
            DateTime dailyLastPaidDate = DateTime.Now;
            DateTime lastPaidDate = DateTime.Now;
            string cmd = "SELECT * FROM contracts c JOIN customers x ON c.`contractNumber` = x.`contractNumber` WHERE dueDate8 != '2039-12-31' AND balanceDue < '25.00' AND c.`deceasedDate` < '1900-01-01' AND c.`lapsed` <> 'Y';";
            DataTable dt = G1.get_db_data(cmd);
            dt.Columns.Add("paid", Type.GetType("System.Double"));
            DataTable dx = dt.Clone();
            DataTable payDt = null;
            double Trust85Paid = 0D;
            double difference = 0D;
            DateTime payDate = DateTime.Now;
            this.Cursor = Cursors.WaitCursor;
            barImport.Minimum = 0;
            barImport.Maximum = dt.Rows.Count;
            barImport.Value = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                barImport.Value = i;
                barImport.Refresh();
                try
                {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                amtOfMonthlyPayt = dt.Rows[i]["amtOfMonthlyPayt"].ObjToDouble();
                numPayments = dt.Rows[i]["numberOfPayments"].ObjToInt32();
                //dailyLastPaidDate = dt.Rows[i]["lastDatePaid8"].ObjToDateTime();
                if ( 1 == 1)
                    {
                        bool trustThreshold = false;
                        bool balanceThreshold = false;
                        if ( CheckForcedPayoff ( contractNumber, amtOfMonthlyPayt, numPayments, ref maxTrust85, ref totalTrust85, ref contractValue, ref trustThreshold, ref balanceThreshold ))
                        {
                            dt.Rows[i]["nowDue"] = totalTrust85;
                            dt.Rows[i]["creditBalance"] = contractValue;
                            dt.Rows[i]["paid"] = maxTrust85;
                            dx.ImportRow(dt.Rows[i]);
                        }
                        continue;
                    }

                numPaid = Customers.GetNumberOfPayments(contractNumber, amtOfMonthlyPayt);
                    if (numPaid >= numPayments)
                    {
                        lastPaidDate = DailyHistory.GetTrustLastPaid(contractNumber, ref beginningBalance, ref endingBalance);
                        if (endingBalance <= 0D)
                            continue;
                        contractValue = DailyHistory.GetContractValue(contractNumber);

                        cmd = "Select * from `payments` where `contractNumber` = '" + contractNumber + "' AND `payDate8` > '" + lastPaidDate.ObjToDateTime().ToString("yyyy-MM-dd") + "' ORDER BY `payDate8` DESC;";
                        payDt = G1.get_db_data(cmd);

                        Trust85Paid = 0D;
                        for (int j = 0; j < payDt.Rows.Count; j++)
                        {
                            if (payDt.Rows[j]["fill"].ObjToString().ToUpper() == "D")
                                continue;
                            payDate = payDt.Rows[j]["payDate8"].ObjToDateTime();
                            if (payDate < lastPaidDate)
                                continue;
                            Trust85Paid += payDt.Rows[j]["trust85P"].ObjToDouble();
                        }

                        maxTrust85 = contractValue * .85D;
                        difference = (endingBalance + Trust85Paid) - maxTrust85;
                        if (difference < 5D)
                        {
                            dt.Rows[i]["nowDue"] = endingBalance + Trust85Paid;
                            dt.Rows[i]["creditBalance"] = contractValue;
                            dt.Rows[i]["paid"] = maxTrust85;
                            dx.ImportRow(dt.Rows[i]);
                        }
                    }
                }
                catch ( Exception ex)
                {
                }
            }
            barImport.Value = dt.Rows.Count;
            barImport.Refresh();
            if (dx.Rows.Count > 0)
            {
                PotentialPayoffs payoffForm = new PotentialPayoffs(dx);
                payoffForm.Show();
            }
            else
            {
                MessageBox.Show("There are no potential payoffs available!", "Potential Payoffs Dialog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        public static bool CheckForcedPayoff ( string contractNumber, double amtOfMonthlyPayt, int numPayments, ref double maxTrust85, ref double totalTrust85, ref double contractValue, ref bool trustThreshold, ref bool balanceThreshold, double addingTrust85 = 0D)
        {
            bool payOff = false;
            balanceThreshold = false;
            trustThreshold = false;
            maxTrust85 = 0D;
            totalTrust85 = 0D;
            contractValue = 0D;
            string cmd = "";
            if ( amtOfMonthlyPayt <= 0D && numPayments <= 0 )
            {
                cmd = "Select * from `contracts` where `contractNumber` = '" + contractNumber + "';";
                DataTable dt = G1.get_db_data(cmd);
                if (dt.Rows.Count <= 0)
                    return payOff;
                amtOfMonthlyPayt = dt.Rows[0]["amtOfMonthlyPayt"].ObjToDouble();
                numPayments = dt.Rows[0]["numberOfPayments"].ObjToInt32();
            }
            int numPaid = 0;
            //int numPaid = Customers.GetNumberOfPayments(contractNumber, amtOfMonthlyPayt);
            //if (addingTrust85 > 0D)
            //    numPaid++;
            numPaid = numPayments + 1; // Force this code to run
            if (numPaid >= numPayments)
            {
                double beginningBalance = 0D;
                double endingBalance = 0D;
                DateTime lastPaidDate = DailyHistory.GetTrustLastPaid(contractNumber, ref beginningBalance, ref endingBalance);
                if (endingBalance <= 0D)
                    lastPaidDate = new DateTime(2000, 1, 1);

                //endingBalance = 5606.69D; // Just for Testing

                contractValue = DailyHistory.GetContractValuePlus(contractNumber);

                cmd = "Select * from `payments` where `contractNumber` = '" + contractNumber + "' AND `payDate8` > '" + lastPaidDate.ObjToDateTime().ToString("yyyy-MM-dd") + "' ORDER BY `payDate8` DESC;";
                DataTable payDt = G1.get_db_data(cmd);

                double Trust85Paid = 0D;
                DateTime payDate = DateTime.Now;
                for (int j = 0; j < payDt.Rows.Count; j++)
                {
                    if (payDt.Rows[j]["fill"].ObjToString().ToUpper() == "D")
                        continue;
                    payDate = payDt.Rows[j]["payDate8"].ObjToDateTime();
                    if (payDate < lastPaidDate)
                        continue;
                    Trust85Paid += payDt.Rows[j]["trust85P"].ObjToDouble();
                }

                maxTrust85 = contractValue * .85D;
                totalTrust85 = endingBalance + Trust85Paid + addingTrust85;
                totalTrust85 = G1.RoundValue(totalTrust85);

                double difference = maxTrust85 - totalTrust85;
                difference = G1.RoundValue(difference);
                //difference = Math.Abs(difference);
                if (difference <= LoginForm.trust85Threshold )
                {
                    if (difference > 0D)
                    {
                        payOff = true;
                        trustThreshold = true;
                    }
                }
            }
            return payOff;
        }
        /***********************************************************************************************/
        private void deleteContractToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!LoginForm.administrator)
            {
                MessageBox.Show("***Warning*** You do not have permission to delete a contract!");
                return;
            }

            string customerRecord = "";
            string contractRecord = "";

            string contract = "";

            DataRow dr = gridMain.GetFocusedDataRow();
            if ( dr != null )
                contract = dr["contractNumber"].ObjToString();
            if (String.IsNullOrWhiteSpace(contract))
            {
                using (Ask askForm = new Ask("Enter Contract # to DELETE ?"))
                {
                    askForm.Text = "";
                    askForm.ShowDialog();
                    if (askForm.DialogResult != System.Windows.Forms.DialogResult.OK)
                        return;
                    contract = askForm.Answer;
                    if (String.IsNullOrWhiteSpace(contract))
                        return;
                }

                //MessageBox.Show("***Warning*** This contract number is empty!!");
                //return;
            }

            string contractsFile = "contracts";
            string customersFile = "customers";
            string paymentsFile = "payments";
            bool insurance = false;

            if (DailyHistory.isInsurance(contract) || cmbType.Text.ToUpper() == "INSURANCE")
            {
                contractsFile = "icontracts";
                customersFile = "icustomers";
                paymentsFile = "ipayments";
                insurance = true;
            }

            DialogResult result = MessageBox.Show("***Question*** Are you sure you want to DELETE this Contract (" + contract + ") ?", "Delete Contract Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            if (result == DialogResult.No)
                return;

            string cmd = "Select * from `" + customersFile + "` where `contractNumber` = '" + contract + "';";
            DataTable dt = G1.get_db_data(cmd);
            if (dt.Rows.Count <= 0)
                MessageBox.Show("***ERROR*** No Customers exist for Contract " + contract + "!");
            else
                customerRecord = dt.Rows[0]["record"].ObjToString();

            cmd = "Select * from `" + contractsFile + "` where `contractNumber` = '" + contract + "';";
            dt = G1.get_db_data(cmd);
            if (dt.Rows.Count <= 0)
                MessageBox.Show("***ERROR*** No Contracts exist for Contract " + contract + "!");
            else
                contractRecord = dt.Rows[0]["record"].ObjToString();

            if (!String.IsNullOrWhiteSpace(contractRecord))
            {
                G1.get_db_data("Delete from `" + contractsFile + "` where `record` = '" + contractRecord + "';");
                G1.get_db_data("Delete from `" + paymentsFile + "` where `contractNumber` = '" + contract + "';");
            }

            if (!String.IsNullOrWhiteSpace(customerRecord))
            {
                G1.get_db_data("Delete from `" + customersFile + "` where `record` = '" + customerRecord + "';");
                G1.get_db_data("Delete from `cust_services` where `contractNumber` = '" + contract + "';");
            }
            if (insurance)
                G1.AddToAudit(LoginForm.username, "Customers", "Delete Insurance Customer", "Deleted", contract);
            else
                G1.AddToAudit(LoginForm.username, "Customers", "Delete Trust Customer", "Deleted", contract);

            MessageBox.Show("***INFO*** Contract (" + contract + ") Deleted!", "Delete Contract Dialog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

        }
        /***********************************************************************************************/
        private void SetupMenuPreferences ()
        {
            if (LoginForm.isRobby)
                return;
            string preference = G1.getPreference(LoginForm.username, "Trust85 Report", "Allow Access");
            if (preference != "YES")
                trust85ReportToolStripMenuItem.Enabled = false;

            preference = G1.getPreference(LoginForm.username, "Payment Reports", "Allow Access");
            if (preference != "YES")
                toolStripMenuItem6.Enabled = false;

            preference = G1.getPreference(LoginForm.username, "Historic Commissions", "Allow Access");
            if (preference != "YES")
                historicCommissionsToolStripMenuItem.Enabled = false;

            preference = G1.getPreference(LoginForm.username, "New Business Report", "Allow Access");
            if (preference != "YES")
            {
                toolStripMenuItem17.Enabled = false;
                byAgentMenu.Enabled = false;
                byDetailToolStripMenuItem.Enabled = false;
            }

            preference = G1.getPreference(LoginForm.username, "Fee Allocation Report", "Allow Access");
            if (preference != "YES")
                feeAllocationReportToolStripMenuItem.Enabled = false;

            preference = G1.getPreference(LoginForm.username, "Miscellaneous Reports", "Allow Access");
            if (preference != "YES")
                miscellaneousReportsMenu.Enabled = false;

            preference = G1.getPreference(LoginForm.username, "Insurance Reports", "Allow Access");
            if (preference != "YES")
            {
                insuranceReportsToolStripMenuItem.HideDropDown();
                insuranceReportsToolStripMenuItem.Enabled = false;
                insuranceReportsToolStripMenuItem.Visible = false;
            }
            if ( G1.isField())
            {
                columnsToolStripMenuItem.Visible = false;
                reportsToolStripMenuItem.Visible = false;
                insuranceReportsToolStripMenuItem.Visible = false;
                editToolStripMenuItem.Visible = false;
                commonToolStripMenuItem.Visible = false;

                yearendReportsToolStripMenuItem.Visible = false;
                potentialPaidOffToolStripMenuItem.Visible = false;

                SetFieldUserFormat();

                btnRecalc.Hide();
                chkMismatches.Hide();
                chkMismatchDates.Hide();
                txtThreshold.Hide();
                barImport.Hide();
                label1.Hide();
                cmbSelectColumns.Hide();
                btnSelectColumns.Hide();

                chkIncludePaid.Hide();
                chkLessZero.Hide();
                chkSortLastName.Hide();
            }
        }
        /****************************************************************************************/
        private void SetFieldUserFormat()
        {
            string type = cmbType.Text.ToUpper();
            if ( type == "TRUSTS" || type == "FUNERALS" || type == "SINGLE PREMIUM" )
            {
                ClearAllPositions(gridMain);

                G1.AddNewColumn(gridMain, "address1", "Address1", "", FormatType.None, 175, true);
                G1.SetColumnWidth(gridMain, "address1", 175);
                G1.AddNewColumn(gridMain, "address2", "Address2", "", FormatType.None, 125, true);
                G1.SetColumnWidth(gridMain, "address2", 125);
                G1.AddNewColumn(gridMain, "city", "City", "", FormatType.None, 125, true);
                G1.SetColumnWidth(gridMain, "city", 125);
                G1.AddNewColumn(gridMain, "state", "State", "", FormatType.None, 125, true);
                G1.SetColumnWidth(gridMain, "state", 125);
                G1.AddNewColumn(gridMain, "zip1", "Zip", "", FormatType.None, 125, true);
                G1.SetColumnWidth(gridMain, "zip1", 125);

                int i = 1;
                G1.SetColumnPosition(gridMain, "num", ++i);
                G1.SetColumnPosition(gridMain, "contractNumber", ++i);
                G1.SetColumnPosition(gridMain, "ServiceId1", ++i);
                G1.SetColumnPosition(gridMain, "lapsed", ++i);
                //G1.SetColumnPosition(gridMain, "fullname", ++i);
                G1.SetColumnPosition(gridMain, "lastName", ++i);
                G1.SetColumnPosition(gridMain, "firstName", ++i);
                G1.SetColumnPosition(gridMain, "ssno", ++i);
                G1.SetColumnPosition(gridMain, "birthDate", ++i);
                G1.SetColumnPosition(gridMain, "issueDate8", ++i);
                G1.SetColumnPosition(gridMain, "dueDate8", ++i);
                G1.SetColumnPosition(gridMain, "deceasedDate", ++i);
                G1.SetColumnPosition(gridMain, "address1", ++i);
                G1.SetColumnPosition(gridMain, "address2", ++i);
                G1.SetColumnPosition(gridMain, "city", ++i);
                G1.SetColumnPosition(gridMain, "state", ++i);
                G1.SetColumnPosition(gridMain, "zip1", ++i);

                if (G1.isField())
                {
                    //gridMain.OptionsMenu.EnableColumnMenu = false;
                    gridMain.OptionsMenu.EnableColumnMenu = true;
                }
                else
                    gridMain.OptionsMenu.EnableColumnMenu = true;
                if (!LoginForm.administrator)
                    dgv.ContextMenuStrip = null;
            }
            else
            {
                ClearAllPositions(gridMain2);

                G1.AddNewColumn(gridMain2, "address1", "Address1", "", FormatType.None, 175, true);
                G1.SetColumnWidth(gridMain2, "address1", 175);
                G1.AddNewColumn(gridMain2, "address2", "Address2", "", FormatType.None, 125, true);
                G1.SetColumnWidth(gridMain2, "address2", 125);
                G1.AddNewColumn(gridMain2, "city", "City", "", FormatType.None, 125, true);
                G1.SetColumnWidth(gridMain2, "city", 125);
                G1.AddNewColumn(gridMain2, "state", "State", "", FormatType.None, 125, true);
                G1.SetColumnWidth(gridMain2, "state", 125);
                G1.AddNewColumn(gridMain2, "zip1", "Zip", "", FormatType.None, 125, true);
                G1.SetColumnWidth(gridMain2, "zip1", 125);

                int i = 1;
                G1.SetColumnPosition(gridMain2, "num", ++i);
                //G1.SetColumnPosition(gridMain, "contractNumber", ++i);
                G1.SetColumnPosition(gridMain2, "payer", ++i);
                G1.SetColumnPosition(gridMain2, "lapsed", ++i);
                //G1.SetColumnPosition(gridMain, "fullname", ++i);
                G1.SetColumnPosition(gridMain2, "lastName", ++i);
                G1.SetColumnPosition(gridMain2, "firstName", ++i);
                G1.SetColumnPosition(gridMain2, "ssno", ++i);
                G1.SetColumnPosition(gridMain2, "amtOfMonthlyPayt", ++i);
                //G1.SetColumnPosition(gridMain2, "annualPremium", ++i);
                G1.SetColumnPosition(gridMain2, "birthDate", ++i);
                G1.SetColumnPosition(gridMain2, "issueDate8", ++i);
                G1.SetColumnPosition(gridMain2, "dueDate8", ++i);
                G1.SetColumnPosition(gridMain2, "lapseDate8", ++i);
                G1.SetColumnPosition(gridMain2, "deceasedDate", ++i);
                G1.SetColumnPosition(gridMain2, "reinstateDate8", ++i);
                G1.SetColumnPosition(gridMain2, "address1", ++i);
                G1.SetColumnPosition(gridMain2, "address2", ++i);
                G1.SetColumnPosition(gridMain2, "city", ++i);
                G1.SetColumnPosition(gridMain2, "state", ++i);
                G1.SetColumnPosition(gridMain2, "zip1", ++i);

                if ( G1.isField () )
                    gridMain2.OptionsMenu.EnableColumnMenu = true;
                else
                    gridMain2.OptionsMenu.EnableColumnMenu = true;
                if (!LoginForm.administrator)
                    dgv2.ContextMenuStrip = null;
            }

            SetupEdit(chkFullEdit.Checked );

        }
        /****************************************************************************************/
        private void ClearAllPositions(DevExpress.XtraGrid.Views.BandedGrid.AdvBandedGridView gMain = null)
        {
            if (gMain == null)
                gMain = gridMain;
            for (int i = 0; i < gMain.Columns.Count; i++)
            {
                gMain.Columns[i].Visible = false;
            }
        }
        /***********************************************************************************************/
        private void showForcedPaidOffContractsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ForcedPayoffs payoffForm = new ForcedPayoffs();
            payoffForm.Show();
        }
        /***********************************************************************************************/
        private void verifyTrustBBWithDailyHistoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            DataRow[] dRows = dt.Select("contractNumber LIKE '%LI'");
            if (dRows.Length > 0)
                dt = dRows.CopyToDataTable();
            dt.Columns.Add("DailyHistory", Type.GetType("System.Double"));
            dt.Columns.Add("TBB", Type.GetType("System.Double"));
            dt.Columns.Add("Diff", Type.GetType("System.Double"));

            string contractNumber = "";
            string cmd = "";
            double dailyHistory = 0D;
            double tbb = 0D;
            DataTable dx = null;
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                cmd = "Select * from `trust2013r` where `contractNumber` = '" + contractNumber + "' ORDER BY `payDate8` DESC;";
                dx = G1.get_db_data(cmd);
                if (dx.Rows.Count <= 0)
                    continue;
                AddDailyHistory(contractNumber, dx);
                dailyHistory = 0D;
                tbb = 0D;
                for ( int j=0; j<dx.Rows.Count; j++)
                {
                    dailyHistory += dx.Rows[j]["dailyHistory"].ObjToDouble();
                    tbb += dx.Rows[j]["paymentCurrMonth"].ObjToDouble();
                }
                dt.Rows[i]["DailyHistory"] = dailyHistory;
                dt.Rows[i]["tbb"] = tbb;
                dt.Rows[i]["diff"] = G1.RoundValue (tbb - dailyHistory);
                //if (i >= 5 )
                    //break;
            }
            dRows = dt.Select ( "Diff <> '0'");
            if (dRows.Length > 0)
                dx = dRows.CopyToDataTable();
        }
        /****************************************************************************************/
        private void AddDailyHistory(string contractNumber, DataTable dt)
        {
            DataTable payDt = GetDailyHistory(contractNumber);
            if (G1.get_column_number(dt, "dailyHistory") < 0)
                dt.Columns.Add("dailyHistory", Type.GetType("System.Double"));
            if (G1.get_column_number(payDt, "myDate") < 0)
                payDt.Columns.Add("myDate");
            for (int i = 0; i < payDt.Rows.Count; i++)
            {
                payDt.Rows[i]["myDate"] = payDt.Rows[i]["payDate8"].ObjToDateTime().ToString("yyyyMM");
            }
            DateTime oldDate = DateTime.Now;
            DateTime date = DateTime.Now;
            DataRow[] dRows = null;
            string date1 = "";
            string date2 = "";
            int days = 0;
            double trust85 = 0D;
            string fill = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                date = dt.Rows[i]["payDate8"].ObjToDateTime();
                oldDate = new DateTime(date.Year, date.Month, 1);
                date1 = oldDate.ToString("yyyyMM");
                dRows = payDt.Select("myDate='" + date1 + "' AND fill<>'D'");
                if (dRows.Length > 0)
                {
                    trust85 = 0D;
                    for (int j = 0; j < dRows.Length; j++)
                        trust85 += dRows[j]["trust85P"].ObjToDouble();
                    dt.Rows[i]["dailyHistory"] = trust85;
                }
                else
                    dt.Rows[i]["dailyHistory"] = 0D;
            }
        }
        /****************************************************************************************/
        private DataTable GetDailyHistory (string contractNumber )
        {
            string cmd = "Select * from `contracts` where `contractNumber` = '" + contractNumber + "';";
            DataTable dx = G1.get_db_data(cmd);
            if (dx.Rows.Count <= 0)
                return null;

            DataTable contractDt = dx.Copy();

            double payment = dx.Rows[0]["amtOfMonthlyPayt"].ObjToString().ObjToDouble();
            int numPayments = dx.Rows[0]["numberOfPayments"].ObjToString().ObjToInt32();
            double totalInterest = dx.Rows[0]["totalInterest"].ObjToString().ObjToDouble();
            string dueDate = dx.Rows[0]["dueDate8"].ObjToString();
            string issueDate = dx.Rows[0]["issueDate8"].ObjToString();
            DateTime iDate = DailyHistory.GetIssueDate(dx.Rows[0]["issueDate8"].ObjToDateTime(), contractNumber, dx);
            issueDate = iDate.ToString("MM/dd/yyyy");
            DateTime lastDate = issueDate.ObjToDateTime();
            string apr = dx.Rows[0]["APR"].ObjToString();
            double dAPR = apr.ObjToDouble() / 100.0D;
            double contractValue = DailyHistory.GetContractValue(dx.Rows[0]);
            //Trust85Max = contractValue * 0.85D;
            //Trust85Max = G1.RoundValue(Trust85Max);

            cmd = "Select * from `payments` where `contractNumber` = '" + contractNumber + "' order by `paydate8` DESC, `tmstamp` DESC;";
            dx = G1.get_db_data(cmd);

            if (numPayments <= 0 && dx.Rows.Count > 0)
                numPayments = dx.Rows.Count;
            double startBalance = DailyHistory.GetFinanceValue(contractNumber);

            DailyHistory.CalculateNewStuff(dx, dAPR, numPayments, startBalance, lastDate);

            //if (dx.Rows.Count > 0)
            //    balanceDue = dx.Rows[0]["newBalance"].ObjToDouble();
            //for (int i = 0; i < dx.Rows.Count; i++)
            //{
            //    iDate = dx.Rows[i]["payDate8"].ObjToDateTime();
            //    if (iDate >= workDate2)
            //        continue;
            //    Trust85Real += dx.Rows[i]["calculatedTrust85"].ObjToDouble();
            //    balanceDue = dx.Rows[i]["newBalance"].ObjToDouble();
            //}
            return dx;
        }
        /***********************************************************************************************/
        private void byAgentMenu_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            NewByAgent newForm = new NewByAgent();
            newForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void byDetailToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            NewByDetail newForm = new NewByDetail();
            newForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            NewBusiness formNewBusiness = new NewBusiness();
            formNewBusiness.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem18_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            NewByLocation newForm = new NewByLocation();
            newForm.Show();
            this.Cursor = Cursors.Default;
            //MessageBox.Show("***INFO*** Report Not Available Yet!", "New Business by Location Report Dialog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }
        /***********************************************************************************************/
        private void findMismatchedPayersPoliciesToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }
        /***********************************************************************************************/
        private void chkDownPayment_CheckedChanged(object sender, EventArgs e)
        {
            string contractNumber = "";
            DataTable dx = null;
            string cmd = "";
            DataTable dt = (DataTable)dgv.DataSource;
            if ( G1.get_column_number ( dt, "dp") < 0 )
                dt.Columns.Add("dp", Type.GetType("System.Double"));

            double contractValue = 0D;
            double downPayment = 0D;
            double dp = 0D;

            this.Cursor = Cursors.WaitCursor;

            for ( int i=0; i<dt.Rows.Count; i++)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                if (String.IsNullOrWhiteSpace(contractNumber))
                    continue;
                downPayment = dt.Rows[i]["downPayment"].ObjToDouble();
                contractValue = dt.Rows[i]["contractValue"].ObjToDouble();
                cmd = "Select * from `payments` where `contractNumber` = '" + contractNumber + "' AND `downPayment` > '0';";
                dx = G1.get_db_data(cmd);
                if ( dx.Rows.Count > 0 )
                {
                    dp = dx.Rows[0]["downPayment"].ObjToDouble();
                    if (dp == 0D && downPayment == contractValue)
                        dt.Rows[i]["dp"] = 999D;
                    else
                        dt.Rows[i]["dp"] = dp;
                }
                else
                {
                    if ( downPayment == contractValue)
                        dt.Rows[i]["dp"] = 999D;
                }
            }
            dgv.DataSource = dt;
            gridMain.RefreshData();
            dgv.Refresh();

            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void activeContractsByLocationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ActiveContracts activeForm = new ActiveContracts();
            activeForm.Show();
        }
        /***********************************************************************************************/
        private void monthlyUnityProcessMenu_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            Unity unityForm = new Unity();
            unityForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void trustTotalsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            TrustTotals trustForm = new TrustTotals();
            trustForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void trustCompanyDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            TrustData trustForm = new TrustData();
            trustForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void trustInterest15ReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            TrustInterestReport trustForm = new TrustInterestReport();
            trustForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void trustContractsWODemographicsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //string cmd = "SELECT * FROM `contracts` WHERE NOT EXISTS (SELECT * FROM `customers` b WHERE b.`contractNumber` = contracts.`contractNumber`) ORDER BY contracts.`deceasedDate`;";
            string cmd = "SELECT * FROM `contracts` WHERE NOT EXISTS (SELECT * FROM `customers` b WHERE b.`contractNumber` = contracts.`contractNumber`) ORDER BY contracts.`deceasedDate`;";
            //string cmd = "SELECT * FROM `trust_data` WHERE NOT EXISTS (SELECT * FROM `contracts` b WHERE b.`contractNumber` = contracts.`contractNumber`) ORDER BY contracts.`deceasedDate`;";
            //SELECT A.ABC_ID, A.VAL FROM A WHERE NOT EXISTS (SELECT* FROM B WHERE B.ABC_ID = A.ABC_ID AND B.VAL = A.VAL)
            DataTable dt = G1.get_db_data(cmd);
            dt.Columns.Add("num");
            dt.Columns.Add("bDate");
            dt.Columns.Add("ssno");
            dt.Columns.Add("agreement");
            dt.Columns.Add("select");
            dt.Columns.Add("fullname");
            dt.Columns.Add("paid", Type.GetType("System.Double"));
            dt.Columns.Add("purchase", Type.GetType("System.Double"));
            dt.Columns.Add("dueDate");
            dt.Columns.Add("DOLP");
            dt.Columns.Add("cbal", Type.GetType("System.Double"));
            dt.Columns.Add("newduedate");
            dt.Columns.Add("days", Type.GetType("System.Int32"));
            dt.Columns.Add("cint", Type.GetType("System.Double"));
            dt.Columns.Add("financed", Type.GetType("System.Double"));
            dt.Columns.Add("trust85", Type.GetType("System.Double"));
            dt.Columns.Add("contractValue", Type.GetType("System.Double"));
            dt.Columns.Add("percentPaid", Type.GetType("System.Double"));
            dt.Columns.Add("idate");
            dt.Columns.Add("realDOLP");
            dt.Columns.Add("DDATE");

            G1.NumberDataTable( dt );
            dgv.DataSource = dt;

            btnRefresh.Text = "Fix";
            btnRefresh.BackColor = Color.Green;
        }
        /***********************************************************************************************/
        private void FixData()
        {
            DataTable dt = (DataTable)dgv.DataSource;
            if (G1.get_column_number(dt, "firstName") < 0)
                dt.Columns.Add("firstName");
            if (G1.get_column_number(dt, "lastName") < 0)
                dt.Columns.Add("lastName");
            if (G1.get_column_number(dt, "address1") < 0)
                dt.Columns.Add("address1");
            if (G1.get_column_number(dt, "city") < 0)
                dt.Columns.Add("city");
            if (G1.get_column_number(dt, "state") < 0)
                dt.Columns.Add("state");
            if (G1.get_column_number(dt, "zip1") < 0)
                dt.Columns.Add("zip1");



            barImport.Show();
            barImport.Refresh();

            string contractNumber = "";
            string cmd = "";
            DataTable dx = null;
            int count = 0;

            barImport.Minimum = 0;
            barImport.Maximum = dt.Rows.Count;
            barImport.Value = 0;

            for ( int i=0; i<dt.Rows.Count; i++)
            {
                Application.DoEvents();
                barImport.Value = (i + 1);
                barImport.Refresh();

                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                if (String.IsNullOrWhiteSpace(contractNumber))
                    continue;
                cmd = "Select * from `trust2013r` WHERE `contractNumber` = '" + contractNumber + "' ORDER by `payDate8` DESC limit 1;";
                dx = G1.get_db_data(cmd);
                if ( dx.Rows.Count > 0 )
                {
                    dt.Rows[i]["firstName"] = dx.Rows[0]["firstName"].ObjToString();
                    dt.Rows[i]["lastName"] = dx.Rows[0]["lastName"].ObjToString();
                    dt.Rows[i]["ssno"] = dx.Rows[0]["ssn2013"].ObjToString();
                    dt.Rows[i]["address1"] = dx.Rows[0]["address2013"].ObjToString();
                    dt.Rows[i]["city"] = dx.Rows[0]["city2013"].ObjToString();
                    dt.Rows[i]["state"] = dx.Rows[0]["state2013"].ObjToString();
                    dt.Rows[i]["zip1"] = dx.Rows[0]["zip2013"].ObjToString();
                    dt.Rows[i]["fullname"] = dx.Rows[0]["firstName"].ObjToString() + " " + dx.Rows[0]["lastName"].ObjToString();
                    count++;
                }
            }
            dgv.DataSource = dt;
            gridMain.RefreshData();
            gridMain.RefreshEditor(true);

            btnRefresh.Text = "Create Customers";
            btnRefresh.Refresh();

            MessageBox.Show("*** INFO *** " + count.ToString() + " Found out of " + dt.Rows.Count.ToString() + "!", "Fix Blank Customers Dialog", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }
        /***********************************************************************************************/
        private void CreateTheData ()
        {
            DataTable dt = (DataTable)dgv.DataSource;
            string contractNumber = "";
            string cmd = "";
            DataTable dx = null;
            string firstName = "";
            string lastName = "";
            string record = "";
            int count = 0;

            barImport.Show();
            barImport.Refresh();

            barImport.Minimum = 0;
            barImport.Maximum = dt.Rows.Count;
            barImport.Value = 0;

            for ( int i=0; i<dt.Rows.Count; i++)
            {
                Application.DoEvents();
                barImport.Value = (i + 1);
                barImport.Refresh();

                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                if (String.IsNullOrWhiteSpace(contractNumber))
                    continue;

                firstName = dt.Rows[i]["firstName"].ObjToString();
                lastName = dt.Rows[i]["lastName"].ObjToString();
                if (String.IsNullOrWhiteSpace(lastName))
                    continue;

                cmd = "Select * from `customers` WHERE `contractNumber` = '" + contractNumber + "';";
                dx = G1.get_db_data(cmd);
                if (dx.Rows.Count > 0)
                    continue;

                record = G1.create_record("customers", "contractNumber", contractNumber);
                if (G1.BadRecord("customers", record))
                    continue;

                dx = G1.get_db_data(cmd);
                if (dx.Rows.Count <= 0)
                    continue;

                G1.update_db_table("customers", "record", record, new string[] { "firstName", firstName, "lastName", lastName, "contractNumber", contractNumber });
                count++;
            }

            MessageBox.Show("*** INFO *** " + count.ToString() + " Created !", "Customers Created Dialog", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }
        /***********************************************************************************************/
        private void verificationModuleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Verifications verifyForm = new Verifications();
            verifyForm.Show();
        }
        /***********************************************************************************************/
        private void editAgentMeetingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            AgentMeetings meetingForm = new AgentMeetings();
            meetingForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private DXMenuItem GetItemByStringId(DXPopupMenu menu, GridStringId id )
        {
            foreach (DXMenuItem item in menu.Items)
                if (item.Caption == GridLocalizer.Active.GetLocalizedString(id))
                    return item;
            return null;
        }
        /***********************************************************************************************/
        private DXMenuItem GetItemByName(DXPopupMenu menu, string name )
        {
            foreach (DXMenuItem item in menu.Items)
                if (item.Caption.ToUpper() == name)
                    return item;
            return null;
        }
        /***********************************************************************************************/
        private bool RemoveItemByName(DXPopupMenu menu, string name)
        {
            bool found = false;
            for (int i = menu.Items.Count - 1; i >= 0; i--)
            {
                if ( menu.Items[i].Caption.ToUpper() == name )
                {
                    menu.Items.RemoveAt(i);
                    found = true;
                    break;
                }
            }
            return found;
        }
        /***********************************************************************************************/
        private void gridMain_PopupMenuShowing(object sender, DevExpress.XtraGrid.Views.Grid.PopupMenuShowingEventArgs e)
        {
            if (!G1.isField())
                return;
            if (e.MenuType == GridMenuType.Column)
            {
                // Customize  
                DXMenuItem miCustomize = GetItemByStringId(e.Menu, GridStringId.MenuColumnColumnCustomization);
                if (miCustomize != null)
                    miCustomize.Visible = false;

                // Group By This Column  
                DXMenuItem miGroup = GetItemByStringId(e.Menu, GridStringId.MenuColumnGroup);
                if (miGroup != null)
                    miGroup.Enabled = false;

                RemoveItemByName(e.Menu, "COLUMN/BAND CHOOSER");
                RemoveItemByName(e.Menu, "SHOW AUTO FILTER ROW");
                RemoveItemByName(e.Menu, "SHOW GROUP BY BOX");
                RemoveItemByName(e.Menu, "BEST FIT (ALL COLUMNS)");
                RemoveItemByName(e.Menu, "HIDE THIS COLUMN");
                RemoveItemByName(e.Menu, "GROUP BY THIS COLUMN");
                RemoveItemByName(e.Menu, "BEST FIT");

                //DXMenuItem miGroup2 = GetItemByName(e.Menu, "COLUMN/BAND CHOOSER" );
                //if (miGroup2 != null)
                //{
                //    miGroup2.Enabled = false;
                //}

                //DXMenuItem miGroup3 = GetItemByName(e.Menu, "SHOW AUTO FILTER ROW");
                //if (miGroup3 != null)
                //    miGroup3.Enabled = false;

                //miGroup3 = GetItemByName(e.Menu, "SHOW GROUP BY BOX");
                //if (miGroup3 != null)
                //    miGroup3.Enabled = false;

                //miGroup3 = GetItemByName(e.Menu, "BEST FIT (ALL COLUMNS)");
                //if (miGroup3 != null)
                //    miGroup3.Enabled = false;

                //miGroup3 = GetItemByName(e.Menu, "HIDE THIS COLUMN");
                //if (miGroup3 != null)
                //    miGroup3.Enabled = false;
            }
        }
        /***********************************************************************************************/
        private void gridMain2_PopupMenuShowing(object sender, PopupMenuShowingEventArgs e)
        {
            if (!G1.isField())
                return;
            if (e.MenuType == GridMenuType.Column)
            {
                // Customize  
                DXMenuItem miCustomize = GetItemByStringId(e.Menu, GridStringId.MenuColumnColumnCustomization);
                if (miCustomize != null)
                    miCustomize.Visible = false;

                // Group By This Column  
                DXMenuItem miGroup = GetItemByStringId(e.Menu, GridStringId.MenuColumnGroup);
                if (miGroup != null)
                    miGroup.Enabled = false;

                RemoveItemByName(e.Menu, "COLUMN/BAND CHOOSER");
                RemoveItemByName(e.Menu, "SHOW AUTO FILTER ROW");
                RemoveItemByName(e.Menu, "SHOW GROUP BY BOX");
                RemoveItemByName(e.Menu, "BEST FIT (ALL COLUMNS)");
                RemoveItemByName(e.Menu, "HIDE THIS COLUMN");
                RemoveItemByName(e.Menu, "GROUP BY THIS COLUMN");
                RemoveItemByName(e.Menu, "BEST FIT");
            }
        }
        /***********************************************************************************************/
        private void chkShowTEB_CheckedChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.CheckBox cBox = (System.Windows.Forms.CheckBox)sender;
            if (!cBox.Checked)
            {
                gridMain.Columns["cbal"].Visible = false;
                gridMain.RefreshEditor(true);
                return;
            }

            DataTable dt = (DataTable)dgv.DataSource;
            if (dt == null)
                return;

            if ( G1.get_column_number ( dt, "cbal") < 0 )
                dt.Columns.Add("cbal", Type.GetType("System.Double"));
            gridMain.Columns["cbal"].Visible = true;
            gridMain.Columns["cbal"].Caption = "TEB";
            gridMain.RefreshEditor(true);

            string contractNumber = "";
            string cmd = "";
            DataTable dx = null;
            double teb = 0D;

            this.Cursor = Cursors.WaitCursor;

            for ( int i=0; i<dt.Rows.Count; i++)
            {
                contractNumber = dt.Rows[i]["contractNumber"].ObjToString();
                cmd = "Select * from `trust2013r` where `contractNumber` = '" + contractNumber + "' ORDER BY `payDate8` DESC;";
                dx = G1.get_db_data(cmd);
                if (dx.Rows.Count <= 0)
                    continue;
                teb = dx.Rows[0]["endingBalance"].ObjToDouble();
                dt.Rows[i]["cbal"] = teb;
                //if (teb > 0D)
                //    return;
            }

            gridMain.RefreshEditor(true);

            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void toolStripMenuItem19_Click(object sender, EventArgs e)
        {
            TrustDeceased trustForm = new TrustDeceased();
            trustForm.Show();
        }
        /***********************************************************************************************/
        private void trustContractLogReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            TrustContractLogReport logReport = new TrustContractLogReport();
            logReport.Show();

            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void chkFullEdit_CheckedChanged(object sender, EventArgs e)
        {
            bool fullEdit = false;
            if (chkFullEdit.Checked)
                fullEdit = true;

            SetupEdit(fullEdit);
        }
        /***********************************************************************************************/
        private void SetupEdit ( bool allow )
        {
            bool fullEdit = allow;

            gridMain.Columns["firstName"].OptionsColumn.AllowEdit = fullEdit;
            gridMain.Columns["lastName"].OptionsColumn.AllowEdit = fullEdit;
            if (G1.get_column_number(gridMain, "address1") > 0)
            {
                gridMain.Columns["address1"].OptionsColumn.AllowEdit = fullEdit;
                gridMain.Columns["address2"].OptionsColumn.AllowEdit = fullEdit;
                gridMain.Columns["city"].OptionsColumn.AllowEdit = fullEdit;
                gridMain.Columns["state"].OptionsColumn.AllowEdit = fullEdit;
                gridMain.Columns["zip1"].OptionsColumn.AllowEdit = fullEdit;
            }

            gridMain2.Columns["firstName"].OptionsColumn.AllowEdit = fullEdit;
            gridMain2.Columns["lastName"].OptionsColumn.AllowEdit = fullEdit;
            if (G1.get_column_number(gridMain2, "address1") > 0)
            {
                gridMain2.Columns["address1"].OptionsColumn.AllowEdit = fullEdit;
                gridMain2.Columns["address2"].OptionsColumn.AllowEdit = fullEdit;
                gridMain2.Columns["city"].OptionsColumn.AllowEdit = fullEdit;
                gridMain2.Columns["state"].OptionsColumn.AllowEdit = fullEdit;
                gridMain2.Columns["zip1"].OptionsColumn.AllowEdit = fullEdit;
            }
        }
        /***********************************************************************************************/
        private void btnSave_Click(object sender, EventArgs e)
        {
            string type = cmbType.Text;
            DataTable dt = (DataTable)dgv.DataSource;
            if ( modTrusts )
            {
                string firstName = "";
                string lastName = "";
                string address1 = "";
                string address2 = "";
                string city = "";
                string state = "";
                string zip = "";
                string record = "";
                string mod = "";
                for ( int i=0; i<dt.Rows.Count; i++)
                {
                    try
                    {
                        mod = dt.Rows[i]["mod"].ObjToString();
                        if (mod.ToUpper() != "Y")
                            continue;
                        firstName = dt.Rows[i]["firstName"].ObjToString();
                        lastName = dt.Rows[i]["lastName"].ObjToString();
                        address1 = dt.Rows[i]["address1"].ObjToString();
                        address2 = dt.Rows[i]["address2"].ObjToString();
                        city = dt.Rows[i]["city"].ObjToString();
                        state = dt.Rows[i]["state"].ObjToString();
                        zip = dt.Rows[i]["zip1"].ObjToString();

                        record = dt.Rows[i]["record"].ObjToString();
                        G1.update_db_table("customers", "record", record, new string[] { "firstName", firstName, "lastName", lastName, "address1", address1, "address2", address2, "city", city, "state", state, "zip1", zip });
                    }
                    catch ( Exception ex)
                    {
                    }
                }
            }

            if (modInsurance)
            {
                dt = (DataTable)dgv2.DataSource;

                string firstName = "";
                string lastName = "";
                string address1 = "";
                string address2 = "";
                string city = "";
                string state = "";
                string zip = "";
                string record = "";
                string mod = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        mod = dt.Rows[i]["mod"].ObjToString();
                        if (mod.ToUpper() != "Y")
                            continue;
                        firstName = dt.Rows[i]["firstName"].ObjToString();
                        lastName = dt.Rows[i]["lastName"].ObjToString();
                        address1 = dt.Rows[i]["address1"].ObjToString();
                        address2 = dt.Rows[i]["address2"].ObjToString();
                        city = dt.Rows[i]["city"].ObjToString();
                        state = dt.Rows[i]["state"].ObjToString();
                        zip = dt.Rows[i]["zip1"].ObjToString();

                        record = dt.Rows[i]["record"].ObjToString();
                        G1.update_db_table("icustomers", "record", record, new string[] { "firstName", firstName, "lastName", lastName, "address1", address1, "address2", address2, "city", city, "state", state, "zip1", zip });
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }

            btnSave.Hide();
            btnSave.Refresh();
        }
        /***********************************************************************************************/
        private void gridMain2_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            if (loading)
                return;
            DataRow dr = gridMain2.GetFocusedDataRow();
            DataTable dt = (DataTable)dgv2.DataSource;
            //if (e.Column.FieldName.ToUpper() == "DUEDATE")
            //{
            //    int rowhandle = gridMain.FocusedRowHandle;
            //    int row = gridMain.GetDataSourceRowIndex(rowhandle);
            //    string cnum = dt.Rows[row]["contractNumber"].ObjToString();
            //    if (!String.IsNullOrWhiteSpace(cnum))
            //    {
            //        string date = e.Value.ObjToString();
            //        if (G1.validate_date(date))
            //        {
            //            MySqlDateTime myDate = (MySqlDateTime)G1.DTtoMySQLDT(date);
            //            string dueDate = myDate.Month.ToString("D2") + "/" + myDate.Day.ToString("D2") + "/" + myDate.Year.ToString("D4");
            //            string record = dt.Rows[row]["record1"].ObjToString();
            //            G1.update_db_table(contractsFile, "record", record, new string[] { "dueDate8", dueDate });
            //            BandedGridView view = sender as BandedGridView;
            //            loading = true;
            //            view.SetRowCellValue(e.RowHandle, view.Columns["dueDate"], dueDate);
            //            loading = false;
            //        }
            //    }
            //}
            if (!chkFullEdit.Checked)
            {
                int rowhandle = gridMain2.FocusedRowHandle;
                int row = gridMain2.GetDataSourceRowIndex(rowhandle);
                string column = e.Column.FieldName;
                return;
            }
            string type = cmbType.Text;
            if (type.ToUpper() == "TRUSTS")
                modTrusts = true;
            else
                modInsurance = true;

            dr["mod"] = "Y";
            btnSave.Show();
            btnSave.Refresh();
        }
        /***********************************************************************************************/
    }
}
