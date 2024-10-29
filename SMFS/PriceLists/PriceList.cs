using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using GeneralLib;
using DevExpress.XtraPrinting;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using DevExpress.XtraReports.UI;
using System.Drawing.Printing;
using DevExpress.XtraPrinting.Native;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Export;
using DevExpress.XtraPrinting.Preview;
using DevExpress.XtraPrinting.Control;
using DevExpress.XtraCharts;
using DevExpress.XtraReports.Native;
using DevExpress.ReportServer.ServiceModel.DataContracts;
/***********************************************************************************************/
namespace SMFS
{
    /***********************************************************************************************/
    public partial class PriceList : DevExpress.XtraEditors.XtraForm
    {
        private bool allowSave = true;
        private bool isPackage = false;
        private string workPackage = "";
        private bool modified = false;
        private DataTable workDt = null;
        private string workDescription = "";
        private bool loading = false;
        private bool showSecondColumn = false;
        private string workFunRec = "";
        private string workPrices = "";
        private string workGroupName = "";
        private string workCasketName = "";
        private bool workMassPrint = false;
        private string workAsOfDate = "";
        private string funeralHome = "";
        private string funeralAtNeedCode = "";
        private double basicService = 0D;
        private double basicServices = 0D;
        /***********************************************************************************************/
        private string funeralHomeName = "";
        private string funeralAddress = "";
        private string funeralPOBox = "";
        private string funeralCity = "";
        private string funeralState = "";
        private string funeralZip = "";
        private string funeralPhone = "";
        PrintableComponentLink link;
        private DevExpress.XtraRichEdit.RichEditControl rtb9;
        /***********************************************************************************************/
        public PriceList(string description, string whatPrices = "", string workFuneralRecord = "", bool massPrint = false, string asOfDate = "" )
        {
            InitializeComponent();
            workDescription = description;
            workFunRec = workFuneralRecord;
            workPrices = whatPrices.ToUpper();
            workMassPrint = massPrint;
            workAsOfDate = asOfDate;
            if (workMassPrint)
            {
                this.Hide();
                loading = true;
                LoadData();
                loading = false;
                LoadRtbExtra();
                this.Text = workDescription;
                this.Close();
                //btnSave.Hide();
                //if (String.IsNullOrWhiteSpace(workPrices))
                //    btnFuture.Hide();
            }
        }
        /***********************************************************************************************/
        private void PriceList_Load(object sender, EventArgs e)
        {
            loading = true;
            LoadData();
            loading = false;
            LoadRtbExtra();
            this.Text = workDescription;
            btnSave.Hide();
            if (String.IsNullOrWhiteSpace(workPrices))
                btnFuture.Hide();
        }
        /***********************************************************************************************/
        private void LoadData()
        {
            if (!String.IsNullOrWhiteSpace(workAsOfDate))
            {
                DateTime date = workAsOfDate.ObjToDateTime();
                this.dateTimePicker1.Value = date;
                chkAsOfDate.Checked = true;
            }
            allowSave = true;
            isPackage = false;
            string cmd = "Select * from `pricelist` where `priceList` = '" + workDescription + "' order by `order`, `record`;";
            DataTable dt = G1.get_db_data(cmd);
            if (workDescription.ToUpper() == "GENERAL PRICE LIST")
            {
                Services.FixAllData(dt);
                FunServices.RunServiceTranslator(dt);
            }

            dt.Columns.Add("num");
            dt.Columns.Add("mod");
            dt.Columns.Add("found");

            allowSave = CheckForPackages(dt);
            if (String.IsNullOrWhiteSpace(workFunRec))
                allowSave = true;

            dt = SetupFuneralHome(dt);

            G1.NumberDataTable(dt);
            dgv.DataSource = dt;

            if (workMassPrint)
            {
                if ( workPrices.ToUpper() == "FUTURE")
                {
                    //workPrices = "CURRENT";
                    //btnFuture.Text = "Show Future";
                    //btnFuture.BackColor = Color.Tomato;

                    //loading = true;
                    //LoadData();
                    //loading = false;
                    //LoadRtbExtra();
                    //this.Text = workDescription;


                    //PriceList_Load(null, null);
                }
                btnPrint_Click(null, null);
                //this.Close();
            }
        }
        /***********************************************************************************************/
        private DataTable SetupFuneralHome(DataTable dt)
        {
            funeralHomeName = "Funeral Home Name";
            funeralAddress = "Funeral Address";
            funeralPOBox = "1234";
            funeralCity = "Funeral City";
            funeralState = "MS";
            funeralZip = "ZIP";
            funeralPhone = "Funeral Phone";
            if (String.IsNullOrWhiteSpace(workFunRec))
            {
                //CheckAndProcessPackages(dt);
                return dt;
            }
            string cmd = "Select * from `funeralHomes` where `record` = '" + workFunRec + "';";
            DataTable dx = G1.get_db_data(cmd);
            if (dx.Rows.Count <= 0)
                return dt;
            workGroupName = dx.Rows[0]["groupname"].ObjToString();
            workCasketName = dx.Rows[0]["casketgroup"].ObjToString();
            funeralHome = dx.Rows[0]["LocationCode"].ObjToString();
            funeralAtNeedCode = dx.Rows[0]["atNeedCode"].ObjToString();

            funeralHomeName = dx.Rows[0]["name"].ObjToString();
            funeralAddress = dx.Rows[0]["address"].ObjToString();
            funeralPOBox = dx.Rows[0]["POBox"].ObjToString();
            funeralCity = dx.Rows[0]["city"].ObjToString();
            funeralState = dx.Rows[0]["state"].ObjToString();
            funeralZip = dx.Rows[0]["zip"].ObjToString();
            funeralPhone = dx.Rows[0]["phoneNumber"].ObjToString();


            this.Text = workDescription + " for " + funeralHome + "-" + workGroupName + "-" + workCasketName;

            dt = SetupPrices(dt);
            return dt;
        }
        /***********************************************************************************************/
        private double GetBasicServices(DataTable gDt)
        {
            double basicServicesPrice = 0D;
            double price = 0D;
            string service = "";
            DataRow[] dR = null;
            DataTable dt = G1.get_db_data("Select * from `funeral_master` where `basicService` = '1';");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                service = dt.Rows[i]["service"].ObjToString();
                if (String.IsNullOrWhiteSpace(service))
                    continue;
                dR = gDt.Select("service='" + service + "'");
                if (dR.Length > 0)
                {
                    price = dR[0]["price"].ObjToDouble();
                    if (workPrices == "FUTURE")
                        price = dR[0]["futurePrice"].ObjToDouble();
                    basicServicesPrice += price;
                }

            }
            return basicServicesPrice;
        }
        /***********************************************************************************************/
        private DataTable SetupPrices(DataTable dt)
        {
            try
            {
                if (G1.get_column_number(dt, "price") < 0)
                    dt.Columns.Add("price", Type.GetType("System.Double"));
                if (G1.get_column_number(dt, "packageprice") < 0)
                    dt.Columns.Add("packageprice", Type.GetType("System.Double"));

                //string cmd = "Select * from `funeral_gplgroups` g JOIN `packages` p ON g.`record` = p.`!serviceRecord` where g.`groupname` = '" + workGroupName + "';";

                string cmd = "Select * from `funeral_gplgroups` where `groupname` = '" + workGroupName + "';";

                DataTable gDt = G1.get_db_data(cmd);

                PullLocationGPL(gDt); // Get Custom Location Prices for GPL Group


                Services.DoTheMath(gDt); // Add Services Together

                string basics = "BASIC SERVICES OF FUNERAL DIRECTOR AND STAFF";

                basicServices = GetBasicServices(gDt);

                cmd = "Select * from `funeral_master`;";
                DataTable ddx = G1.get_db_data(cmd);

                DataRow[] dRows = null;
                string service = "";

                for ( int i=0; i<gDt.Rows.Count; i++)
                {
                    if ( gDt.Rows[i]["SameAsMaster"].ObjToString() == "1")
                    {
                        service = gDt.Rows[i]["service"].ObjToString();
                        if (String.IsNullOrWhiteSpace(service))
                            continue;
                        dRows = ddx.Select("service='" + service + "'");
                        if ( dRows.Length > 0 )
                        {
                            gDt.Rows[i]["price"] = dRows[0]["price"].ObjToDouble();
                            gDt.Rows[i]["futurePrice"] = dRows[0]["futurePrice"].ObjToDouble();
                        }
                    }
                }

                string record = "";
                DataRow[] dR = null;
                if (isPackage)
                {
                    cmd = "Select * from `packages` where `groupname` = '" + workGroupName + "' and `PackageName` = '" + workPackage + "';";
                    DataTable newddx = G1.get_db_data(cmd);
                    for (int i = 0; i < newddx.Rows.Count; i++)
                    {
                        record = newddx.Rows[i]["!serviceRecord"].ObjToString();
                        dR = gDt.Select("!masterRecord='" + record + "'");
                        if (dR.Length > 0)
                        {
                            if (dR[0]["price"].ObjToDouble() <= 0D)
                                dR[0]["price"] = newddx.Rows[i]["price"].ObjToString();
                            if (dR[0]["futurePrice"].ObjToDouble() <= 0D)
                                dR[0]["futurePrice"] = newddx.Rows[i]["futurePrice"].ObjToString();
                        }
                    }
                }


                dt = CheckAndProcessPackages(dt);

                if (isPackage)
                {
                    if (G1.get_column_number(dt, "DELETED") < 0)
                        dt.Columns.Add("DELETED");
                    if (G1.get_column_number(dt, "ModMod") < 0)
                        dt.Columns.Add("ModMod");
                    PriceList.replaceCredits(dt, "PACKAGE");
                }

                //gDt = dt.Copy();

                DataTable tempDt = null;

                service = "";
                string priceList = "";
                dRows = null;
                double price = 0D;
                double futurePrice = 0D;
                for (int i = 0; i < gDt.Rows.Count; i++)
                {
                    service = gDt.Rows[i]["service"].ObjToString();
                    if (String.IsNullOrWhiteSpace(service))
                        continue;
                    if (service.ToUpper() == basics)
                    {
                    }
                    if (service.ToUpper() == "TOTAL LISTED PRICE")
                        continue;
                    if (service.ToUpper() == "PACKAGE DISCOUNT")
                        continue;
                    if (service.ToUpper() == "PACKAGE PRICE")
                        continue;
                    price = gDt.Rows[i]["price"].ObjToDouble();
                    if (workPrices == "FUTURE")
                        price = gDt.Rows[i]["futurePrice"].ObjToDouble();
                    if (!isPackage)
                    {
                        if (price <= 0D)
                            price = gDt.Rows[i]["price"].ObjToDouble();
                    }
                    if (service.Trim().ToUpper() == basics)
                        basicService = price;

                    dRows = dt.Select("service='" + service + "'");
                    if (dRows.Length > 0)
                    {
                        tempDt = dRows.CopyToDataTable();
                        for (int j = 0; j < dRows.Length; j++)
                        {
                            dRows[j]["found"] = "Y";
                            priceList = dRows[j]["priceList"].ObjToString();
                            if ( !String.IsNullOrWhiteSpace ( priceList))
                                dRows[j]["price"] = price;
                        }
                    }
                }

                ProcessMerchansice(dt);

                showSecondColumn = false;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    service = dt.Rows[i]["service"].ObjToString();
                    if (service.Trim().ToUpper() == "{INCLUDE BOTH PRICES}")
                    {
                        showSecondColumn = true;
                        break;
                    }
                }
                if (!showSecondColumn)
                    gridMain.Columns["packageprice"].Visible = false;

            }
            catch (Exception ex)
            {

            }
            return dt;
        }
        /***********************************************************************************************/
        private bool CheckForPackages ( DataTable gDt)
        {
            bool found = true;
            string service = "";
            for (int i = 0; i < gDt.Rows.Count; i++)
            {
                service = gDt.Rows[i]["service"].ObjToString();
                if (service.ToUpper().IndexOf("PACKAGE=") >= 0)
                {
                    isPackage = true;
                    found = false;
                    int idx = service.ToUpper().IndexOf("PACKAGE=");
                    workPackage = service.Substring(idx + 8);
                    break;
                }
            }
            return found;
        }
        /***********************************************************************************************/
        private DataTable CheckAndProcessPackages ( DataTable gDt )
        {
            string service = "";
            string package = "";
            int idx = 0;
            DataTable nDt = gDt.Clone();
            DataTable xDt = null;
            DataRow dRow = null;
            string mainPriceListTitle = "";
            if ( gDt.Rows.Count > 0 )
                mainPriceListTitle = gDt.Rows[0]["pricelist"].ObjToString();
            for (int i = 0; i < gDt.Rows.Count; i++)
            {
                try
                {
                    service = gDt.Rows[i]["service"].ObjToString();
                    if (service.ToUpper().IndexOf("PACKAGE=") >= 0)
                    {
                        idx = service.ToUpper().IndexOf("PACKAGE=");
                        package = service.Substring(idx + 8);
                        if (!String.IsNullOrWhiteSpace(package))
                        {
                            xDt = LoadPackage(gDt, i, workGroupName, package);
                            for (int j = 0; j < xDt.Rows.Count; j++)
                            {
                                if (workPrices == "FUTURE")
                                    xDt.Rows[j]["price"] = xDt.Rows[j]["futurePrice"].ObjToDouble();
                                xDt.Rows[j]["pricelist"] = package;
                                nDt.ImportRow(xDt.Rows[j]);
                                dRow = nDt.NewRow();
                                dRow["service"] = ".";
                                nDt.Rows.Add(dRow);
                            }
                        }
                    }
                    else
                    {
                        nDt.ImportRow(gDt.Rows[i]);
                    }
                }
                catch ( Exception ex)
                {
                }
            }

            return nDt;
        }
        /***********************************************************************************************/
        private DataTable LoadPackage(DataTable gDt, int mainRow, string group, string package)
        {
            string serviceRecord = "";
            if (String.IsNullOrWhiteSpace(package))
                package = "Master";
            if (String.IsNullOrWhiteSpace(group))
                group = "Group 3 GPL"; // Just for testing
            if (String.IsNullOrWhiteSpace(group) || String.IsNullOrWhiteSpace(package))
            {
                MessageBox.Show("***ERROR*** Empty group or package!");
                return gDt;
            }

            string list = "";

            string cmd = "Select * from `packages` where `groupname` = 'master' and `PackageName` = '" + package + "';";
            DataTable dx = G1.get_db_data(cmd);
            for (int i = 0; i < dx.Rows.Count; i++)
            {
                serviceRecord = dx.Rows[i]["!serviceRecord"].ObjToString();
                if (String.IsNullOrWhiteSpace(serviceRecord))
                    continue;
                list += "'" + serviceRecord + "',";
            }

            //string cmd = "Select * from `packages` where `groupname` = '" + group + "' and `PackageName` = '" + package + "';";
            //DataTable dx = G1.get_db_data(cmd);
            //for (int i = 0; i < dx.Rows.Count; i++)
            //{
            //    serviceRecord = dx.Rows[i]["!serviceRecord"].ObjToString();
            //    if (String.IsNullOrWhiteSpace(serviceRecord))
            //        continue;
            //    list += "'" + serviceRecord + "',";
            //}
            list = list.TrimEnd(',');
            if (!String.IsNullOrWhiteSpace(list))
            {
                cmd = "Select * from `packages` p LEFT JOIN `funeral_master` s ON p.`!serviceRecord` = s.`record` where s.`record` IN (" + list + ") ";
                //                cmd = "Select * from `packages` p RIGHT JOIN `services` s ON p.`!serviceRecord` = s.`record` where s.`record` IN (" + list + ") ";
                cmd += " and `groupname` = 'master' and `PackageName` = '" + package + "' ";
                cmd += ";";
            }
            else
            {
                cmd = "Select * from `packages` p JOIN `funeral_master` s ON p.`!serviceRecord` = s.`record` where `service` = 'xyzzyxxxx';";
            }
            DataTable dt = G1.get_db_data(cmd);
            dt.Columns.Add("num");
            dt.Columns.Add("mod");
            dt.Columns.Add("select");
            dt.Columns.Add("total", Type.GetType("System.Double"));
            string service = "";
            string record = "";
            DataRow[] dR = null;
            DataTable ddx = null;

            cmd = "Select * from `funeral_gplgroups` where `groupname` = '" + group + "';";
            ddx = G1.get_db_data(cmd);
            cmd = "Select * from `packages` where `groupname` = '" + group + "' and `PackageName` = '" + package + "';";
            DataTable newddx = G1.get_db_data(cmd);
            for (int i = 0; i < newddx.Rows.Count; i++)
            {
                record = newddx.Rows[i]["!serviceRecord"].ObjToString();
                dR = ddx.Select("!masterRecord='" + record + "'");
                if (dR.Length > 0)
                {
                    if (dR[0]["price"].ObjToDouble() <= 0D)
                        dR[0]["price"] = newddx.Rows[i]["price"].ObjToString();
                    if (dR[0]["futurePrice"].ObjToDouble() <= 0D)
                        dR[0]["futurePrice"] = newddx.Rows[i]["futurePrice"].ObjToString();
                }
            }
            SetupSameAsMaster(dt, ddx);

            G1.NumberDataTable(dt);

            dR = dt.Select("service='Package Price'");
            if (dR.Length > 0)
            {
                double packagePrice = dR[0]["price"].ObjToDouble();
                double packageFuture = dR[0]["futurePrice"].ObjToDouble();
                double totalPrice = 0D;
                double totalFuture = 0D;
                service = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    service = dt.Rows[i]["service"].ObjToString().ToUpper();
                    if (service == "TOTAL LISTED PRICE")
                        continue;
                    if (service == "PACKAGE DISCOUNT")
                        continue;
                    if (service == "PACKAGE PRICE")
                        continue;
                    totalPrice += dt.Rows[i]["price"].ObjToDouble();
                    totalFuture += dt.Rows[i]["futurePrice"].ObjToDouble();
                }
                double totalListedPrice = totalPrice;
                double packageDiscount = totalPrice - packagePrice;
                if (packageDiscount > 0D)
                    packageDiscount = packageDiscount * (-1D);

                double totalListedFuture = totalFuture;
                double packageFutureDiscount = totalFuture - packageFuture;
                if (packageFutureDiscount > 0D)
                    packageFutureDiscount = packageFutureDiscount * (-1D);

                dR = dt.Select("service='Total Listed Price'");
                if (dR.Length > 0)
                {
                    dR[0]["price"] = totalListedPrice.ToString();
                    dR[0]["futurePrice"] = totalListedFuture;
                }

                dR = dt.Select("service='Package Discount'");
                if (dR.Length > 0)
                {
                    dR[0]["price"] = packageDiscount.ToString();
                    dR[0]["futurePrice"] = packageFutureDiscount;
                }
            }
            double size = gDt.Rows[mainRow]["size"].ObjToDouble();
            DataRow dRow = null;
            DataTable xDt = gDt.Clone();
            int newRow = 0;
            for ( int i=0; i<dt.Rows.Count; i++)
            {
                try
                {
                    dRow = xDt.NewRow();
                    xDt.Rows.Add(dRow);
                    service = dt.Rows[i]["service"].ObjToString();
                    if ( service.ToUpper() == "TOTAL LISTED PRICE")
                    {
                        xDt.Rows[newRow]["size"] = 7D;
                        xDt.Rows[newRow]["underline"] = "y";
                        xDt.Rows[newRow]["service"] = ".";
                        newRow++;
                        dRow = xDt.NewRow();
                        xDt.Rows.Add(dRow);
                    }
                    G1.HardCopyDtRow(dt, i, xDt, newRow);
                    xDt.Rows[newRow]["size"] = size;
                    service = xDt.Rows[newRow]["service"].ObjToString();
                    //if ( service.ToUpper() == "URN CREDIT")
                    //    xDt.Rows[newRow]["underline"] = "n";

                    if ( service.ToUpper() == "TOTAL LISTED PRICE")
                        xDt.Rows[newRow]["bold"] = "y";

                    if (service.ToUpper() == "PACKAGE DISCOUNT")
                    {
                        xDt.Rows[newRow]["bold"] = "y";
                        xDt.Rows[newRow]["underline"] = "y";
                    }

                    if (service.ToUpper() == "PACKAGE PRICE")
                        xDt.Rows[newRow]["bold"] = "y";
                    newRow++;
                }
                catch ( Exception ex)
                {
                }
            }
            //gDt.Rows.RemoveAt(mainRow + dt.Rows.Count);
            xDt.AcceptChanges();
            return xDt;
        }
        /***********************************************************************************************/
        private void SetupSameAsMaster(DataTable dt, DataTable ddx = null)
        {
            bool GPLddx = false;
            if (dt != null)
            {
                if (ddx == null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                        dt.Rows[i]["SameAsMaster"] = "0";
                }
                else
                {
                    if (G1.get_column_number(ddx, "groupname") >= 0)
                        GPLddx = true;
                    DataRow[] dR = null;
                    string select = "";
                    string service = "";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (!GPLddx)
                        {
                            select = dt.Rows[i]["SameAsMaster"].ObjToString();
                            if (select == "1")
                            {
                                dt.Rows[i]["SameAsMaster"] = "1";
                                service = dt.Rows[i]["service"].ObjToString();
                                dR = ddx.Select("service='" + service + "'");
                                if (dR.Length > 0)
                                {
                                    dt.Rows[i]["price"] = dR[0]["price"];
                                }
                            }
                            else
                                dt.Rows[i]["SameAsMaster"] = "0";
                        }
                        else
                        {
                            service = dt.Rows[i]["service"].ObjToString();
                            dR = ddx.Select("service='" + service + "'");
                            if (dR.Length > 0)
                            {
                                select = dR[0]["SameAsMaster"].ObjToString();
                                if (select == "1")
                                {
                                    dt.Rows[i]["SameAsMaster"] = "1";
                                    dt.Rows[i]["price"] = dR[0]["price"];
                                    dt.Rows[i]["futurePrice"] = dR[0]["futurePrice"];
                                    dt.Rows[i]["pastPrice"] = dR[0]["pastPrice"];
                                }
                                else
                                {
                                    dt.Rows[i]["SameAsMaster"] = "0";
                                    dt.Rows[i]["price"] = dR[0]["price"];
                                    dt.Rows[i]["futurePrice"] = dR[0]["futurePrice"];
                                    dt.Rows[i]["pastPrice"] = dR[0]["pastPrice"];
                                }
                            }
                            else
                                dt.Rows[i]["SameAsMaster"] = "0";
                        }
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void PullLocationGPL(DataTable dx)
        {
            string gplGroup = workGroupName;
            if (String.IsNullOrWhiteSpace(gplGroup))
                return;
            if (gplGroup.Trim().ToUpper() == "MASTER")
                return;

            string location = "(" + funeralAtNeedCode + ") " + funeralHome;
            if (String.IsNullOrWhiteSpace(location))
                return;

            string cmd = "Select * from `gpl_locations` where `gpl` = '" + gplGroup + "' and `location` = '" + location + "';";
            DataTable dt = G1.get_db_data(cmd);
            if (dt.Rows.Count <= 0)
                return;
            if (dx != null)
            {
                DataRow[] dRows = null;
                string service = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    service = dt.Rows[i]["service"].ObjToString();
                    dRows = dx.Select("service='" + service + "'");
                    if (dRows.Length > 0)
                    {
                        dRows[0]["price"] = dt.Rows[i]["price"].ObjToString();
                        dRows[0]["futurePrice"] = dt.Rows[i]["futurePrice"].ObjToString();
                        dRows[0]["pastPrice"] = dt.Rows[i]["pastPrice"].ObjToString();
                        //dRows[0]["data"] = "CUSTOM";
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void ProcessMerchansice(DataTable dt)
        {
            string cmd = "Select * from `casket_packages` p RIGHT JOIN `casket_master` s ON p.`!masterRecord` = s.`record` where ";
            cmd += " `groupname` = '" + workCasketName + "' ";
            cmd += ";";

            double markup = 0D;
            double cost = 0D;
            double price = 0D;
            double packageCost = 0D;
            double rounding = 0D;
            string service = "";
            DataRow[] dRows = null;

            DataTable ddb = null;

            DataTable gDt = G1.get_db_data(cmd);
            PullLocationCaskets(gDt);

            cmd = "Select * from `casket_master`;";
            DataTable cDt = G1.get_db_data(cmd);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    service = dt.Rows[i]["service"].ObjToString();
                    if (service.Trim() == "Wilbert Bronze - Size 30")
                    {
                    }
                    dRows = gDt.Select("casketdesc='" + service + "'");
                    if (dRows.Length <= 0)
                    {
                        dRows = cDt.Select("casketdesc='" + service + "'");
                        if (dRows.Length > 0)
                        {
                            rounding = dRows[0]["round"].ObjToDouble();
                            //markup = dRows[0]["markup"].ObjToDouble();
                            markup = 1D;
                            if (workPrices == "FUTURE")
                            {
                                price = dRows[0]["futureCasketCost"].ObjToDouble();
                                if (price <= 0D)
                                {
                                    price = dRows[0]["casketcost"].ObjToDouble();
                                    if (rounding > 0D)
                                        price = Caskets.RoundTo(price, rounding);
                                }
                                else
                                {
                                    if (rounding > 0D)
                                        price = Caskets.RoundTo(price, rounding);
                                    price = price * markup;

                                }
                            }
                            else
                            {
                                price = dRows[0]["casketcost"].ObjToDouble();
                                if (rounding > 0D)
                                    price = Caskets.RoundTo(price, rounding);
                                price = price * markup;
                            }
                            dt.Rows[i]["price"] = price;
                            packageCost = price + basicServices;
                            dt.Rows[i]["packageprice"] = packageCost;
                            dt.Rows[i]["found"] = "Y";
                        }
                    }
                    else
                    {
                        ddb = dRows.CopyToDataTable();
                        dRows[0]["casketCost"] = dRows[0]["price"].ObjToDouble();
                        rounding = dRows[0]["round"].ObjToDouble();
                        markup = dRows[0]["markup"].ObjToDouble();
                        price = dRows[0]["casketCost"].ObjToDouble();
                        cost = dRows[0]["casketcost"].ObjToDouble();
                        if (workPrices == "FUTURE")
                        {
                            cost = dRows[0]["futureCasketCost"].ObjToDouble();
                            if (cost <= 0D)
                                cost = dRows[0]["casketcost"].ObjToDouble();
                            markup = dRows[0]["futuremarkup"].ObjToDouble();
                        }
                        price = cost;
                        price = price * markup;
                        if (rounding > 0D)
                            price = Caskets.RoundTo(price, rounding);
                        dt.Rows[i]["price"] = price;
                        packageCost = price + basicServices;
                        dt.Rows[i]["packageprice"] = packageCost;
                        dt.Rows[i]["found"] = "Y";
                    }
                }
                catch (Exception ex)
                {
                }
            }
        }
        /***********************************************************************************************/
        private void PullLocationCaskets(DataTable dx)
        {
            string casketGroup = workCasketName;
            if (String.IsNullOrWhiteSpace(casketGroup))
                return;
            if (casketGroup.Trim().ToUpper() == "MASTER")
                return;

            string location = "(" + funeralAtNeedCode + ") " + funeralHome;
            if (String.IsNullOrWhiteSpace(location))
                return;

            string cmd = "Select * from `casket_locations` where `casketGroup` = '" + casketGroup + "' and `location` = '" + location + "';";
            DataTable dt = G1.get_db_data(cmd);
            if (dt.Rows.Count <= 0)
                return;
            if (dx != null)
            {
                DataRow[] dRows = null;
                string casketcode = "";
                string casketdesc = "";
                bool gotData = true;
                if (G1.get_column_number(dx, "type") < 0)
                    dx.Columns.Add("type");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    try
                    {
                        casketcode = dt.Rows[i]["casketcode"].ObjToString();
                        casketdesc = dt.Rows[i]["casketdesc"].ObjToString();
                        if (!String.IsNullOrWhiteSpace(casketcode))
                            dRows = dx.Select("casketcode='" + casketcode + "'");
                        else
                            dRows = dx.Select("casketdesc='" + casketdesc + "'");
                        if (dRows.Length > 0)
                        {
                            dRows[0]["price"] = dt.Rows[i]["price"].ObjToString();
                            dRows[0]["futureCasketCost"] = dt.Rows[i]["futureCasketCost"].ObjToString();
                            dRows[0]["pastCasketCost"] = dt.Rows[i]["pastCasketCost"].ObjToString();
                            //dRows[0]["casket"] = dt.Rows[i]["casket"].ObjToString();
                            dRows[0]["casketcost"] = dt.Rows[i]["casketcost"].ObjToString();
                            dRows[0]["casketcost"] = dt.Rows[i]["casket"].ObjToString();
                            dRows[0]["price"] = dt.Rows[i]["casket"].ObjToString();
                            dRows[0]["markup"] = dt.Rows[i]["markup"].ObjToString();
                            dRows[0]["futuremarkup"] = dt.Rows[i]["futuremarkup"].ObjToString();
                            if (gotData)
                                dRows[0]["type"] = "CUSTOM";
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
        }
        /***********************************************************************************************/
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            G1.ShowHideFindPanel(gridMain);
            //if (gridMain.OptionsFind.AlwaysVisible == true)
            //    gridMain.OptionsFind.AlwaysVisible = false;
            //else
            //    gridMain.OptionsFind.AlwaysVisible = true;
        }
        /***********************************************************************************************/
        private void pictureAdd_Click(object sender, EventArgs e)
        {
            string group = "";
            string package = "";
            DataTable dt = (DataTable)dgv.DataSource;
            DataTable dx = dt.Clone();
            Services serviceForm = new Services(group, package, true, dx);
            serviceForm.SelectDone += ServiceForm_SelectDone;
            serviceForm.Show();
        }
        /***************************************************************************************/
        private void ServiceForm_SelectDone(DataTable dt, string what )
        {
            workDt = (DataTable)dgv.DataSource;

            string select = "";
            string service = "";
            double price = 0D;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                select = dt.Rows[i]["select"].ObjToString();
                if (select == "1")
                {
                    service = dt.Rows[i]["service"].ObjToString();
                    DataRow newRow = workDt.NewRow();
                    newRow["service"] = service;
                    newRow["size"] = 0D;
                    newRow["layin"] = 0;
                    newRow["indent"] = 0;
                    newRow["price"] = price;
                    workDt.Rows.Add(newRow);
                }
            }

            if (!String.IsNullOrWhiteSpace(workFunRec))
                workDt = SetupPrices(workDt);

            G1.NumberDataTable(workDt);
            dgv.DataSource = workDt;
            dgv.Refresh();
            modified = true;
            if ( allowSave )
                btnSave.Show();
        }
        /***********************************************************************************************/
        //private void pictureAdd_Click(object sender, EventArgs e)
        //{
        //    DataTable dt = (DataTable)dgv.DataSource;
        //    int lines = 1;
        //    for (int i = 0; i < lines; i++)
        //    {
        //        DataRow dRow = dt.NewRow();
        //        dRow["num"] = dt.Rows.Count.ObjToInt32() + 1;
        //        dt.Rows.Add(dRow);
        //    }
        //    dgv.DataSource = dt;

        //    int row = dt.Rows.Count - 1;
        //    gridMain.SelectRow(row);
        //    gridMain.FocusedRowHandle = row;
        //    dgv.RefreshDataSource();
        //    dgv.Refresh();
        //}
        /***********************************************************************************************/
        private void pictureDelete_Click(object sender, EventArgs e)
        {
            DataRow dr = gridMain.GetFocusedDataRow();
            string record = dr["record"].ObjToString();
            string service = dr["service"].ObjToString();
            DialogResult result = MessageBox.Show("***Question*** Are you sure you want to delete this price (" + service + ") ?", "Delete Price Item Dialog", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
                return;

            DataTable dt = (DataTable)dgv.DataSource;
            int[] rows = gridMain.GetSelectedRows();
            int dtRow = 0;
            int firstRow = 0;
            if (rows.Length > 0)
                firstRow = rows[0];
            int row = 0;
            try
            {
                loading = true;
                for (int i = 0; i < rows.Length; i++)
                {
                    row = rows[i];
                    dtRow = gridMain.GetDataSourceRowIndex(row);
                    if (dtRow < 0 || dtRow > (dt.Rows.Count - 1))
                    {
                        continue;
                    }
                    var dRow = gridMain.GetDataRow(row);
                    if (dRow != null)
                        dRow["mod"] = "D";
                    dt.Rows[dtRow]["mod"] = "D";
                    modified = true;
                    if ( allowSave )
                        btnSave.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("*ERROR*** " + ex.Message.ToString());
            }

            //            gridMain.FocusedRowHandle = firstRow;
            loading = false;
            if (firstRow > (dt.Rows.Count - 1))
                firstRow = (dt.Rows.Count - 1);
            dgv.DataSource = dt;
            gridMain.RefreshData();
            dgv.Refresh();

            gridMain.FocusedRowHandle = firstRow;
            gridMain.SelectRow(firstRow);
        }
        /***********************************************************************************************/
        private void picRowUp_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
                return;
            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            if (rowHandle == 0)
                return; // Already at the first row
            //MoveRowUp(dt, rowHandle);
            massRowsUp(dt, rowHandle);
            dt.AcceptChanges();
            dgv.DataSource = dt;
            gridMain.ClearSelection();
            gridMain.SelectRow(rowHandle - 1);
            gridMain.FocusedRowHandle = rowHandle - 1;
            gridMain.RefreshData();
            dgv.Refresh();
        }
        /***********************************************************************************************/
        private void massRowsUp(DataTable dt, int row)
        {
            int[] rows = gridMain.GetSelectedRows();
            int firstRow = 0;
            if (rows.Length > 0)
                firstRow = rows[0];
            try
            {
                G1.NumberDataTable(dt);
                dt.Columns.Add("Count", Type.GetType("System.Int32"));
                for (int i = 0; i < dt.Rows.Count; i++)
                    dt.Rows[i]["Count"] = i.ToString();
                int moverow = rows[0];
                for (int i = 0; i < rows.Length; i++)
                {
                    row = rows[i];
                    dt.Rows[row]["Count"] = (row - 1).ToString();
                    //dt.Rows[row - 1]["Count"] = row.ToString();
                    //var dRow = gridMain.GetDataRow(row);
                    dt.Rows[row]["mod"] = "M";
                    modified = true;
                    if ( allowSave )
                        btnSave.Show();
                }
                dt.Rows[moverow - 1]["Count"] = (moverow + (rows.Length - 1)).ToString();
                G1.sortTable(dt, "Count", "asc");
                dt.Columns.Remove("Count");
                G1.NumberDataTable(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show("*ERROR*** " + ex.Message.ToString());
            }

            //            gridMain.FocusedRowHandle = firstRow;
            gridMain.SelectRow(firstRow);
            dgv.DataSource = dt;
        }
        /***********************************************************************************************/
        private void picRowDown_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
                return;
            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            if (rowHandle == (dt.Rows.Count - 1))
                return; // Already at the last row
            MoveRowDown(dt, rowHandle);
            dt.AcceptChanges();
            dgv.DataSource = dt;
            gridMain.ClearSelection();
            gridMain.SelectRow(rowHandle + 1);
            gridMain.FocusedRowHandle = rowHandle + 1;
            gridMain.RefreshData();
            dgv.Refresh();
        }
        /***************************************************************************************/
        private void MoveRowDown(DataTable dt, int row)
        {
            dt.Columns.Add("Count", Type.GetType("System.Int32"));
            for (int i = 0; i < dt.Rows.Count; i++)
                dt.Rows[i]["Count"] = i.ToString();
            dt.Rows[row]["Count"] = (row + 1).ToString();
            dt.Rows[row + 1]["Count"] = row.ToString();
            G1.sortTable(dt, "Count", "asc");
            dt.Columns.Remove("Count");
            for (int i = 0; i < dt.Rows.Count; i++)
                dt.Rows[i]["Num"] = (i + 1).ToString();
        }
        /***********************************************************************************************/
        private void btnInsert_Click(object sender, EventArgs e)
        {
            DataRow dRow = null;
            DataTable dt = (DataTable)dgv.DataSource;
            if (dt.Rows.Count <= 0)
            {
                dRow = dt.NewRow();
                dt.Rows.Add(dRow);
                G1.NumberDataTable(dt);
                dt.AcceptChanges();
                dgv.DataSource = dt;
                gridMain.ClearSelection();
                gridMain.RefreshData();
                gridMain.FocusedRowHandle = 0;
                gridMain.SelectRow(0);
                dgv.Refresh();
                modified = true;
                if ( allowSave )
                    btnSave.Show();
                return;
            }
            DataRow dr = gridMain.GetFocusedDataRow();
            int rowHandle = gridMain.FocusedRowHandle;
            int dtRow = gridMain.GetDataSourceRowIndex(rowHandle);
            if (dtRow < 0 || dtRow > (dt.Rows.Count - 1))
                return;
            //if (rowHandle == (dt.Rows.Count - 1))
            //    return; // Already at the last row
            dRow = dt.NewRow();
            dt.Rows.InsertAt(dRow, dtRow);
            G1.NumberDataTable(dt);
            dt.AcceptChanges();
            dgv.DataSource = dt;
            gridMain.ClearSelection();
            gridMain.RefreshData();
            gridMain.FocusedRowHandle = rowHandle + 1;
            gridMain.SelectRow(rowHandle + 1);
            dgv.Refresh();
            modified = true;
            if ( allowSave )
                btnSave.Show();
        }
        /***********************************************************************************************/
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
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

            Printer.setupPrinterMargins(50, 50, 80, 50);

            pageMarginLeft = Printer.pageMarginLeft;
            pageMarginRight = Printer.pageMarginRight;
            pageMarginTop = Printer.pageMarginTop;
            pageMarginBottom = Printer.pageMarginBottom;

            printableComponentLink1.Margins.Left = pageMarginLeft;
            printableComponentLink1.Margins.Right = pageMarginRight;
            printableComponentLink1.Margins.Top = pageMarginTop;
            printableComponentLink1.Margins.Bottom = pageMarginBottom;

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
            printableComponentLink1.PrintingSystemBase = printingSystem1;
            printableComponentLink1.CreateDetailHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateDetailHeaderArea);
            printableComponentLink1.CreateMarginalHeaderArea += new DevExpress.XtraPrinting.CreateAreaEventHandler(this.printableComponentLink1_CreateMarginalHeaderArea);
            printableComponentLink1.BeforeCreateAreas += new System.EventHandler(this.printableComponentLink1_BeforeCreateAreas);
            printableComponentLink1.AfterCreateAreas += new System.EventHandler(this.printableComponentLink1_AfterCreateAreas);
            printableComponentLink1.Landscape = false;

            Printer.setupPrinterMargins(50, 50, 80, 50);

            pageMarginLeft = Printer.pageMarginLeft;
            pageMarginRight = Printer.pageMarginRight;
            pageMarginTop = Printer.pageMarginTop;
            pageMarginBottom = Printer.pageMarginBottom;

            printableComponentLink1.Margins.Left = pageMarginLeft;
            printableComponentLink1.Margins.Right = pageMarginRight;
            printableComponentLink1.Margins.Top = pageMarginTop;
            printableComponentLink1.Margins.Bottom = pageMarginBottom;

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

            font = new Font("Ariel", 10, FontStyle.Bold);
            Printer.DrawQuad(6, 8, 4, 4, "Price List", Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);


            //            Printer.DrawQuadTicks();
            DateTime date = DateTime.Now;
            string workDate = date.Month.ToString("D2") + "/" + date.Year.ToString("D4");
            Printer.SetQuadSize(24, 12);
            font = new Font("Ariel", 9, FontStyle.Bold);
            //            Printer.DrawQuad(20, 8, 5, 4, "Report Month:" + workDate, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(16, 8, 3, 4, lblPayment.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(19, 8, 3, 4, lblTrust85.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);
            //Printer.DrawQuad(22, 8, 3, 4, lblTrust100.Text, Color.Black, BorderSide.None, font, HorizontalAlignment.Left, VertAlignment.Bottom);

            Printer.SetQuadSize(12, 12);
            Printer.DrawQuadBorder(1, 1, 12, 12, BorderSide.All, 1, Color.Black);
            Printer.DrawQuadBorder(12, 1, 1, 12, BorderSide.Right, 1, Color.Black);
        }
        /***********************************************************************************************/
        private double oldValue = 0D;
        /***********************************************************************************************/
        private void gridMain_CellValueChanging(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (loading)
                return;
            string columnName = e.Column.FieldName.ToUpper();
            if (columnName != "PRICE")
            {
                oldValue = 0D;
                return;
            }
            DataRow dr = gridMain.GetFocusedDataRow();
            DataTable dt = (DataTable)dgv.DataSource;
            int row = e.RowHandle;
            oldValue = dt.Rows[row]["price"].ObjToDouble();
        }
        /***********************************************************************************************/
        private void gridMain_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (loading)
                return;
            modified = true;
            DataRow dr = gridMain.GetFocusedDataRow();
            DataTable dt = (DataTable)dgv.DataSource;
            if (e.Column.FieldName.ToUpper() == "PRICE")
            {
                int row = gridMain.FocusedRowHandle;
                string str = e.Value.ObjToString();
                if (String.IsNullOrWhiteSpace(str))
                {
                    dr["price"] = oldValue;
                    return;
                }
                else
                    dr["mod"] = "Y";
            }
            if ( allowSave)
                btnSave.Show();
        }
        /***********************************************************************************************/
        private void gridMain_CustomRowFilter(object sender, DevExpress.XtraGrid.Views.Base.RowFilterEventArgs e)
        {
            int row = e.ListSourceRow;
            DataTable dt = (DataTable)dgv.DataSource;
            if (G1.get_column_number(dt, "mod") < 0)
                return;
            string delete = dt.Rows[row]["mod"].ObjToString();
            if (delete.ToUpper() == "D")
            {
                e.Visible = false;
                e.Handled = true;
            }
        }
        /***********************************************************************************************/
        private void PriceLists_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!modified)
                return;
            if (!allowSave)
                return;
            DialogResult result = MessageBox.Show("***Question***\nPrice Lists have been modified!\nWould you like to save your changes?", "Price Lists Dialog", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }
            modified = false;
            if (result == DialogResult.No)
                return;
            SavePriceLists();
        }
        /***********************************************************************************************/
        private void SavePriceLists()
        {
            string record = "";
            string service = "";
            string mod = "";
            byte[] bytes = null;
            string underline = "";
            string bold = "";
            double indent = 0D;
            double size = 0D;
            double layin = 0D;
            double price = 0D;

            DataTable dt = (DataTable)dgv.DataSource;
            this.Cursor = Cursors.WaitCursor;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                record = dt.Rows[i]["record"].ObjToString();
                mod = dt.Rows[i]["mod"].ObjToString();
                if (mod.ToUpper() == "D")
                {
                    if (!String.IsNullOrWhiteSpace(record))
                        G1.delete_db_table("pricelist", "record", record);
                    continue;
                }
                if (String.IsNullOrWhiteSpace(record))
                    record = G1.create_record("pricelist", "service", "-1");
                if (G1.BadRecord("pricelist", record))
                    continue;
                dt.Rows[i]["record"] = record;
                service = dt.Rows[i]["service"].ObjToString();
                underline = dt.Rows[i]["underline"].ObjToString();
                size = dt.Rows[i]["size"].ObjToDouble();
                indent = dt.Rows[i]["indent"].ObjToDouble();
                bold = dt.Rows[i]["bold"].ObjToString();
                layin = dt.Rows[i]["layin"].ObjToDouble();
                price = dt.Rows[i]["price"].ObjToDouble();

                G1.update_db_table("pricelist", "record", record, new string[] { "pricelist", workDescription, "service", service, "size", size.ToString(), "indent", indent.ToString(), "underline", underline, "bold", bold, "layin", layin.ToString(), "order", i.ToString() });
                if (mod == "Y")
                    G1.update_db_table("pricelist", "record", record, new string[] { "price", price.ToString() });

                bytes = dt.Rows[i]["header"].ObjToBytes();
                G1.update_blob("priceList", "record", record, "header", bytes);

                bytes = dt.Rows[i]["tail"].ObjToBytes();
                G1.update_blob("priceList", "record", record, "tail", bytes);
            }

            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                mod = dt.Rows[i]["mod"].ObjToString();
                if (mod.ToUpper() == "D")
                    dt.Rows.RemoveAt(i);
                else
                    dt.Rows[i]["mod"] = "";
            }
            dgv.DataSource = dt;
            this.Cursor = Cursors.Default;
        }
        /***********************************************************************************************/
        private void gridMain_DoubleClick(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DataRow dr = gridMain.GetFocusedDataRow();

            int rowHandle = gridMain.FocusedRowHandle;
            int row = gridMain.GetDataSourceRowIndex(rowHandle);

            DataTable dt = (DataTable)dgv.DataSource;
            string record = dr["record"].ObjToString();

            RichTextBox rtb1 = new RichTextBox();
            RichTextBox rtb2 = new RichTextBox();
            byte[] bytes = null;
            MemoryStream stream = null;
            DevExpress.XtraRichEdit.RichEditControl rtb = new RichEditControl();
            string str = "";

            //            str = G1.get_db_blob("priceList", record, "header");
            bytes = dt.Rows[row]["header"].ObjToBytes();
            if (bytes != null)
            {
                str = G1.ConvertToString(bytes);
                //str = G1.DecompressString(str);

                if (!String.IsNullOrWhiteSpace(str))
                {
                    stream = new MemoryStream(bytes);
                    rtb.Document.Delete(rtb.Document.Range);
                    rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);
                    rtb1.Rtf = rtb.Document.RtfText;
                }
            }
            bytes = dt.Rows[row]["tail"].ObjToBytes();
            if (bytes != null)
            {
                str = G1.ConvertToString(bytes);
                //str = G1.DecompressString(str);

                if (!String.IsNullOrWhiteSpace(str))
                {
                    stream = new MemoryStream(bytes);
                    rtb.Document.Delete(rtb.Document.Range);
                    rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);
                    rtb2.Rtf = rtb.Document.RtfText;
                }
            }
            this.Cursor = Cursors.Default;

            using (PriceSingle askForm = new PriceSingle(rtb1, rtb2))
            {
                askForm.Text = "";
                askForm.ShowDialog();
                if (askForm.DialogResult != System.Windows.Forms.DialogResult.Yes)
                    return;

                modified = true;
                if ( allowSave )
                    btnSave.Show();
                rtb1.Rtf = askForm.rtf1;
                rtb2.Rtf = askForm.rtf2;

                byte[] b = Encoding.UTF8.GetBytes(rtb1.Rtf);
                dt.Rows[row]["header"] = b;
                //G1.update_blob("priceList", "record", record, "header", b);

                b = Encoding.UTF8.GetBytes(rtb2.Rtf);
                dt.Rows[row]["tail"] = b;
                //                G1.update_blob("priceList", "record", record, "tail", b);
            }
        }
        /***********************************************************************************************/
        public void AddParagraphMark(DevExpress.XtraRichEdit.RichEditControl rtb1)
        {
            string rtf = rtb1.Document.RtfText;
            //rtf = rtf.Replace("rtf1", "rtf1\\ansi This is in \\b0bold\\b0");
            //            rtf += "{\\b0}";
            // @"{\rtf1\ansi This is in \b bold\b0.}";
            rtb1.Document.RtfText = rtf;
        }
        /***********************************************************************************************/
        public void AddPageBreak(DevExpress.XtraRichEdit.RichEditControl rtb1)
        {
            string pb = DevExpress.Office.Characters.PageBreak.ToString();
            DocumentPosition pos = rtb1.Document.CaretPosition;
            rtb1.Document.InsertText(pos, pb);
        }
        /***********************************************************************************************/
        private void AddFuneralHeading(DevExpress.XtraRichEdit.RichEditControl rtb1, bool includeSubHeading = true)
        {
            DevExpress.XtraRichEdit.API.Native.Document document = rtb1.Document;
            //document.AppendSection();
            AddHeaderLine(rtb1, funeralHomeName, true, 18F, "Times New Roman");

            AddHeaderLine(rtb1, funeralAddress, false, 14F, "Times New Roman");
            if (!String.IsNullOrWhiteSpace(funeralPOBox))
                AddHeaderLine(rtb1, "P.O. Box " + funeralPOBox, false, 14F, "Times New Roman");
            string str = funeralCity + ", " + funeralState + "  " + funeralZip;
            AddHeaderLine(rtb1, str, false, 14F, "Times New Roman");
            AddHeaderLine(rtb1, funeralPhone, false, 14F, "Times New Roman");
            rtb1.Document.AppendText("\n");

            if (includeSubHeading)
            {
                str = workDescription;
                //if (workPrices.ToUpper() == "FUTURE")
                //    str += " (Future)";

                AddHeaderLine(rtb1, str, true, 22F, "Times New Roman");
                rtb1.Document.AppendText("\n");
            }

            //            document.AppendSection();
        }
        /***********************************************************************************************/
        private void LoadRtbExtra()
        {
            DataTable dt = (DataTable)dgv.DataSource;

            if (G1.get_column_number(dt, "top") < 0)
                dt.Columns.Add("top");
            if (G1.get_column_number(dt, "bottom") < 0)
                dt.Columns.Add("bottom");
            if (G1.get_column_number(dt, "RowHeight") < 0)
                dt.Columns.Add("RowHeight", Type.GetType("System.Int32"));

            string str = "";
            byte[] bytes = null;
            MemoryStream stream = null;
            DevExpress.XtraRichEdit.RichEditControl rtb = new RichEditControl();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["RowHeight"] = 1;
                bytes = dt.Rows[i]["header"].ObjToBytes();
                if (bytes != null)
                {
                    str = G1.ConvertToString(bytes);
                    if (!String.IsNullOrWhiteSpace(str))
                    {
                        stream = new MemoryStream(bytes);
                        rtb.Document.Delete(rtb.Document.Range);
                        rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);
                        str = rtb.Text.Trim();
                        if (!String.IsNullOrWhiteSpace(str))
                        {
                            dt.Rows[i]["top"] = rtb.Document.RtfText;
                        }
                    }
                }
                bytes = dt.Rows[i]["tail"].ObjToBytes();
                if (bytes != null)
                {
                    str = G1.ConvertToString(bytes);

                    if (!String.IsNullOrWhiteSpace(str))
                    {
                        stream = new MemoryStream(bytes);
                        rtb.Document.Delete(rtb.Document.Range);
                        rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);

                        str = rtb.Text;
                        if (!String.IsNullOrWhiteSpace(str))
                            dt.Rows[i]["bottom"] = rtb.Document.RtfText;
                    }
                }
            }
            dgv.DataSource = dt;
        }
        /***********************************************************************************************/
        private void TestPrint(object sender, EventArgs e)
        {
            //this.Cursor = Cursors.WaitCursor;
            //DataTable dx = (DataTable)dgv.DataSource;
            //PriceListPrint priceForm = new PriceListPrint(dx, this.dateTimePicker1.Value, this.chkAsOfDate.Checked);
            //priceForm.Show();
            //this.Cursor = Cursors.Default;
            PrintMethod(false);
        }
        /***********************************************************************************************/
        private void generateFormattedPriceListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintMethod(true);
        }
        /***********************************************************************************************/
        private void PrintMethod(bool normalPrintMethod)
        {
            this.Cursor = Cursors.WaitCursor;
            //DataTable dx = (DataTable)dgv.DataSource;
            //RibbonForm1 ribbonForm = new RibbonForm1(dx, this.dateTimePicker1.Value, this.chkAsOfDate.Checked);
            //ribbonForm.Show();
            //PriceListPrint priceForm = new PriceListPrint(dx, this.dateTimePicker1.Value, this.chkAsOfDate.Checked);
            //priceForm.Show();
            //this.Cursor = Cursors.Default;
            //if (1 == 1)
            //    return;
            DataTable dt = (DataTable)dgv.DataSource;

            if (G1.get_column_number(dt, "packageprice") < 0)
                dt.Columns.Add("packageprice", Type.GetType("System.Double"));

            DevExpress.XtraRichEdit.RichEditControl rtb1 = new RichEditControl();

            rtb1.Options.HorizontalRuler.Visibility = RichEditRulerVisibility.Hidden;
            rtb1.Options.VerticalRuler.Visibility = RichEditRulerVisibility.Hidden;

            rtb1.Document.Sections[0].Margins.Left = Units.InchesToDocumentsF(0.5f);
            rtb1.Document.Sections[0].Margins.Right = Units.InchesToDocumentsF(0.75f);
            rtb1.Document.Sections[0].Margins.Top = Units.InchesToDocumentsF(0.5f);
            rtb1.Document.Sections[0].Margins.Bottom = Units.InchesToDocumentsF(0.5f);

            rtb1.Document.Sections[0].PageNumbering.ContinueNumbering = true;
            rtb1.Document.Sections[0].PageNumbering.FirstPageNumber = 1;
            rtb1.Document.Sections[0].PageNumbering.NumberingFormat = NumberingFormat.CardinalText;
            rtb1.Document.Fields.Update();


            bool gotPrices = false;
            if (G1.get_column_number(dt, "price") >= 0)
                gotPrices = true;

            AddFuneralHeading(rtb1);

            string str = "";
            byte[] bytes = null;
            MemoryStream stream = null;
            DevExpress.XtraRichEdit.RichEditControl rtb = new RichEditControl();

            bool underline = false;
            bool bold = false;
            float size = 9f;
            int indent = 0;
            int ipad = 0;
            string sValue = "";
            string desc = "";
            double pad = 0D;
            double dvalue = 0D;
            double bvalue = 0D;
            string holdingStr = "";
            double majorPad = 95;
            int holdCount = 0;
            int layin = 0;
            string saveDesc = "";
            string found = "";
            char c = (char)127;
            string baseString = " ";
            string tail = "";
            bool got5Percent = false;
            int tailIdx = 0;
            var sb = new StringBuilder(baseString);
            sb[0] = c;
            baseString = sb.ToString();

            DateTime date1 = this.dateTimePicker1.Value;

            int year = date1.Year;
            int day = date1.Day;
            string month = date1.ToString("MMMMMMMMMMMMM");

            string asOfDate = month + " " + day.ToString() + ", " + year.ToString("D4");

            bool includeBoth = false;

            bool problem = false;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    got5Percent = false;
                    saveDesc = "";
                    underline = false;
                    bold = false;
                    size = 9f;
                    indent = 0;
                    layin = 0;
                    problem = false;
                    bytes = dt.Rows[i]["header"].ObjToBytes();
                    size = dt.Rows[i]["size"].ObjToFloat();
                    if (size <= 0f)
                        size = 9f;
                    str = dt.Rows[i]["underline"].ObjToString();
                    if (str.ToUpper().IndexOf("Y") >= 0)
                        underline = true;
                    str = dt.Rows[i]["bold"].ObjToString();
                    if (str.ToUpper().IndexOf("Y") >= 0)
                        bold = true;
                    indent = dt.Rows[i]["indent"].ObjToInt32();
                    if (indent <= 0)
                        indent = 0;
                    layin = dt.Rows[i]["layin"].ObjToInt32();
                    if (layin <= 0)
                        layin = 0;
                    found = dt.Rows[i]["found"].ObjToString();
                    if (bytes != null)
                    {
                        str = G1.ConvertToString(bytes);

                        if (!String.IsNullOrWhiteSpace(str))
                        {
                            stream = new MemoryStream(bytes);
                            rtb.Document.Delete(rtb.Document.Range);
                            rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);
                            str = rtb.Text.Trim();
                            if (chkAsOfDate.Checked)
                            {
                                int idx = str.IndexOf("as of");
                                if (idx > 0)
                                {
                                    int xdx = str.IndexOf(" but");
                                    if (xdx > 0)
                                    {
                                        string text = str.Substring(idx, xdx - idx);
                                        rtb.RtfText = rtb.RtfText.Replace(text, "as of " + asOfDate);
                                        str = rtb.Text.Trim();
                                    }
                                }
                            }
                            if (!String.IsNullOrWhiteSpace(str))
                            {
                                rtb1.Document.AppendRtfText(rtb.Document.RtfText);
                            }
                            //rtb1.Document.AppendText("\n");
                        }
                    }
                    str = dt.Rows[i]["service"].ObjToString();
                    if (str.ToUpper().IndexOf("OTHER") == 0)
                    {
                    }
                    if (str == ".")
                        str = baseString;
                    if (!String.IsNullOrWhiteSpace(str))
                    {
                        if (str.Trim().ToUpper() == "{HEADER}")
                        {
                            AddPageBreak(rtb1);
                            AddFuneralHeading(rtb1);
                        }
                        else if (str.Trim().ToUpper() == "{SUB-HEADER}")
                        {
                            AddPageBreak(rtb1);
                            AddFuneralHeading(rtb1, false);
                        }
                        else if (str.Trim().ToUpper() == "{BREAK}")
                            AddPageBreak(rtb1);
                        else if (str.Trim().ToUpper() == "{EMPTY}")
                            str = baseString;
                        else if (str.Trim().ToUpper() == "{INCLUDE BOTH PRICES}")
                        {
                            includeBoth = true;
                        }
                        else if (str.Trim().ToUpper().IndexOf("{RANGE}") >= 0)
                        {
                            int idx = str.IndexOf('}');
                            str = str.Substring(idx + 1);
                            holdingStr = str.Trim();
                            pad = majorPad;
                            if (size != 9F)
                            {
                                pad = majorPad / size * 9D;
                                pad = pad - 1D;
                            }
                            pad = G1.RoundDown(pad);
                            pad = pad + indent;
                            ipad = Convert.ToInt32(pad);

                            holdingStr = holdingStr.PadRight(ipad);
                        }
                        else
                        {
                            dvalue = 3095D;
                            sValue = "$" + G1.ReformatMoney(dvalue);
                            if (gotPrices)
                            {
                                desc = dt.Rows[i]["price"].ObjToString();
                                if (!String.IsNullOrWhiteSpace(desc))
                                {
                                    dvalue = dt.Rows[i]["price"].ObjToDouble();
                                    sValue = "$" + G1.ReformatMoney(dvalue);
                                    if (includeBoth)
                                    {
                                        bvalue = dt.Rows[i]["packageprice"].ObjToDouble();
                                        if (bvalue > 9999.99D)
                                        {
                                            problem = true;
                                            sValue = " " + sValue;
                                            sValue += "    $" + G1.ReformatMoney(bvalue);
                                        }
                                        else
                                            sValue += "     $" + G1.ReformatMoney(bvalue);
                                    }
                                }
                                else
                                    sValue = "";
                            }
                            else
                            {
                                if (includeBoth)
                                {
                                    //sValue += "     $" + G1.ReformatMoney(dvalue);
                                }
                            }
                            if (found != "Y")
                            {
                                if (sValue == "$0.00")
                                    sValue = baseString;
                                else if (sValue == "$0.00     $0.00")
                                    sValue = baseString;
                            }

                            desc = "";
                            pad = majorPad;
                            if (size != 9F)
                            {
                                pad = majorPad / size * 9D;
                                pad = pad - 1D;
                            }
                            pad = G1.RoundDown(pad);
                            //pad = Math.Truncate(pad);
                            pad = pad + indent;
                            ipad = Convert.ToInt32(pad);

                            desc = desc.PadRight(ipad);
                            if (!String.IsNullOrWhiteSpace(holdingStr))
                            {
                                if (holdCount == 0)
                                {
                                    holdingStr = holdingStr.TrimEnd();
                                    holdingStr += "          " + sValue + "   to   ";
                                    holdCount++;
                                }
                                else
                                {
                                    desc = "";
                                    holdingStr += sValue;
                                    if (indent > 0)
                                        desc = " ".PadRight(indent);
                                    desc += holdingStr;
                                    AddNormalLine(rtb1, desc + "\n", bold, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                    //AddNormalLine(rtb1, desc + "\n", bold, size, "Times New Roman", ParagraphAlignment.Left, underline);
                                    holdCount = 0;
                                    holdingStr = "";
                                }
                            }
                            else
                            {
                                if (layin <= 0)
                                {
                                    tail = GetTail(dt, i);
                                    if (!String.IsNullOrWhiteSpace(tail))
                                        str = str.Replace(tail, "");
                                    else if (str.IndexOf("5% Of Total Assignment") > 0)
                                    {
                                        got5Percent = true;
                                        str = str.Replace("5% Of Total Assignment", "");
                                        sValue = "5% Of Total Assignment";
                                    }
                                    desc = G1.lay_in_string(desc, str, indent, str.Length);
                                    int ll = desc.Length;
                                    //sValue += "~";
                                    //if (problem)
                                    //    indent = indent - 1;
                                    if (got5Percent)
                                        desc += "  ";
                                    desc = G1.lay_in_string(desc, sValue, desc.Length - (sValue.Length + indent), sValue.Length);
                                    ll = desc.Length;
                                    if (desc.IndexOf("{") < 0 && desc.IndexOf("}") < 0)
                                    {
                                        AddNormalLine(rtb1, desc + "\n", bold, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                        //AddNormalLine(rtb1, desc + "\n", bold, size, "Times New Roman", ParagraphAlignment.Left, underline);
                                        AddParagraphMark(rtb1);
                                    }
                                }
                                else
                                    saveDesc = str;
                                G1.Toggle_Bold(rtb1, false, false);
                            }
                        }
                    }
                    //AddNormalText(rtb1, str + "\n", false, 12f);
                    //rtb1.Document.AppendText(str + "\n");

                    bytes = dt.Rows[i]["tail"].ObjToBytes();
                    if (bytes != null)
                    {
                        str = G1.ConvertToString(bytes);

                        if (!String.IsNullOrWhiteSpace(str))
                        {
                            stream = new MemoryStream(bytes);
                            rtb.Document.Delete(rtb.Document.Range);
                            rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);

                            str = rtb.Text;
                            if (!String.IsNullOrWhiteSpace(str))
                            {
                                if (layin > 0)
                                {
                                    pad = majorPad;
                                    if (size != 9F)
                                    {
                                        pad = majorPad / size * 9D;
                                        pad = pad - 1D;
                                    }
                                    pad = G1.RoundDown(pad);
                                    pad = pad + indent;
                                    ipad = Convert.ToInt32(pad);
                                    //desc = G1.lay_in_string(saveDesc, str, layin, str.Length);
                                    desc = desc.PadRight(layin);
                                    //SetFont(rtb1, "Lucida Console", size, bold, false);
                                    saveDesc = saveDesc.PadRight(layin);
                                    AddNormal(rtb1, saveDesc, bold, size, "Lucida Console", ParagraphAlignment.Left, underline);

                                    //rtb1.Document.AppendText(saveDesc);
                                    //SetFont(rtb1, "Lucida Console", size, false, false);
                                    AddNormal(rtb1, str, false, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                    //rtb1.Document.AppendText(str);
                                    sValue = sValue.PadLeft(ipad - layin - str.Length);
                                    //                                SetFont(rtb1, "Lucida Console", size, bold, false);
                                    AddNormal(rtb1, sValue + "\n", bold, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                    rtb1.Document.AppendText("\n");
                                    //rtb1.Document.AppendText(sValue + "\n");
                                    //AddNormalText(rtb1, desc, bold, size);
                                    //str = str.PadRight(ipad - layin);
                                    //SetFont(rtb1, "Lucida Console", size, false, false);
                                    //AddNormalText(rtb1, str, false, size);
                                    //SetFont(rtb1, "Lucida Console", size, bold, false);
                                    //AddNormalText(rtb1, sValue + "\n", bold, size);
                                    //AddNormalLine(rtb1, desc + "\n", bold, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                }
                                else
                                {
                                    //AddParagraphMark(rtb);
                                    rtb1.Document.AppendRtfText(rtb.Document.RtfText);
                                }
                            }
                            else if (str == "\r\n")
                                rtb1.Document.AppendText("\n");
                        }
                    }
                }
                catch (Exception ex)
                {

                }
            }


            str = rtb1.Document.RtfText;

            if (!normalPrintMethod)
            {
                this.Cursor = Cursors.WaitCursor;
                ViewRTF aForm = new ViewRTF(str);
                aForm.Show();
                this.Cursor = Cursors.Default;
                return;
            }

            this.Cursor = Cursors.WaitCursor;

            rtb9 = rtb1;

            //Section firstSection = rtb9.Document.Sections[0];
            //// Create an empty header. 
            //SubDocument newHeader = firstSection.BeginUpdateFooter();
            //newHeader.AppendText("Page - [Page #]");
            //firstSection.EndUpdateFooter(newHeader);
            //// Check whether the document already has a header (the same header for all pages). 
            //if (firstSection.HasHeader(HeaderFooterType.Primary))
            //{
            //    SubDocument headerDocument = firstSection.BeginUpdateFooter();
            //    headerDocument.AppendText("Page - [Page #]");
            //    rtb9.Document.ChangeActiveDocument(headerDocument);
            //    rtb9.Document.CaretPosition = headerDocument.CreatePosition(0);
            //    firstSection.EndUpdateFooter(headerDocument);
            //}

            //rtb9.Document.AppendSection();
            Section firstSection = rtb9.Document.Sections[0];
            // Modify the header of the HeaderFooterType.First type. 
            SubDocument myHeader = firstSection.BeginUpdateFooter(HeaderFooterType.Primary);
            DocumentRange range = myHeader.InsertText(myHeader.CreatePosition(0), " PAGE ");
            Field fld = myHeader.Fields.Create(range.End, "PAGE \\* ARABICDASH");
            myHeader.Fields.Update();
            firstSection.EndUpdateFooter(myHeader);
            // Display the header of the HeaderFooterType.First type on the first page. 
            firstSection.DifferentFirstPage = false;

            //rtb9.Document.AppendSection();
            firstSection = rtb9.Document.Sections[0];
            // Modify the header of the HeaderFooterType.First type. 
            myHeader = firstSection.BeginUpdateFooter(HeaderFooterType.First);
            range = myHeader.InsertText(myHeader.CreatePosition(0), " PAGE ");
            fld = myHeader.Fields.Create(range.End, "PAGE \\* ARABICDASH");
            myHeader.Fields.Update();
            firstSection.EndUpdateFooter(myHeader);
            // Display the header of the HeaderFooterType.First type on the first page. 
            firstSection.DifferentFirstPage = false;


            link = new PrintableComponentLink(new PrintingSystem());

            link.Component = rtb9;
            //link.PrintingSystem.AfterMarginsChange += new MarginsChangeEventHandler(PrintingSystem_AfterMarginsChange);
            //link.PrintingSystem.AfterBuildPages += PrintingSystem_AfterBuildPages;
            PrintTool tool = new PrintTool(link.PrintingSystem);

            //tool.PrintingSystem.PageMargins.Left = 5;
            //link.MinMargins.Left = 5;
            //link.MinMargins.Right = 5;
            //link.MinMargins.Top = 5;
            //link.MinMargins.Bottom = 5;

            //link.Margins.Left = 10;
            //link.Margins.Right = 10;
            //link.Margins.Top = 10;
            //link.Margins.Bottom = 10;

            //link.PrintingSystem.PageMargins.Left = 5;
            //link.PrintingSystem.PageMargins.Top = 5;
            //link.PrintingSystem.PageMargins.Bottom = 5;
            //link.PrintingSystem.PageMargins.Top = 5;

            link.CreateDocument();

            //link.MinMargins.Left = 5;
            //link.MinMargins.Right = 5;
            //link.MinMargins.Top = 5;
            //link.MinMargins.Bottom = 5;

            //link.Margins.Left = 10;
            //link.Margins.Right = 10;
            //link.Margins.Top = 10;
            //link.Margins.Bottom = 10;

            //link.PrintingSystem.PageMargins.Left = 5;
            //link.PrintingSystem.PageMargins.Top = 5;
            //link.PrintingSystem.PageMargins.Bottom = 5;
            //link.PrintingSystem.PageMargins.Top = 5;

            //            link.ShowPreviewDialog();
            link.ShowPreview();
            //tool.ShowPreview();
            //tool.ShowPreviewDialog();
            //rtb9.ShowPrintPreview();

            ////ViewRTF aForm = new ViewRTF(str);
            ////aForm.Show();
            this.Cursor = Cursors.Default;
        }
        /***************************************************************************************/
        private string GetTail(DataTable dt, int i)
        {
            string tail = "";
            string str = "";
            byte[] bytes = null;
            MemoryStream stream = null;
            DevExpress.XtraRichEdit.RichEditControl rtb = new RichEditControl();
            bytes = dt.Rows[i]["tail"].ObjToBytes();
            if (bytes != null)
            {
                str = G1.ConvertToString(bytes);

                if (!String.IsNullOrWhiteSpace(str))
                {
                    stream = new MemoryStream(bytes);
                    rtb.Document.Delete(rtb.Document.Range);
                    rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);

                    tail = rtb.Text;
                    tail = tail.Replace("\r", "");
                    tail = tail.Replace("\n", "");
                }
            }
            return tail;
        }
        /***************************************************************************************/
        void PrintingSystem_AfterBuildPages(object sender, EventArgs e)
        {
            link.PrintingSystem.SetCommandVisibility(PrintingSystemCommand.PageMargins, CommandVisibility.All);
        }
        void PrintingSystem_AfterMarginsChange(object sender, MarginsChangeEventArgs e)
        {
            // Change document margins in the source RichEditControl
            SectionMargins margins = rtb9.Document.Sections[0].Margins;
            switch (e.Side)
            {
                case MarginSide.Left:
                    margins.Left = Units.HundredthsOfInchToDocuments((int)e.Value);
                    break;
                case MarginSide.Right:
                    margins.Right = Units.HundredthsOfInchToDocuments((int)e.Value);
                    break;
                case MarginSide.Top:
                    margins.Top = Units.HundredthsOfInchToDocuments((int)e.Value);
                    break;
                case MarginSide.Bottom:
                    margins.Bottom = Units.HundredthsOfInchToDocuments((int)e.Value);
                    break;
                default:
                    break;
            }
            link.CreateDocument();
        }
        /***************************************************************************************/
        private void AddHeader(DevExpress.XtraRichEdit.RichEditControl rtb, string str, bool bold, float size, string fontname = "")
        {
            try
            {
                rtb.Document.BeginUpdate();

                //The target range is the first paragraph
                DocumentPosition pos = rtb.Document.CaretPosition;
                DocumentRange range = rtb.Document.CreateRange(pos, 0);

                SetFont(rtb, fontname, size);

                // Create and customize an object  
                // that sets character formatting for the selected range
                ParagraphProperties pp = rtb.Document.BeginUpdateParagraphs(range);
                // Center paragraph
                pp.Alignment = ParagraphAlignment.Center;
                // Set single spacing
                pp.LineSpacingType = ParagraphLineSpacing.Multiple;
                pp.LineSpacingMultiplier = 1;
                // Set left indent at 0.0".
                // Default unit is 1/300 of an inch (a document unit).
                pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
                // Set tab stop at 1.0"
                TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);
                TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();
                tbi.Alignment = TabAlignmentType.Center;
                tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
                tbiColl.Add(tbi);
                pp.EndUpdateTabs(tbiColl);
                rtb.Document.EndUpdateParagraphs(pp);

                int j = rtb.Document.Sections.Count - 1;

                Section section = rtb.Document.Sections[0];
                SubDocument headerDocument = section.BeginUpdateHeader();
                rtb.Document.ChangeActiveDocument(headerDocument);
                // rtb.Document.CaretPosition = headerDocument.CreatePosition(0);


                string[] Lines = str.Split('\n');
                for (int i = 0; i < Lines.Length; i++)
                {
                    str = Lines[i].Trim();
                    headerDocument.AppendText(str + "\n");
                }

                section.EndUpdateHeader(headerDocument);
                section.DifferentFirstPage = true;
                section.LinkHeaderToPrevious();
                rtb.Document.EndUpdate();
            }
            catch (Exception ex)
            {
            }
        }
        /***************************************************************************************/
        private void AddHeaderLine(DevExpress.XtraRichEdit.RichEditControl rtb, string str, bool bold, float size, string fontname = "")
        {
            try
            {
                rtb.Document.BeginUpdate();

                //The target range is the first paragraph
                DocumentPosition pos = rtb.Document.CaretPosition;
                DocumentRange range = rtb.Document.CreateRange(pos, 0);

                SetFont(rtb, fontname, size);

                // Create and customize an object  
                // that sets character formatting for the selected range
                ParagraphProperties pp = rtb.Document.BeginUpdateParagraphs(range);
                // Center paragraph
                pp.Alignment = ParagraphAlignment.Center;
                // Set single spacing
                pp.LineSpacingType = ParagraphLineSpacing.Multiple;
                pp.LineSpacingMultiplier = 1;
                // Set left indent at 0.0".
                // Default unit is 1/300 of an inch (a document unit).
                pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
                // Set tab stop at 1.0"
                TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);
                TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();
                tbi.Alignment = TabAlignmentType.Center;
                tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
                tbiColl.Add(tbi);
                pp.EndUpdateTabs(tbiColl);
                rtb.Document.EndUpdateParagraphs(pp);



                string[] Lines = str.Split('\n');
                for (int i = 0; i < Lines.Length; i++)
                {
                    str = Lines[i].Trim();
                    rtb.Document.AppendText(str + "\n");
                }

                rtb.Document.EndUpdate();
            }
            catch (Exception ex)
            {
            }
        }
        /***************************************************************************************/
        private void SetFont(DevExpress.XtraRichEdit.RichEditControl rtb, string fontname, float size, bool bold = false, bool underline = false)
        {
            //if (!String.IsNullOrWhiteSpace(fontname))
            //{
            //    Paragraph currentParagraph = rtb.Document.GetParagraph(rtb.Document.CaretPosition);

            //    ParagraphProperties ppp = rtb.Document.BeginUpdateParagraphs(currentParagraph.Range);
            //    ppp.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Left;
            //    //                ppp.BackColor = Color.LightGray;
            //    rtb.Document.EndUpdateParagraphs(ppp);
            //    CharacterProperties cp = rtb.Document.BeginUpdateCharacters(currentParagraph.Range);
            //    cp.Bold = bold;
            //    if (underline == false)
            //        cp.Underline = UnderlineType.None;
            //    else
            //        cp.Underline = UnderlineType.Single;
            //    cp.FontSize = size;
            //    cp.FontName = fontname;
            //    rtb.Document.EndUpdateCharacters(cp);
            //}
            //else
            //{
            //    Paragraph currentParagraph = rtb.Document.GetParagraph(rtb.Document.CaretPosition);

            //    ParagraphProperties ppp = rtb.Document.BeginUpdateParagraphs(currentParagraph.Range);
            //    ppp.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Left;
            //    //                ppp.BackColor = Color.LightGray;
            //    rtb.Document.EndUpdateParagraphs(ppp);
            //    CharacterProperties cp = rtb.Document.BeginUpdateCharacters(currentParagraph.Range);
            //    cp.Bold = bold;
            //    cp.FontSize = size;
            //    cp.FontName = "Times New Roman";
            //    rtb.Document.EndUpdateCharacters(cp);
            //}
        }
        /***************************************************************************************/
        private void SetFontText(DevExpress.XtraRichEdit.RichEditControl rtb, string str, string fontname, float size, bool bold = false, bool underline = false)
        {
            //if (!String.IsNullOrWhiteSpace(fontname))
            //{
            //    rtb.Document.BeginUpdate();
            //    Paragraph currentParagraph = rtb.Document.GetParagraph(rtb.Document.CaretPosition);

            //    //                DocumentPosition pos = rtb.Document.Range.Start;
            //    DocumentPosition pos = rtb.Document.CaretPosition;
            //    int start = rtb.Document.Length;

            //    string[] Lines = str.Split('\n');
            //    for (int i = 0; i < Lines.Length; i++)
            //    {
            //        //                    str = Lines[i];
            //        if (!String.IsNullOrWhiteSpace(Lines[i]))
            //        {
            //            str = Lines[i];
            //            rtb.Document.AppendText(Lines[i]);
            //        }
            //    }
            //    //DocumentRange range = rtb.Document.CreateRange(pos, str.Length);
            //    DocumentPosition pos2 = rtb.Document.CaretPosition;
            //    int stop = rtb.Document.Length - 1;
            //    int len = stop;
            //    //int len = currentParagraph.Range.Length;
            //    DocumentRange range = rtb.Document.CreateRange(len - str.Length, str.Length);

            //    //ParagraphProperties ppp = rtb.Document.BeginUpdateParagraphs(currentParagraph.Range);
            //    ParagraphProperties ppp = rtb.Document.BeginUpdateParagraphs(range);
            //    ppp.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Left;
            //    ppp.Style.Bold = bold;
            //    CharacterProperties cp = rtb.Document.BeginUpdateCharacters(range);
            //    cp.Bold = bold;
            //    if (underline == false)
            //        cp.Underline = UnderlineType.None;
            //    else
            //        cp.Underline = UnderlineType.Single;
            //    cp.FontSize = size;
            //    cp.FontName = fontname;
            //    rtb.Document.EndUpdateCharacters(cp);
            //    rtb.Document.EndUpdateParagraphs(ppp);
            //    rtb.Document.EndUpdate();
            //}
            //else
            //{
            //    Paragraph currentParagraph = rtb.Document.GetParagraph(rtb.Document.CaretPosition);

            //    ParagraphProperties ppp = rtb.Document.BeginUpdateParagraphs(currentParagraph.Range);
            //    ppp.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Left;
            //    //                ppp.BackColor = Color.LightGray;
            //    rtb.Document.EndUpdateParagraphs(ppp);
            //    CharacterProperties cp = rtb.Document.BeginUpdateCharacters(currentParagraph.Range);
            //    cp.Bold = bold;
            //    cp.FontSize = size;
            //    cp.FontName = "Times New Roman";
            //    rtb.Document.EndUpdateCharacters(cp);
            //}
        }
        /***********************************************************************************************/
        private void AddNormalLine(DevExpress.XtraRichEdit.RichEditControl rtb, string str, bool bold, float size, string fontname = "", ParagraphAlignment align = ParagraphAlignment.Left, bool underline = false)
        {
            SetFontText(rtb, str, fontname, size, bold, underline);
            rtb.Document.AppendText("\n");
            if (1 == 1)
                return;
            rtb.Document.BeginUpdate();

            //The target range is the first paragraph
            DocumentPosition pos = rtb.Document.CaretPosition;
            DocumentRange range = rtb.Document.CreateRange(pos, 0);

            //SetFont(rtb, fontname, size, bold, underline );

            // Create and customize an object  
            // that sets character formatting for the selected range
            ParagraphProperties pp = rtb.Document.BeginUpdateParagraphs(range);
            pp.Style.FontName = fontname;
            pp.Style.FontSize = size;
            pp.Style.Bold = bold;
            // Center paragraph
            //                pp.Alignment = ParagraphAlignment.Left;
            pp.Alignment = align;
            // Set triple spacing
            pp.LineSpacingType = ParagraphLineSpacing.Multiple;
            pp.LineSpacingMultiplier = 1;
            // Set left indent at 0.5".
            // Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
            // Set tab stop at 1.5"
            TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);
            TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();
            tbi.Alignment = TabAlignmentType.Left;
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
            tbiColl.Add(tbi);
            pp.EndUpdateTabs(tbiColl);

            ParagraphStyle chapterStyle = rtb.Document.ParagraphStyles.CreateNew();
            chapterStyle.Name = "MyTitleStyle";
            chapterStyle.ForeColor = Color.Black;
            chapterStyle.FontSize = size;
            chapterStyle.FontName = "Times New Roman";
            if (!String.IsNullOrWhiteSpace(fontname))
            {
                chapterStyle.FontSize = size;
                chapterStyle.FontName = fontname;
                chapterStyle.Name = fontname;
            }
            //            chapterStyle.Alignment = ParagraphAlignment.Left;
            chapterStyle.Alignment = align;
            chapterStyle.SpacingBefore = Units.InchesToDocumentsF(0.0f);
            chapterStyle.SpacingAfter = Units.InchesToDocumentsF(0.0f);
            chapterStyle.OutlineLevel = 1;
            chapterStyle.Bold = bold;

            //Add the object to the document collection
            rtb.Document.ParagraphStyles.Add(chapterStyle);
            chapterStyle.Name = fontname;

            pp.Style = chapterStyle;

            //G1.Toggle_Bold(rtb, bold, false);

            string[] Lines = str.Split('\n');
            for (int i = 0; i < Lines.Length; i++)
            {
                str = Lines[i];
                //if ( align == ParagraphAlignment.Left )
                //    str = Lines[i].Trim();
                if (!String.IsNullOrWhiteSpace(str))
                {
                    rtb.Document.AppendText(str + "\n");
                }
            }

            rtb.Document.EndUpdateParagraphs(pp);
            rtb.Document.EndUpdate();
        }
        /***********************************************************************************************/
        //private void AddNormalLine(DevExpress.XtraRichEdit.RichEditControl rtb, string str, bool bold, float size, string fontname = "", ParagraphAlignment align = ParagraphAlignment.Left, bool underline = false )
        //{
        //    rtb.Document.BeginUpdate();

        //    //The target range is the first paragraph
        //    DocumentPosition pos = rtb.Document.CaretPosition;
        //    DocumentRange range = rtb.Document.CreateRange(pos, 0);

        //    //SetFont(rtb, fontname, size, bold, underline );

        //    // Create and customize an object  
        //    // that sets character formatting for the selected range
        //    ParagraphProperties pp = rtb.Document.BeginUpdateParagraphs(range);
        //    pp.Style.FontName = fontname;
        //    pp.Style.FontSize = size;
        //    pp.Style.Bold = bold;
        //    // Center paragraph
        //    //                pp.Alignment = ParagraphAlignment.Left;
        //    pp.Alignment = align;
        //    // Set triple spacing
        //    pp.LineSpacingType = ParagraphLineSpacing.Multiple;
        //    pp.LineSpacingMultiplier = 1;
        //    // Set left indent at 0.5".
        //    // Default unit is 1/300 of an inch (a document unit).
        //    pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
        //    // Set tab stop at 1.5"
        //    TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);
        //    TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();
        //    tbi.Alignment = TabAlignmentType.Left;
        //    tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
        //    tbiColl.Add(tbi);
        //    pp.EndUpdateTabs(tbiColl);

        //    ParagraphStyle chapterStyle = rtb.Document.ParagraphStyles.CreateNew();
        //    chapterStyle.Name = "MyTitleStyle";
        //    chapterStyle.ForeColor = Color.Black;
        //    chapterStyle.FontSize = size;
        //    chapterStyle.FontName = "Times New Roman";
        //    if (!String.IsNullOrWhiteSpace(fontname))
        //    {
        //        chapterStyle.FontSize = size;
        //        chapterStyle.FontName = fontname;
        //        chapterStyle.Name = fontname;
        //    }
        //    //            chapterStyle.Alignment = ParagraphAlignment.Left;
        //    chapterStyle.Alignment = align;
        //    chapterStyle.SpacingBefore = Units.InchesToDocumentsF(0.0f);
        //    chapterStyle.SpacingAfter = Units.InchesToDocumentsF(0.0f);
        //    chapterStyle.OutlineLevel = 1;
        //    chapterStyle.Bold = bold;

        //    //Add the object to the document collection
        //    rtb.Document.ParagraphStyles.Add(chapterStyle);
        //    chapterStyle.Name = fontname;

        //    pp.Style = chapterStyle;

        //    //G1.Toggle_Bold(rtb, bold, false);

        //    string[] Lines = str.Split('\n');
        //    for (int i = 0; i < Lines.Length; i++)
        //    {
        //        str = Lines[i];
        //        //if ( align == ParagraphAlignment.Left )
        //        //    str = Lines[i].Trim();
        //        if (!String.IsNullOrWhiteSpace(str))
        //        {
        //            rtb.Document.AppendText(str + "\n");
        //        }
        //    }

        //    rtb.Document.EndUpdateParagraphs(pp);
        //    rtb.Document.EndUpdate();
        //}
        /***********************************************************************************************/
        private void DoSomething(DevExpress.XtraRichEdit.RichEditControl rtb)
        {
            rtb.Document.BeginUpdate();
            rtb.Document.AppendText("Modified Paragraph\nNormal\nNormal");
            rtb.Document.EndUpdate();

            //The target range is the first paragraph 
            DocumentPosition pos = rtb.Document.Range.Start;
            DocumentRange range = rtb.Document.CreateRange(pos, 0);

            // Create and customize an object   
            // that sets character formatting for the selected range 
            ParagraphProperties pp = rtb.Document.BeginUpdateParagraphs(range);
            // Center paragraph 
            pp.Alignment = ParagraphAlignment.Center;
            // Set triple spacing 
            pp.LineSpacingType = ParagraphLineSpacing.Multiple;
            pp.LineSpacingMultiplier = 3;
            // Set left indent at 0.5". 
            // Default unit is 1/300 of an inch (a document unit). 
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.5f);
            // Set tab stop at 1.5" 
            TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);
            TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();
            tbi.Alignment = TabAlignmentType.Center;
            tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);
            tbiColl.Add(tbi);
            pp.EndUpdateTabs(tbiColl);

            //Finalize modifications 
            // with this method call 
            rtb.Document.EndUpdateParagraphs(pp);
        }
        /***********************************************************************************************/
        private void AddNormal(DevExpress.XtraRichEdit.RichEditControl rtb, string str, bool bold, float size, string fontname = "", ParagraphAlignment align = ParagraphAlignment.Left, bool underline = false)
        {
            SetFontText(rtb, str, fontname, size, bold, underline);
            if (1 == 1)
                return;

            rtb.Document.BeginUpdate();

            //The target range is the first paragraph
            rtb.Document.Paragraphs.Append();
            DocumentPosition pos = rtb.Document.CaretPosition;
            DocumentRange range = rtb.Document.CreateRange(pos, str.Length);

            SetFont(rtb, fontname, size, bold, underline);

            // Create and customize an object  
            // that sets character formatting for the selected range
            ParagraphProperties pp = rtb.Document.BeginUpdateParagraphs(range);
            pp.Style.FontName = fontname;
            pp.Style.FontSize = size;
            pp.Style.Bold = bold;
            // Center paragraph
            //                pp.Alignment = ParagraphAlignment.Left;
            pp.Alignment = align;
            // Set triple spacing
            pp.LineSpacingType = ParagraphLineSpacing.Multiple;
            pp.LineSpacingMultiplier = 1;
            // Set left indent at 0.5".
            // Default unit is 1/300 of an inch (a document unit).
            pp.LeftIndent = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
            // Set tab stop at 1.5"
            //TabInfoCollection tbiColl = pp.BeginUpdateTabs(true);
            //TabInfo tbi = new DevExpress.XtraRichEdit.API.Native.TabInfo();
            //tbi.Alignment = TabAlignmentType.Left;
            //tbi.Position = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.0f);
            //tbiColl.Add(tbi);
            //pp.EndUpdateTabs(tbiColl);

            ParagraphStyle chapterStyle = rtb.Document.ParagraphStyles.CreateNew();
            chapterStyle.Name = "MyTitleStyle";
            chapterStyle.ForeColor = Color.Black;
            chapterStyle.FontSize = size;
            chapterStyle.FontName = "Times New Roman";
            if (!String.IsNullOrWhiteSpace(fontname))
            {
                chapterStyle.FontSize = size;
                chapterStyle.FontName = fontname;
                chapterStyle.Name = fontname;
            }
            //            chapterStyle.Alignment = ParagraphAlignment.Left;
            chapterStyle.Alignment = align;
            chapterStyle.SpacingBefore = Units.InchesToDocumentsF(0.0f);
            chapterStyle.SpacingAfter = Units.InchesToDocumentsF(0.0f);
            chapterStyle.OutlineLevel = 1;
            chapterStyle.Bold = bold;

            //Add the object to the document collection
            rtb.Document.ParagraphStyles.Add(chapterStyle);
            chapterStyle.Name = fontname;

            pp.Style = chapterStyle;

            //G1.Toggle_Bold(rtb, bold, false); 

            string[] Lines = str.Split('\n');
            for (int i = 0; i < Lines.Length; i++)
            {
                str = Lines[i];
                //if ( align == ParagraphAlignment.Left )
                //    str = Lines[i].Trim();
                if (!String.IsNullOrWhiteSpace(str))
                    rtb.Document.AppendText(str);
            }

            //rtb.Document.EndUpdateParagraphs(pp);
            int count = rtb.Document.Paragraphs.Count;

            rtb.Document.EndUpdate();
            rtb.Document.Paragraphs[count - 1].Style = chapterStyle;

        }
        /***********************************************************************************************/
        private void AddNormalText(DevExpress.XtraRichEdit.RichEditControl rtb, string str, bool bold, float size)
        {
            if (bold)
                G1.Toggle_Bold(rtb, true, false);
            else
                G1.Toggle_Bold(rtb, false, false);
            rtb.Document.DefaultParagraphProperties.Alignment = DevExpress.XtraRichEdit.API.Native.ParagraphAlignment.Left;
            rtb.Document.AppendText(str);
        }
        /***********************************************************************************************/
        private void btnSave_Click(object sender, EventArgs e)
        {
            SavePriceLists();
            modified = false;
            btnSave.Visible = false;
        }
        /***********************************************************************************************/
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Caskets casketForm = new Caskets(true);
            casketForm.SelectDone += CasketForm_SelectDone;
            casketForm.Show();
        }
        private void CasketForm_SelectDone(DataTable dt)
        {
            workDt = (DataTable)dgv.DataSource;

            string select = "";
            string service = "";

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                select = dt.Rows[i]["select"].ObjToString();
                if (select == "1")
                {
                    service = dt.Rows[i]["casketdesc"].ObjToString();
                    DataRow newRow = workDt.NewRow();
                    newRow["service"] = service;
                    workDt.Rows.Add(newRow);
                }
            }
            dgv.DataSource = workDt;
            dgv.Refresh();
            modified = true;
            if ( allowSave)
                btnSave.Show();
        }
        /***********************************************************************************************/
        private void btnFuture_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgv.DataSource;
            if (workPrices.ToUpper() == "CURRENT")
            {
                workPrices = "FUTURE";
                btnFuture.Text = "Show Current";
                btnFuture.BackColor = Color.Lime;

                LoadData();

                //dt = (DataTable)dgv.DataSource;
                //dt = SetupPrices(dt);
                //dgv.DataSource = dt;
                //dgv.Refresh();
            }
            else
            {
                workPrices = "CURRENT";
                btnFuture.Text = "Show Future";
                btnFuture.BackColor = Color.Tomato;

                PriceList_Load(null, null);
                //LoadData();

                //dt = SetupPrices(dt);
                //dgv.DataSource = dt;
                //dgv.Refresh();
            }
        }
        /***********************************************************************************************/
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if ( chkAsOfDate.Checked )
            {
                string group = workGroupName;
                DateTime date = dateTimePicker1.Value;
                string cmd = "Select * from `effectivedates` where `category` = '" + workDescription + "';";
                DataTable dt = G1.get_db_data(cmd);
                string record = "";
                if (dt.Rows.Count > 0)
                    record = dt.Rows[0]["record"].ObjToString();
                else
                {
                    record = G1.create_record("effectivedates", "category", "-1");
                    if (G1.BadRecord("effectivedates", record))
                        return;
                    G1.update_db_table("effectivedates", "record", record, new string[] { "category", workDescription, "effectiveDate", date.ToString("yyyy-MM-dd") });
                }
                //G1.update_db_table("effectivedates", "record", record, new string[] {"category", workDescription, "effectiveDate", date.ToString("yyyy-MM-dd") });
            }
            if (!chkUseNewVersion.Checked)
                generateFormattedPriceListToolStripMenuItem_Click(null, null);
            else
            {
                //string cmd = "Select * from `casket_master`;";
                //DataTable dt = G1.get_db_data(cmd);
                DataTable dt = (DataTable)dgv.DataSource;
                CreateReport3(dt);
            }
        }
        /***********************************************************************************************/
        private void btnAdd_Click(object sender, EventArgs e)
        {
            DataRow dRow = null;
            DataTable dt = (DataTable)dgv.DataSource;
            dRow = dt.NewRow();
            dt.Rows.Add(dRow);
            G1.NumberDataTable(dt);
            dt.AcceptChanges();
            dgv.DataSource = dt;
            gridMain.ClearSelection();
            gridMain.RefreshData();
            gridMain.FocusedRowHandle = dt.Rows.Count - 1;
            gridMain.SelectRow(dt.Rows.Count - 1);
            dgv.Refresh();
            modified = true;
            if ( allowSave )
                btnSave.Show();
        }
        /***********************************************************************************************/
        private void gridMain_CalcRowHeight(object sender, RowHeightEventArgs e)
        {
            //GridView view = sender as GridView;
            //if (view == null)
            //    return;
            //if (e.RowHandle >= 0)
            //{
            //    int height = (int)view.GetDataRow(e.RowHandle)["RowHeight"].ObjToInt32();
            //    if (height < 1)
            //        height = 1;
            //    if (height >= 1)
            //    {
            //        height = 15 * height;
            //    }
            //    e.RowHeight = height;
            //}
        }
        /***********************************************************************************************/
        public XtraReport report = null;
        public XRSubreport subReport = null;
        public DevExpress.XtraReports.UI.XRPageBreak xrPageBreak2 = new XRPageBreak();
        public DevExpress.XtraReports.UI.GroupFooterBand xrGroupFooter = new GroupFooterBand();
        /***********************************************************************************************/
        public void CreateReport3(DataTable dt)
        {
            try
            {
                xrPageBreak2 = new XRPageBreak();
                //xrPageBreak2.BeforePrint += XrPageBreak2_BeforePrint;

                //xrGroupFooter.AfterPrint += XrGroupFooter_AfterPrint;
                xrGroupFooter.PageBreak = PageBreak.AfterBand;


                if (G1.get_column_number(dt, "packageprice") < 0)
                    dt.Columns.Add("packageprice", Type.GetType("System.Double"));

                XtraReport reportMerge = new XtraReport();
                reportMerge.Margins.Top = 50;
                reportMerge.Margins.Bottom = 50;
                reportMerge.Margins.Left = 30;
                reportMerge.Margins.Right = 75;

                DetailBand detailBand = new DetailBand();
                //detailBand.Height = labelDetail.Height;
                detailBand.Height = 0;
                detailBand.Name = "DetailBand";
                detailBand.KeepTogetherWithDetailReports = true;
                reportMerge.Bands.Add(detailBand);

                IList<XtraReport> reportList = new List<XtraReport>();

                //xrPageBreak1.Location = new System.Drawing.Point(0, 58);
                //xrPageBreak1.BeforePrint += XrPageBreak1_BeforePrint;

                Point point = new Point(0,0);
                xrPageBreak2.Location = point;


                //DevExpress.XtraReports.UI.XRPageBreak xrPageBreak1 = new XRPageBreak();
                //xrPageBreak1.Name = "xrPageBreak1";
                //xrPageBreak1.BeforePrint += XrPageBreak1_BeforePrint;

                int pageCount = 0;

                try
                {
                    int startRow = 0;
                    int myLastRow = 0;
                    bool includeBoth = false;
                    string lastBreak = "{HEADER}";
                    for (; ; )
                    {
                        report = new XtraReport();
                        report.Margins.Top = 50;
                        report.Margins.Bottom = 40;
                        report.Margins.Left = 30;
                        report.Margins.Right = 75;

                        detailBand = new DetailBand();
                        detailBand.Height = 0;
                        detailBand.KeepTogetherWithDetailReports = true;
                        report.Bands.Add(detailBand);

                        CreateDetailReport(report, dt, startRow, ref myLastRow, ref lastBreak, pageCount, ref includeBoth );

                        reportList.Add(report);
                        pageCount++;
                        //if (pageCount >= 3)
                        //    break;

                        point = new Point (0, point.Y + report.PageHeight );

                        if (String.IsNullOrWhiteSpace(lastBreak))
                            break;
                        startRow = myLastRow + 1;
                        //if (startRow >= 59)
                        //    break;
                    }
                }
                catch (Exception ex)
                {
                }
                int y = 0;
                int bottom = 0;
                for (int i = 0; i < reportList.Count; i++)
                {
                    subReport = new XRSubreport();
                    subReport.ReportSource = reportList[i];
                    subReport.LocationF = new Point(0, y);
                    subReport.GenerateOwnPages = true;
                    reportMerge.Bands["DetailBand"].Controls.Add(subReport);
                    bottom = reportList[i].PageHeight;
                    y += reportList[i].PageHeight;
                }

                PublishReport(reportMerge);
            }
            catch (Exception ex)
            {
            }
        }
        /***********************************************************************************************/
        private void XrGroupFooter_AfterPrint(object sender, EventArgs e)
        {
        }
        /***********************************************************************************************/
        private void XrPageBreak1_BeforePrint(object sender, PrintEventArgs e)
        {
            DevExpress.XtraReports.UI.XRPageBreak xrPageBreak1 = (XRPageBreak)sender;
            //Point point = xrPageBreak1.Location;
            xrPageBreak1.Visible = true;
        }
        /***********************************************************************************************/
        private void XrPageBreak2_BeforePrint(object sender, PrintEventArgs e)
        {
            //DevExpress.XtraReports.UI.XRPageBreak xrPageBreak1 = (XRPageBreak)sender;
            //Point point = xrPageBreak1.Location;
            //xrPageBreak2.Visible = true;
        }
        /***********************************************************************************************/
        private RichEditDocumentServer richEditDocumentServer;
        private string CreateDetailReport(XtraReport report, DataTable dt, int startRow, ref int myLastRow, ref string lastBreak, int pageCount, ref bool incBoth )
        {
            string saveLastBreak = lastBreak;

            bool gotPrices = false;
            if (G1.get_column_number(dt, "price") >= 0)
                gotPrices = true;

            string str = "";
            byte[] bytes = null;
            MemoryStream stream = null;
            DevExpress.XtraRichEdit.RichEditControl rtb = new RichEditControl();

            DevExpress.XtraReports.UI.PageHeaderBand PageHeader = new DevExpress.XtraReports.UI.PageHeaderBand();
            PageHeader.HeightF = 0F;

            bool underline = false;
            bool bold = false;
            float size = 9f;
            int indent = 0;
            int ipad = 0;
            string sValue = "";
            string desc = "";
            double pad = 0D;
            double dvalue = 0D;
            double bvalue = 0D;
            string holdingStr = "";
            double majorPad = 110;
            int holdCount = 0;
            int layin = 0;
            string saveDesc = "";
            string found = "";
            char c = (char)127;
            string baseString = " ";
            string tail = "";
            bool got5Percent = false;
            int extraWidth = 0;
            int tailIdx = 0;
            var sb = new StringBuilder(baseString);
            sb[0] = c;
            baseString = sb.ToString();
            string finale = "";

            DateTime date1 = this.dateTimePicker1.Value;

            int year = date1.Year;
            int day = date1.Day;
            string month = date1.ToString("MMMMMMMMMMMMM");

            string asOfDate = month + " " + day.ToString() + ", " + year.ToString("D4");

            bool includeBoth = incBoth;

            bool problem = false;


            DetailReportBand detailReportBand = new DetailReportBand();
            detailReportBand.HeightF = 0F;
            detailReportBand.WidthF = GetTotalPageWidth();
            report.Bands.Add(detailReportBand);

            ReportHeaderBand detailReportHeader = new ReportHeaderBand();
            detailReportHeader.HeightF = 0F;
            detailReportHeader.WidthF = GetTotalPageWidth();
            detailReportBand.Bands.Add(detailReportHeader);

            XRTable titleHeader = new XRTable();
            titleHeader.BeginInit();
            float titleY = 0;

            if ( saveLastBreak.Trim().ToUpper() == "{HEADER}")
                titleY = BuildFuneralHeading(titleHeader, saveLastBreak );

            titleHeader.EndInit();
            lastBreak = "";

            XRTable tableHeader = new XRTable();
            tableHeader.BeginInit();
            tableHeader.Rows.Add(new XRTableRow());
            tableHeader.Borders = BorderSide.All;
            tableHeader.BorderColor = Color.DarkGray;
            tableHeader.Font = new Font("Tahoma", 10, System.Drawing.FontStyle.Bold);
            tableHeader.Padding = 0;
            tableHeader.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft;
            tableHeader.HeightF = 0F;
            //tableHeader.LocationF = new PointF(tableHeader.LeftF, tableHeader.TopF + titleY);
            tableHeader.LocationF = new PointF(tableHeader.LeftF, titleHeader.BottomF );
            tableHeader.WidthF = GetTotalPageWidth();

            XRTableCell cellHeader1 = new XRTableCell();
            cellHeader1.Text = "Casket Name";
            cellHeader1.WidthF = GetTotalPageWidth() - 100;
            XRTableCell cellHeader2 = new XRTableCell();
            cellHeader2.Text = "Unit Price";
            cellHeader2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleCenter;
            cellHeader2.WidthF = 100F;

            tableHeader.Rows[0].Cells.AddRange(new XRTableCell[] { cellHeader1, cellHeader2 });
            detailReportHeader.Height = tableHeader.Height + titleHeader.Height;
            detailReportHeader.HeightF = 0F;

            PageHeader.Controls.Add(titleHeader);
            if ( saveLastBreak.Trim().ToUpper() == "{HEADER}")
                report.Bands.Add(PageHeader);

            tableHeader.EndInit();

            XRTable tableDetail = new XRTable();
            tableDetail.BeginInit();
            tableDetail.WidthF = GetTotalPageWidth();
            tableDetail.KeepTogether = true;
            tableDetail.Rows.Clear();

            string service = "";
            string price = "";
            string packagePrice = "";
            int lastRow = dt.Rows.Count;
            bool didSubHeader = false;
            //lastRow = 20;
            for (int i = startRow; i < lastRow; i++)
            {
                myLastRow = i;
                got5Percent = false;
                extraWidth = 0;
                saveDesc = "";
                underline = false;
                bold = false;
                size = 9f;
                indent = 0;
                layin = 0;
                problem = false;
                bytes = dt.Rows[i]["header"].ObjToBytes();
                size = dt.Rows[i]["size"].ObjToFloat();
                if (size <= 0f)
                    size = 9f;
                str = dt.Rows[i]["underline"].ObjToString();
                if (str.ToUpper().IndexOf("Y") >= 0)
                    underline = true;
                str = dt.Rows[i]["bold"].ObjToString();
                if (str.ToUpper().IndexOf("Y") >= 0)
                    bold = true;
                indent = dt.Rows[i]["indent"].ObjToInt32();
                if (indent <= 0)
                    indent = 0;
                layin = dt.Rows[i]["layin"].ObjToInt32();
                if (layin <= 0)
                    layin = 0;
                found = dt.Rows[i]["found"].ObjToString();

                service = dt.Rows[i]["service"].ObjToString();
                if (service == "Crescent Crowne Single Urn Vault")
                {
                }
                price = dt.Rows[i]["price"].ObjToString();
                packagePrice = dt.Rows[i]["packagePrice"].ObjToString();
                if (saveLastBreak.Trim().ToUpper() == "{SUB-HEADER}" && !didSubHeader )
                {
                    tableDetail.KeepTogether = true;
                    tableDetail.LocationF = new Point(0, 0);
                    BuildTableHeading(tableDetail, saveLastBreak);
                    didSubHeader = true;
                }
                if (bytes != null)
                {
                    str = G1.ConvertToString(bytes);

                    if (!String.IsNullOrWhiteSpace(str))
                    {
                        stream = new MemoryStream(bytes);
                        rtb.Document.Delete(rtb.Document.Range);
                        rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);
                        str = rtb.Text.Trim();
                        if (chkAsOfDate.Checked)
                        {
                            int idx = str.IndexOf("as of");
                            if (idx > 0)
                            {
                                int xdx = str.IndexOf(" but");
                                if (xdx > 0)
                                {
                                    string text = str.Substring(idx, xdx - idx);
                                    rtb.RtfText = rtb.RtfText.Replace(text, "as of " + asOfDate);
                                    //str = rtb.Text.Trim();
                                }
                            }
                        }
                        if (!String.IsNullOrWhiteSpace(str))
                        {
                            if (richEditDocumentServer == null)
                                richEditDocumentServer = new RichEditDocumentServer();

                            XRTableRow xrRowR = new XRTableRow();
                            xrRowR.WidthF = GetTotalPageWidth();
                            XRTableCell rCell = new XRTableCell();
                            XRRichText richtext = new XRRichText();
                            richtext.DataBindings.Clear();

                            richEditDocumentServer.RtfText = rtb.RtfText;
                            //ApplyRTFModification(richEditDocumentServer);
                            richtext.Rtf = richEditDocumentServer.RtfText;

                            //richtext.Text = rtb.Document.RtfText;
                            richtext.Location = new Point(0, 0);
                            richtext.CanGrow = true;
                            richtext.Size = rCell.Size;
                            richtext.WidthF = GetTotalPageWidth();
                            rCell.Controls.Add(richtext);
                            xrRowR.Cells.Add(rCell);
                            tableDetail.Rows.Add(xrRowR);
                        }
                    }
                }

                str = dt.Rows[i]["service"].ObjToString();
                if (str.IndexOf("Other Preparation Of The Body (Includes Cosmetology, Dressing And Casketing)") == 0)
                {
                }
                if (str.IndexOf("Canton Alternative Container - Cardboard - Laminate Venee") == 0)
                {
                }
                if (str == ".")
                    str = baseString;

                try
                {
                    if (!String.IsNullOrWhiteSpace(str))
                    {
                        if (str.Trim().ToUpper() == "{HEADER}")
                        {
                            lastBreak = str.Trim().ToUpper();
                            break;
                        }
                        else if (str.Trim().ToUpper() == "{SUB-HEADER}")
                        {
                            lastBreak = str.Trim().ToUpper();
                            break;
                        }
                        else if (str.Trim().ToUpper() == "{BREAK}")
                        {
                            lastBreak = str.Trim().ToUpper();
                            break;
                        }
                        else if (str.Trim().ToUpper() == "{EMPTY}")
                            str = baseString;
                        else if (str.Trim().ToUpper() == "{INCLUDE BOTH PRICES}")
                        {
                            includeBoth = true;
                            incBoth = true;
                        }
                        else if (str.Trim().ToUpper().IndexOf("{RANGE}") >= 0)
                        {
                            int idx = str.IndexOf('}');
                            str = str.Substring(idx + 1);
                            holdingStr = str.Trim();
                            pad = majorPad;
                            if (size != 9F)
                            {
                                pad = majorPad / size * 9D;
                                pad = pad - 1D;
                            }
                            pad = G1.RoundDown(pad);
                            pad = pad + CalculateIndent(indent, "Times New Roman", size);
                            ipad = Convert.ToInt32(pad);

                            holdingStr = holdingStr.PadRight(ipad);
                        }
                        else
                        {
                            dvalue = 3095D;
                            sValue = "$" + G1.ReformatMoney(dvalue);
                            if (gotPrices)
                            {
                                desc = dt.Rows[i]["price"].ObjToString();
                                if (!String.IsNullOrWhiteSpace(desc))
                                {
                                    dvalue = dt.Rows[i]["price"].ObjToDouble();
                                    sValue = "$" + G1.ReformatMoney(dvalue);
                                    if (includeBoth)
                                    {
                                        bvalue = dt.Rows[i]["packageprice"].ObjToDouble();
                                        if (bvalue > 9999.99D)
                                        {
                                            problem = true;
                                            sValue = " " + sValue;
                                            sValue += "                    $" + G1.ReformatMoney(bvalue);
                                        }
                                        else if (bvalue < 10D)
                                        {
                                            sValue += "                           $" + G1.ReformatMoney(bvalue);
                                        }
                                        else if (bvalue < 100D)
                                        {
                                            sValue += "                          $" + G1.ReformatMoney(bvalue);
                                        }
                                        else if (bvalue < 1000D)
                                        {
                                            sValue += "                        $" + G1.ReformatMoney(bvalue);
                                        }
                                        else
                                            sValue += "                     $" + G1.ReformatMoney(bvalue);
                                        if (dvalue == 0D && bvalue == 0D)
                                            sValue = "";
                                    }
                                }
                                else
                                    sValue = "";
                            }
                            else
                            {
                                if (includeBoth)
                                {
                                    //sValue += "     $" + G1.ReformatMoney(dvalue);
                                }
                            }
                            if (found != "Y")
                            {
                                if (sValue == "$0.00")
                                    sValue = baseString;
                                else if (sValue == "$0.00     $0.00")
                                    sValue = baseString;
                            }

                            desc = "";
                            pad = majorPad;
                            if (size != 9F)
                            {
                                pad = majorPad / size * 9D;
                                pad = pad - 1D;
                            }
                            pad = G1.RoundDown(pad);
                            //pad = Math.Truncate(pad);
                            pad = pad + indent;
                            ipad = Convert.ToInt32(pad);

                            desc = desc.PadRight(ipad);
                            if (!String.IsNullOrWhiteSpace(holdingStr))
                            {
                                if (holdCount == 0)
                                {
                                    holdingStr = holdingStr.TrimEnd();
                                    //holdingStr += "          " + sValue + "   to   ";
                                      holdingStr += "                                   " + sValue + "      to      ";
                                    holdCount++;
                                }
                                else
                                {
                                    XRTableRow xrRow = new XRTableRow();
                                    xrRow.WidthF = GetTotalPageWidth();
                                    if ( includeBoth )
                                    {
                                    }

                                    desc = "";
                                    holdingStr += sValue;
                                    //if (indent > 0)
                                    //{
                                    //    desc = " ".PadRight(indent*3);
                                    //}
                                    desc += holdingStr;
                                    if (desc.IndexOf("Basic Alternative Container - Cardboard") == 0)
                                    {
                                    }

                                    AddNormalLine(tableDetail, titleY, desc, bold, size, "Times New Roman", indent );
                                    //XRTableCell cell = new XRTableCell();
                                    //cell.Text = desc;
                                    //cell.WidthF = GetTotalPageWidth() - 100;
                                    //xrRow.Cells.Add(cell);

                                    //XRTableCell cell2 = new XRTableCell();
                                    ////cell.ExpressionBindings.Add(new ExpressionBinding("BeforePrint", "Text", "FormatString('{0:$0.00}', [UnitPrice])"));
                                    //cell2.Text = holdingStr;
                                    ////cell.Width = cellWidth[1];
                                    //cell2.WidthF = 100F;
                                    //cell2.TextAlignment = TextAlignment.MiddleRight;
                                    //xrRow.Cells.Add(cell2);
                                    //tableDetail.Rows.Add(xrRow);
                                    //AddNormalLine(rtb1, desc + "\n", bold, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                    holdCount = 0;
                                    holdingStr = "";
                                }
                            }
                            else
                            {
                                if (layin <= 0)
                                {
                                    tail = GetTail(dt, i);
                                    if (!String.IsNullOrWhiteSpace(tail))
                                        str = str.Replace(tail, "");
                                    else if (str.IndexOf("5% Of Total Assignment") > 0)
                                    {
                                        got5Percent = true;
                                        str = str.Replace("5% Of Total Assignment", "");
                                        sValue = "5% Of Total Assignment";
                                    }
                                    if ( sValue != baseString )
                                        desc = G1.lay_in_string(desc, str, indent, str.Length);
                                    else
                                    {
                                        desc = str;
                                    }
                                    int ll = desc.Length;
                                    //sValue += "~";
                                    //if (problem)
                                    //    indent = indent - 1;
                                    if (got5Percent)
                                    {
                                        desc += "  ";
                                        extraWidth = 110;
                                    }
                                    if (includeBoth)
                                        extraWidth = 160;
                                    if (sValue == baseString && extraWidth == 0)
                                        extraWidth = -70;
                                    //desc = G1.lay_in_string(desc, sValue, desc.Length - (sValue.Length + indent), sValue.Length);
                                    ll = desc.Length;
                                    if (desc.IndexOf("{") < 0 && desc.IndexOf("}") < 0)
                                    {
                                        if (desc.Trim() == baseString)
                                            AddNormalLine(tableDetail, titleY, "", "", bold, size, "Times New Roman", underline, indent );
                                        else
                                            AddNormalLine(tableDetail, titleY, desc, sValue, bold, size, "Times New Roman", underline, indent, extraWidth );
                                    }
                                }
                                else
                                    saveDesc = str;
                            }
                        }
                        bytes = dt.Rows[i]["tail"].ObjToBytes();
                        if (bytes != null)
                        {
                            str = G1.ConvertToString(bytes);

                            if (!String.IsNullOrWhiteSpace(str))
                            {
                                stream = new MemoryStream(bytes);
                                rtb.Document.Delete(rtb.Document.Range);
                                rtb.Document.LoadDocument(stream, DevExpress.XtraRichEdit.DocumentFormat.Rtf);

                                str = rtb.Text;
                                if (!String.IsNullOrWhiteSpace(str))
                                {
                                    if (layin > 0)
                                    {
                                        pad = majorPad;
                                        if (size != 9F)
                                        {
                                            pad = majorPad / size * 9D;
                                            pad = pad - 1D;
                                        }
                                        pad = G1.RoundDown(pad);
                                        pad = pad + indent;
                                        ipad = Convert.ToInt32(pad);


                                        XRTableRow xrRow = new XRTableRow();
                                        XRTableCell cell = new XRTableCell();
                                        if ( saveDesc == "Crescent Crowne Single Urn Vault")
                                        {
                                        }
                                        cell.Text = saveDesc;
                                        cell.WidthF = 108 + layin;
                                        cell.Font = new Font("Times New Roman", size, System.Drawing.FontStyle.Bold);
                                        xrRow.Cells.Add(cell);

                                        XRTableCell cell2 = new XRTableCell();
                                        cell2.Text = str;
                                        cell2.WidthF = 517F;
                                        cell2.TextAlignment = TextAlignment.MiddleLeft;
                                        cell2.Font = new Font("Times New Roman", size, System.Drawing.FontStyle.Regular);
                                        xrRow.Cells.Add(cell2);

                                        XRTableCell cell3 = new XRTableCell();
                                        cell3.Text = sValue;
                                        cell3.WidthF = 100F;
                                        cell3.TextAlignment = TextAlignment.MiddleRight;
                                        cell3.Font = new Font("Times New Roman", size, System.Drawing.FontStyle.Bold);
                                        xrRow.Cells.Add(cell3);
                                        tableDetail.Rows.Add(xrRow);



                                        //desc = G1.lay_in_string(saveDesc, str, layin, str.Length);
                                        desc = desc.PadRight(layin);
                                        //SetFont(rtb1, "Lucida Console", size, bold, false);
                                        saveDesc = saveDesc.PadRight(layin);
                                        //AddNormalLine(tableDetail, titleY, saveDesc, "", bold, size, "Times New Roman", underline, indent);
                                        //AddNormal(rtb1, saveDesc, bold, size, "Lucida Console", ParagraphAlignment.Left, underline);

                                        //rtb1.Document.AppendText(saveDesc);
                                        //SetFont(rtb1, "Lucida Console", size, false, false);
                                        //AddNormal(rtb1, str, false, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                        //rtb1.Document.AppendText(str);
                                        sValue = sValue.PadLeft(ipad - layin - str.Length);
                                        //                                SetFont(rtb1, "Lucida Console", size, bold, false);
                                        //AddNormalLine(tableDetail, titleY, saveDesc, sValue, bold, size, "Times New Roman", underline, indent);
                                        //AddNormal(rtb1, sValue + "\n", bold, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                        //rtb1.Document.AppendText("\n");
                                        //rtb1.Document.AppendText(sValue + "\n");
                                        //AddNormalText(rtb1, desc, bold, size);
                                        //str = str.PadRight(ipad - layin);
                                        //SetFont(rtb1, "Lucida Console", size, false, false);
                                        //AddNormalText(rtb1, str, false, size);
                                        //SetFont(rtb1, "Lucida Console", size, bold, false);
                                        //AddNormalText(rtb1, sValue + "\n", bold, size);
                                        //AddNormalLine(rtb1, desc + "\n", bold, size, "Lucida Console", ParagraphAlignment.Left, underline);
                                    }
                                    else
                                    {
                                        //AddParagraphMark(rtb);
                                        if (richEditDocumentServer == null)
                                            richEditDocumentServer = new RichEditDocumentServer();

                                        XRTableRow xrRowR = new XRTableRow();
                                        xrRowR.WidthF = GetTotalPageWidth();
                                        XRTableCell rCell = new XRTableCell();
                                        XRRichText richtext = new XRRichText();
                                        richtext.DataBindings.Clear();

                                        richEditDocumentServer.RtfText = rtb.RtfText;
                                        richtext.Rtf = richEditDocumentServer.RtfText;

                                        //richtext.Text = rtb.Document.RtfText;
                                        richtext.Location = new Point(0, 0);
                                        richtext.CanGrow = true;
                                        richtext.Size = rCell.Size;
                                        richtext.WidthF = GetTotalPageWidth();
                                        rCell.Controls.Add(richtext);
                                        xrRowR.Cells.Add(rCell);
                                        tableDetail.Rows.Add(xrRowR);

                                        //rtb1.Document.AppendRtfText(rtb.Document.RtfText);
                                    }
                                }
                                else if (str == "\r\n")
                                {
                                    //rtb1.Document.AppendText("\n");
                                }
                            }
                        }

                        //XRTableCell cell = new XRTableCell();
                        //cell.Text = price;
                        //cell.WidthF = GetTotalPageWidth() - 100;
                        //xrRow.Cells.Add(cell);

                        //XRTableCell cell2 = new XRTableCell();
                        ////cell.ExpressionBindings.Add(new ExpressionBinding("BeforePrint", "Text", "FormatString('{0:$0.00}', [UnitPrice])"));
                        //cell2.Text = price;
                        ////cell.Width = cellWidth[1];
                        //cell2.WidthF = 100F;
                        //cell2.TextAlignment = TextAlignment.MiddleRight;
                        //xrRow.Cells.Add(cell2);
                    }
                }
                catch (Exception ex)
                {
                }

                //tableDetail.Rows.Add(xrRow);
            }

//            tableDetail.Borders = BorderSide.Left | BorderSide.Right | BorderSide.Bottom;
            tableDetail.Borders = BorderSide.None;
            tableDetail.BorderColor = Color.DarkGray;
            tableDetail.Font = new Font("Tahoma", 10);
            tableDetail.Padding = 0;
            tableDetail.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft;
            tableDetail.WidthF = GetTotalPageWidth();

            DetailBand detailBand = new DetailBand();
            detailBand.Height = tableDetail.Height;
            detailBand.WidthF = 500F;
            detailReportBand.Bands.Add(detailBand);
            tableDetail.WidthF = GetTotalPageWidth();
            tableDetail.LocationF = new PointF(0, 0);
            detailBand.Controls.Add(tableDetail);
            //detailBand.PageBreak = PageBreak.AfterBand;

            //detailBand.Controls.Add(xrPageBreak2);

            tableDetail.EndInit();

            //if (lastBreak.Trim().ToUpper() == "{HEADER}" || lastBreak.Trim().ToUpper() == "{SUB-HEADER}")
            //    report.Bands.Add(PageHeader);

            return lastBreak;
        }
        /***********************************************************************************************/
        private float GetTotalPageWidth()
        {
            float totalWidth = report.PageWidth - report.Margins.Left - report.Margins.Right;
            return totalWidth;
        }
        /***********************************************************************************************/
        private void AdjustTableWidth(XRTable table)
        {
            XtraReport report = table.RootReport;
            table.WidthF = report.PageWidth - report.Margins.Left - report.Margins.Right - 10;
        }
        /***********************************************************************************************/
        private void PublishReport(XtraReport report)
        {
            ReportPrintTool printTool = new ReportPrintTool(report);
            if (workMassPrint)
            {
                string abrev = workDescription;
                string cmd = "Select * from `pricelists` where `description` = '" + workDescription + "';";
                DataTable dx = G1.get_db_data(cmd);
                if (dx.Rows.Count > 0)
                    abrev = dx.Rows[0]["abrev"].ObjToString();
                string asOfDate = this.dateTimePicker1.Value.ToString("yyyyMMdd");
                string filePath = "C:/SMFS_Reports/";
                G1.verify_path(filePath);
                filePath = "C:/SMFS_Reports/PriceLists/";
                G1.verify_path(filePath);
                filePath = "C:/SMFS_Reports/PriceLists/" + funeralHome + "_" + abrev + "_" + asOfDate + ".PDF";

                report.ExportToPdf(filePath);
            }
            else
                printTool.ShowPreviewDialog();
        }
        /***********************************************************************************************/
        private float BuildFuneralHeading( XRTable titleHeader, string lastBreak, bool includeSubHeading = true)
        {
            float titleY = 0F;
            if (lastBreak.Trim().ToUpper() == "{BREAK}")
                return titleY;
            titleY = AddHeaderLine(titleHeader, titleY, funeralHomeName, true, 18F, "Times New Roman");
            titleY = AddHeaderLine(titleHeader, titleY, funeralAddress, false, 14F, "Times New Roman");
            if (!String.IsNullOrWhiteSpace(funeralPOBox))
                titleY = AddHeaderLine(titleHeader, titleY, "P.O. Box " + funeralPOBox, false, 14F, "Times New Roman");
            string str = funeralCity + ", " + funeralState + "  " + funeralZip;
            titleY = AddHeaderLine(titleHeader, titleY, str, false, 14F, "Times New Roman");
            titleY = AddHeaderLine(titleHeader, titleY, funeralPhone, false, 14F, "Times New Roman");
            titleY = AddHeaderLine(titleHeader, titleY, " ", false, 14F, "Times New Roman");


            if (lastBreak.Trim().ToUpper() == "{HEADER}" )
            {
                str = workDescription;
                //if (workPrices.ToUpper() == "FUTURE")
                //    str += " (Future)";
                titleY = AddHeaderLine(titleHeader, titleY, str, true, 22F, "Times New Roman");
                titleY = AddHeaderLine(titleHeader, titleY, " ", false, 10F, "Times New Roman");
            }
            return titleY;
        }
        /***********************************************************************************************/
        private float BuildTableHeading (XRTable titleHeader, string lastBreak, bool includeSubHeading = true)
        {
            float titleY = 0F;
            if (lastBreak.Trim().ToUpper() == "{BREAK}")
                return titleY;
            //AddCenterLine(titleHeader, titleY, funeralHome, "", true, 18F, "Times New Roman");

            titleY = AddCenterLine(titleHeader, titleY, funeralHomeName, true, 18F, "Times New Roman");
            titleY = AddCenterLine(titleHeader, titleY, funeralAddress, false, 14F, "Times New Roman");
            if (!String.IsNullOrWhiteSpace(funeralPOBox))
                titleY = AddCenterLine(titleHeader, titleY, "P.O. Box " + funeralPOBox, false, 14F, "Times New Roman");
            string str = funeralCity + ", " + funeralState + "  " + funeralZip;
            titleY = AddCenterLine(titleHeader, titleY, str, false, 14F, "Times New Roman");
            titleY = AddCenterLine(titleHeader, titleY, funeralPhone, false, 14F, "Times New Roman");
            titleY = AddCenterLine(titleHeader, titleY, " ", false, 14F, "Times New Roman");


            if (lastBreak.Trim().ToUpper() == "{HEADER}")
            {
                str = workDescription;
                //if (workPrices.ToUpper() == "FUTURE")
                //    str += " (Future)";
                titleY = AddCenterLine(titleHeader, titleY, str, true, 22F, "Times New Roman");
                titleY = AddCenterLine(titleHeader, titleY, " ", false, 14F, "Times New Roman");
            }
            return titleY;
        }
        /***********************************************************************************************/
        private float AddHeaderLine (XRTable titleHeader, float titleY, string line, bool bold, float size, string fontname = "")
        {
            //titleHeader.BeginInit();
            titleHeader.Rows.Add(new XRTableRow());
            titleHeader.Borders = BorderSide.None;
            titleHeader.BorderColor = Color.Black;
            if (String.IsNullOrWhiteSpace(fontname))
                fontname = "Times New Roman";
            if ( bold )
                titleHeader.Font = new Font(fontname, size, System.Drawing.FontStyle.Bold);
            else
                titleHeader.Font = new Font(fontname, size, System.Drawing.FontStyle.Regular);
            titleHeader.Padding = 0;
            titleHeader.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleCenter;
            titleHeader.WidthF = GetTotalPageWidth();
            titleHeader.LocationF = new PointF(titleHeader.LeftF, titleHeader.TopF + titleY);

            XRTableCell titleCellHeader1 = new XRTableCell();
            titleCellHeader1.Text = line;
            titleCellHeader1.WidthF = GetTotalPageWidth();
            if (bold)
                titleCellHeader1.Font = new Font(fontname, size, System.Drawing.FontStyle.Bold);
            else
                titleCellHeader1.Font = new Font(fontname, size, System.Drawing.FontStyle.Regular);
            int row = titleHeader.Rows.Count;
            titleHeader.Rows[row-1].Cells.AddRange(new XRTableCell[] { titleCellHeader1 });
            //titleY += titleHeader.Rows[row - 1].HeightF;
            //titleY += titleHeader.HeightF;
            //titleHeader.EndInit();
            return titleY;
        }
        /***********************************************************************************************/
        private float AddNormalLine(XRTable tableDetail, float titleY, string line, bool bold, float size, string fontname = "", int indent = 0 )
        {
            if (line.IndexOf("Basic Alternative Container - Cardboard") == 0)
            {
            }

            //titleHeader.BeginInit();
            XRTableRow xrRow = new XRTableRow();
            xrRow.WidthF = GetTotalPageWidth();

            xrRow.Borders = BorderSide.None;
            xrRow.BorderColor = Color.Black;
            if (String.IsNullOrWhiteSpace(fontname))
                fontname = "Times New Roman";
            if (bold)
                xrRow.Font = new Font(fontname, size, System.Drawing.FontStyle.Bold);
            else
                xrRow.Font = new Font(fontname, size, System.Drawing.FontStyle.Regular);
            xrRow.Padding = 0;
            xrRow.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleLeft;
            xrRow.WidthF = GetTotalPageWidth();
            //titleHeader.LocationF = new PointF(titleHeader.LeftF, titleHeader.TopF + titleY);

            XRTableCell titleCellHeader1 = new XRTableCell();
            titleCellHeader1.Text = line;
            titleCellHeader1.WidthF = GetTotalPageWidth();
            if (bold)
                titleCellHeader1.Font = new Font(fontname, size, System.Drawing.FontStyle.Bold);
            else
                titleCellHeader1.Font = new Font(fontname, size, System.Drawing.FontStyle.Regular);
            //int row = titleHeader.Rows.Count;
            //titleHeader.Rows[row - 1].Cells.AddRange(new XRTableCell[] { titleCellHeader1 });

            XRTableCell cell = new XRTableCell();
            cell.Text = line;
            cell.WidthF = GetTotalPageWidth();
            cell.Padding = CalculateIndent(indent, fontname, size);

            xrRow.Cells.Add(cell);
            tableDetail.Rows.Add(xrRow);

            return titleY;
        }
        /***********************************************************************************************/
        private int CalculateIndent (int indent, string fontName, float size )
        {
            if (indent <= 0)
                return 0;
            string measureString = "Measure String";
            Font stringFont = new Font(fontName, size);

            // Measure string.
            SizeF stringSize = new SizeF();
            Graphics g = this.CreateGraphics();
            stringSize = g.MeasureString(measureString, stringFont);
            float len = measureString.Length;
            int newIndent = (int) (stringSize.Width / len);
            newIndent = newIndent * indent;
            if (newIndent < 0)
                newIndent = 0;
            g.Dispose();
            g = null;
            return newIndent;
        }
        /***********************************************************************************************/
        private float AddNormalLine(XRTable tableDetail, float titleY, string line, string sValue, bool bold, float size, string fontname = "", bool underline = false, int indent = 0, int extraWidth = 0 )
        {
            //titleHeader.BeginInit();
            XRTableRow xrRow = new XRTableRow();
            //xrRow.WidthF = tableDetail.WidthF;

            xrRow.Borders = BorderSide.None;
            xrRow.BorderColor = Color.Black;
            if (String.IsNullOrWhiteSpace(fontname))
                fontname = "Times New Roman";
            if ( String.IsNullOrWhiteSpace ( line ) && String.IsNullOrWhiteSpace ( sValue))
            {
                line = " ";
                sValue = " ";
            }

            XRTableCell cell = new XRTableCell();
            if (bold)
                cell.Font = new Font(fontname, size, System.Drawing.FontStyle.Bold);
            else
                cell.Font = new Font(fontname, size, System.Drawing.FontStyle.Regular);

            if (underline)
                cell.Borders = BorderSide.Bottom;
            cell.Text = line;
            cell.WordWrap = false;
            cell.WidthF = tableDetail.WidthF - 100F - extraWidth;
            cell.Padding = CalculateIndent(indent, fontname, size);
            xrRow.Cells.Add(cell);

            XRTableCell cell2 = new XRTableCell();
            if (bold)
                cell2.Font = new Font(fontname, size, System.Drawing.FontStyle.Bold);
            else
                cell2.Font = new Font(fontname, size, System.Drawing.FontStyle.Regular);
            if (underline)
                cell2.Borders = BorderSide.Bottom;
            string sValue1 = "";
            string sValue2 = "";
            bool isDouble = CheckForDoubleMoney(sValue, ref sValue1, ref sValue2);
            //isDouble = false;
            if (isDouble)
            {
                cell2.Text = sValue1;
                cell2.WordWrap = false;
                cell2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight;
                cell2.WidthF = 50F + extraWidth / 2;
                xrRow.Cells.Add(cell2);

                XRTableCell cell3 = new XRTableCell();
                if (bold)
                    cell3.Font = new Font(fontname, size, System.Drawing.FontStyle.Bold);
                else
                    cell3.Font = new Font(fontname, size, System.Drawing.FontStyle.Regular);
                if (underline)
                    cell3.Borders = BorderSide.Bottom;

                cell3.Text = sValue2;
                cell3.WordWrap = false;
                cell3.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight;
                cell3.WidthF = 80F + extraWidth / 2;
                xrRow.Cells.Add(cell3);
            }
            else
            {
                cell2.Text = sValue;
                cell2.WordWrap = false;
                cell2.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleRight;
                cell2.WidthF = 100F + extraWidth;
                xrRow.Cells.Add(cell2);
            }

            tableDetail.Rows.Add(xrRow);
            return titleY;
        }
        /***********************************************************************************************/
        private bool CheckForDoubleMoney(string sValue, ref string sValue1, ref string sValue2 )
        {
            bool isDouble = false;
            if (String.IsNullOrWhiteSpace(sValue))
                return false;
            sValue1 = "";
            sValue2 = "";
            G1.parse_answer_data(sValue, " ", true );
            int count = G1.of_ans_count;
            if (count == 2)
            {
                sValue1 = G1.of_answer[0].Trim();
                sValue2 = G1.of_answer[1].Trim();
                isDouble = true;
            }
            return isDouble;
        }
        /***********************************************************************************************/
        private float AddCenterLine(XRTable tableDetail, float titleY, string line, bool bold, float size, string fontname = "" )
        {
            //titleHeader.BeginInit();
            XRTableRow xrRow = new XRTableRow();

            xrRow.Borders = BorderSide.None;
            xrRow.BorderColor = Color.Black;
            if (String.IsNullOrWhiteSpace(fontname))
                fontname = "Times New Roman";

            XRTableCell cell = new XRTableCell();
            if (bold)
                cell.Font = new Font(fontname, size, System.Drawing.FontStyle.Bold);
            else
                cell.Font = new Font(fontname, size, System.Drawing.FontStyle.Regular);

            cell.Text = line;
            cell.WordWrap = false;
            cell.TextAlignment = DevExpress.XtraPrinting.TextAlignment.MiddleCenter;

            cell.WidthF = tableDetail.WidthF;
            xrRow.Cells.Add(cell);

            tableDetail.Rows.Add(xrRow);
            return titleY;
        }
        /***********************************************************************************************/
        public static void replaceCredits(DataTable dt, string PackageName)
        {
            //if (1 == 1)
            //    return;

            string service = "";

            if (G1.get_column_number(dt, "DATA") < 0)
                dt.Columns.Add("data");

            DataTable dx = dt.Clone();
            DataRow[] dRows = null;
            int lastRow = 0;
            string deleted = "";
            if (G1.get_column_number(dt, "DELETED") < 0)
                return;

            bool gotUrn = false;
            string newUrn = "";
            double urnPrice = 0D;

            bool gotAlter = false;
            string newAlter = "";
            double AlterPrice = 0D;

            bool honorReplacement = true;

            dRows = dt.Select("ModMod='Y'");

            if (dRows.Length <= 3)
                honorReplacement = false;

            if (String.IsNullOrWhiteSpace(PackageName))
                return;

            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    service = dt.Rows[i]["service"].ObjToString();
                    if ( service.ToUpper() == "URN CREDIT" )
                    {
                        dt.Rows[i]["service"] = "The Alumina - Aluminum";
                    }
                    else if (service.ToUpper() == "TEMPORARY URN CREDIT")
                    {
                        dt.Rows[i]["service"] = "Temporary Urn";
                    }
                    else if (service.ToUpper() == "ALTERNATIVE CONTAINER CREDIT" && deleted != "D")
                    {
                        dt.Rows[i]["service"] = "Basic Alternative Container - Cardboard";
                    }
                    else if (service.ToUpper() == "CREMATION CASKET CREDIT OR RENTAL CASKET WITH REMOVABLE INSERT" && deleted != "D")
                    {
                        dt.Rows[i]["service"] = "Standard Rental Casket";
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }
/***********************************************************************************************/
    }
}