using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;

using System.Data;
using System.Windows.Forms;
using System.IO;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraPrinting;
using DevExpress.Utils;
using MySql.Data.MySqlClient;
using GeneralLib;

// Hello world. This is Adam.
// It says there are no unpushed commits, but I don't believe it.

namespace SMFS
{
    public partial class XtraReport1 : DevExpress.XtraReports.UI.XtraReport
    {
        public static DevExpress.XtraGrid.GridControl workdgv;
        public XtraReport1( DataTable dt, DevExpress.XtraGrid.GridControl dgv )
        {
            workdgv = dgv;
            workdgv.DataSource = dt;
            workdgv.DataSource = dt;
            workdgv.DataSource = dt;
            InitializeComponent();
            workdgv.DataSource = dt;
            this.objectDataSource1.DataSource = dt;
            this.DataSource = dt;
        }
    }
}
