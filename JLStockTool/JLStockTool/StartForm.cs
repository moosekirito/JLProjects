using CheckOne.CFTools;
using CheckOne.ClientLib;
using JZRZ;
using System;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace JLStockTool
{
    public partial class StartForm : Form
    {
        private readonly CheckOneClient _coClient;
        private readonly SCControlInfo _ctrlInfo;
        private UniDBClient _dbClient;
        private readonly InfoFilter _log;

        private DataTable _excelData;

        public StartForm(object sender, string module, CheckOneClient client, SCControlInfo ctrl, object[] param)
        {
            InitializeComponent();

            _coClient = client;
            _ctrlInfo = ctrl;
            _log = new InfoFilter(module);

            LoadBasicConfig();
        }

        private void LoadBasicConfig()
        {
            try
            {
                // 讀取 "法扣資料處理連線"（你系統通用的）
                string connStr = _coClient.CreateRemotingDBConnectString("法扣資料處理連線");
                _dbClient = UniDBOper.GetDBConnectionObject(UniDBType.CheckOneRemoting, connStr);

                _log.GetInfo(LogWriteLevel.Info, "初始化完成", null);
            }
            catch (Exception ex)
            {
                _log.GetInfo(LogWriteLevel.Error, "初始化失敗", ex);
                MessageBox.clsMessageBox.Show("初始化失敗！");
            }
        }

        private void btnLoadExcel_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Filter = "Excel Files|*.xlsx;*.xls;*.csv";
                if (dlg.ShowDialog() != DialogResult.OK) return;

                string full = dlg.FileName;
                string folder = Path.GetDirectoryName(full);
                string file = Path.GetFileName(full);

                _excelData = LoadExcel(folder, file);
                dgv.DataSource = _excelData;

                cmbColumns.Items.Clear();
                foreach (DataColumn col in _excelData.Columns)
                    cmbColumns.Items.Add(col.ColumnName);

                _log.GetInfo(LogWriteLevel.Info, $"Excel 載入: {file}", null);
            }
            catch (Exception ex)
            {
                _log.GetInfo(LogWriteLevel.Error, "Excel 載入失敗", ex);
                MessageBox.clsMessageBox.Show("Excel 載入失敗！");
            }
        }

        private DataTable LoadExcel(string folder, string file)
        {
            string msg = "";
            DataSet ds = FileReaderTrans.FileReaderTrans.ReaderToDataSet(folder + "\\", file, out msg);
            if (!string.IsNullOrEmpty(msg))
                throw new Exception(msg);

            return ds.Tables[0];
        }

        private void btnMatch_Click(object sender, EventArgs e)
        {
            if (_excelData == null)
            {
                MessageBox.clsMessageBox.Show("尚未載入 Excel");
                return;
            }

            string col = cmbColumns.Text;
            if (string.IsNullOrEmpty(col))
            {
                MessageBox.clsMessageBox.Show("請選擇比對欄位！");
                return;
            }

            DataTable dt = _excelData;

            foreach (DataRow row in dt.Rows)
            {
                string stockCode = row[col]?.ToString() ?? "";

                string sql = "EXEC JL_Stock_Query @Code";
                var parms = new DbParameter[]
                {
                    new SqlParameter("@Code", stockCode)
                };

                DataTable result = _dbClient.GetDataTable(sql, parms);

                _log.GetInfo(LogWriteLevel.Info, $"查詢 {stockCode}", null);
            }

            MessageBox.clsMessageBox.Show("比對完成了！");//test

        }

        private void cmbColumns_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
