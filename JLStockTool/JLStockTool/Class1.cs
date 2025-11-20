using CheckOne.CFTools;
using CheckOne.ClientLib;
using JLStockTool;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace CheckOne.FunPlugin
{
    public class FunPlugClass : FunPluginNonUIButton
    {
        //  private UniDBClient _dbClient;

        /// <summary>
        /// 建構式
        /// </summary>
        public FunPlugClass()
        {
            //Plugin描述
            this._DEPICTION.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this._DEPICTION.PluginFunction = "Test1載入視窗";
            this._DEPICTION.PluginDepiction = "顯示Test1功能的視窗。";

        }   //建構式
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="CtrlInfo"></param>
        /// <param name="param"></param>
        public override void Initialize(object sender, SCControlInfo CtrlInfo, object[] param)
        {
            try
            {
                base.Initialize(sender, CtrlInfo, param);
                //產生登打視窗
                if (_OBJ_PARAMETER == null)
                {
                    MessageBox.clsMessageBox.Show("未指定執行檔參數");
                    return;
                }
                CheckOneClient _coClient = (CheckOneClient)_OBJ_PARAMETER[0];
                StartForm frmMainShow = new StartForm((Control)sender, _OBJ_PARAMETER[1].ToString(), _coClient, this._ctrlInfo, null);

                this._SHOW_FORM = frmMainShow;


            }
            catch (Exception ex)
            {
                MessageBox.clsMessageBox.Show(string.Format("DataEntryForm.Initialize(){{{0}}}", ex.ToString()));
            }
        }   //初始化


        /// <summary>
        /// 複寫_funStart()放置被啟動後要處理的工作
        /// </summary>
        protected override void _funStart()
        {
            try
            {
                base._funStart();

#if (DEBUG)
                StringBuilder ctrlInfoString = new StringBuilder();
                ctrlInfoString.AppendFormat("Login User: {0} ({1})\n", _ctrlInfo.LoginUser.UserID, _ctrlInfo.LoginUser.UserName);
                ctrlInfoString.AppendFormat("User Dept: {0} ({1})\n", _ctrlInfo.LoginUser.DeptID, _ctrlInfo.LoginUser.DeptName);
                ctrlInfoString.AppendFormat("User Group: {0} ({1})\n", _ctrlInfo.LoginUser.GroupID, _ctrlInfo.LoginUser.GroupName);
                ctrlInfoString.AppendFormat("EntryLimit: {0} ({1})\n", _ctrlInfo.LoginUser.EntryLimit, String.Join("|", _ctrlInfo.LoginUser.EntryGroup));
                ctrlInfoString.AppendLine("ProList:");
                if (_ctrlInfo.ProList != null)
                {
                    foreach (SCMenuProItem item in _ctrlInfo.ProList)
                    {
                        ctrlInfoString.AppendFormat("\t{0} ({1})\n", item.ProID, item.ProName);
                    }
                }
                MessageBox.clsMessageBox.Show(ctrlInfoString.ToString(), "Debug Info");
#endif
            }
            catch (Exception ex)
            {
                MessageBox.clsMessageBox.Show(ex.ToString());
            }
        }
        /// <summary>
        /// 釋放資源
        /// </summary>
        public override void Dispose()
        {
            FireWorkStatusChanged("Dispose", true, DateTime.Now, "");
            base.Dispose();
        }   //釋放資源
    }
}

