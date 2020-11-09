using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using AddInsLibrary;
using System.Diagnostics;

namespace OutlookAddInSAB
{
    public partial class SettingForm : Form
    {
        #region 定義

        /// <summary>
        /// CC
        /// </summary>
        public string SetCCText { get; set; }

        /// <summary>
        /// ListViewItemリスト
        /// </summary>
        List<ListViewItem> lviAllItem = new List<ListViewItem>();

        /// <summary>
        /// 自身のアドレス帳のオブジェクト取得
        /// 
        /// ※Exchangeに接続している場合
        /// </summary>
        public Outlook.ContactItem CurrentContactItem { get; set; }

        /// <summary>
        /// 自身のアドレス帳のオブジェクト取得
        /// 
        /// ※Exchangeに接続していない
        /// </summary>
        public Outlook.ExchangeUser ExchangeUserItem { get; set; }

        /// <summary>
        /// CCリスト
        /// </summary>
        public List<string> lstResultCCItem = new List<string>();

        /// <summary>
        /// ListViewItemリスト
        /// </summary>
        public List<ListViewItem> lstSuperiors = new List<ListViewItem>();

        /// <summary>
        /// メールの区切り文字
        /// </summary>
        private const char ADDRESS_SPACE = ';';

        /// <summary>
        /// 共通設定ファイルクラス
        /// </summary>
        CommonSettings clsCommonSettings;

        /// <summary>
        /// 共通設定ファイル 読み込む判定フラグ
        /// 
        /// true：読み込み成功、false：読み込み失敗
        /// </summary>
        public bool commonSettingFlg = true;

        #endregion

        public SettingForm()
        {
            InitializeComponent();
        }

        private void SettingForm_Load(object sender, EventArgs e)
        {
            System.Diagnostics.FileVersionInfo ver = FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string AssemblyName = ver.FileVersion;
            this.Text = AddInsLibrary.Properties.Resources.txt_Form_Title + " " + AssemblyName;

            tbxCCList.Text = SetCCText;

            RefreshContactList();

            if (!commonSettingFlg)
            {
                // 共通設定ファイル読み込み失敗の場合、画面を閉じる
                Close();
            }
        }

        #region ボタン押下処理

        /// <summary>
        /// キャンセルボタンクリック
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                SelectedCCItem();
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgCancelButtonError,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
            this.Close();
            this.Dispose();
        }

        /// <summary>
        /// CCボタンクリック
        /// </summary>
        private void btnCC_Click(object sender, EventArgs e)
        {
            try
            {
                SelectedCCItem();
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgCCButtonError,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// 検索ボタンクリック
        /// </summary>
        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(tbxSearchText.Text) != false)
                {
                    // 空欄の場合
                    RefreshContactList();
                    return;
                }

                // 検索処理
                SetSuperiorsList(tbxSearchText.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgSearchButtonError,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// OKボタンクリック
        /// </summary>
        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(tbxCCList.Text) != false)
                {
                    MessageBox.Show(AddInsLibrary.Properties.Resources.msgPermitSelection, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }

                string[] split = tbxCCList.Text.Split(ADDRESS_SPACE);
                foreach (string address in split)
                {
                    lstResultCCItem.Add(address.Trim());
                }

                this.DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgOKButtonError,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        #endregion

        #region メソッド

        /// <summary>
        /// 役職区分取得
        /// </summary>
        /// <param name="userTitle">役職</param>
        /// <returns>取得結果</returns>
        private int GetManagerIndex(string userTitle)
        {
            int ret = -1;

            // 共通設定ファイル読み込み
            CommonSettingRead read = new CommonSettingRead();
            clsCommonSettings = read.Reader();

            if (clsCommonSettings == null)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msg_err_read_common_file,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);

                // 共通設定ファイルの読み込み失敗
                commonSettingFlg = false;

                return ret;
            }

            ClsClassificationList classification = new ClsClassificationList();
            var managerial = classification.ManagerList(clsCommonSettings);

            // 送信者の役職区分を取得
            ret = managerial.FindIndex(x => x.manager == userTitle);

            return ret;
        }

        /// <summary>
        /// ListViewのアイテムをリフレッシュする
        /// </summary>
        public void RefreshContactList()
        {
            if (ExchangeUserItem == null)
            {
                return;
            }

            lstSuperiors.Clear();
            lviAllItem.Clear();

            // 接続状況の判定
            if (Globals.ThisAddIn.Application.Session.Offline)
            {
                // Exchangeに接続していない
                // 「Offline Global Address List」を取得

                MessageBox.Show("Application.Session.Offline = true  オフライン");

                foreach (Outlook.AddressList list in Globals.ThisAddIn.Application.Session.AddressLists)
                {
                    if (list.AddressListType.ToString() != "olExchangeGlobalAddressList")   // 該当のアドレス一覧名はべた書き。。。
                    {
                        continue;
                    }

                    foreach (Outlook.AddressEntry entryItem in list.AddressEntries)
                    {
                        Outlook.ExchangeUser checkItem = entryItem.GetExchangeUser();

                        if (checkItem != null)
                        {
                            if (string.IsNullOrEmpty(checkItem.CompanyName) == false)
                            {
                                if (checkItem.CompanyName.Contains(ExchangeUserItem.CompanyName) != false)
                                {
                                    string name = checkItem.LastName + " " + checkItem.FirstName;
                                    string department = checkItem.Department;
                                    string position = checkItem.JobTitle;

                                    ListViewItem lvi = new ListViewItem();
                                    lvi.ImageKey = checkItem.Name;
                                    lvi.Name = checkItem.PrimarySmtpAddress;
                                    lvi.Text = name;
                                    lvi.SubItems.Add(department);
                                    lvi.SubItems.Add(position);
                                    lviAllItem.Add(lvi);

                                    int Index = GetManagerIndex(position);
                                    if (Index != -1)
                                    {
                                        lstSuperiors.Add(lvi);
                                    }
                                    else
                                    {
                                        if (!commonSettingFlg)
                                        {
                                            // 共通設定ファイル読み込み失敗
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                // Exchangeに接続している
                // グローバルアドレス一覧を取得

                MessageBox.Show("Application.Session.Offline = false  オンライン");

                Outlook.AddressList list = Globals.ThisAddIn.Application.Session.GetGlobalAddressList();
                foreach (Outlook.AddressEntry entryItem in list.AddressEntries)
                {
                    Outlook.ExchangeUser checkItem = entryItem.GetExchangeUser();
                    if (checkItem != null)
                    {
                        if (string.IsNullOrEmpty(checkItem.CompanyName) == false)
                        {
                            if (checkItem.CompanyName.Contains(CurrentContactItem.CompanyName) != false)
                            {
                                string name = checkItem.LastName + " " + checkItem.FirstName;
                                string department = checkItem.Department;
                                string position = checkItem.JobTitle;

                                ListViewItem lvi = new ListViewItem();
                                lvi.ImageKey = checkItem.Name;
                                lvi.Name = checkItem.PrimarySmtpAddress;
                                lvi.Text = name;
                                lvi.SubItems.Add(department);
                                lvi.SubItems.Add(position);
                                lviAllItem.Add(lvi);

                                int Index = GetManagerIndex(position);
                                if (Index != -1)
                                {
                                    lstSuperiors.Add(lvi);
                                }
                                else
                                {
                                    if (!commonSettingFlg)
                                    {
                                        // 共通設定ファイル読み込み失敗
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //foreach (Outlook.AddressList list in Globals.ThisAddIn.Application.Session.AddressLists)
            //{
            //    foreach (Outlook.AddressEntry entryItem in list.AddressEntries)
            //    {
            //        Outlook.ContactItem checkItem = entryItem.GetContact();
            //        if (checkItem != null)
            //        {
            //            if (string.IsNullOrEmpty(checkItem.CompanyName) == false)
            //            {
            //                if (checkItem.CompanyName.Contains(CurrentContactItem.CompanyName) != false)
            //                {
            //                    string name = checkItem.FullName;
            //                    string department = checkItem.Department;
            //                    string position = checkItem.JobTitle;

            //                    ListViewItem lvi = new ListViewItem();
            //                    lvi.ImageKey = checkItem.Email1DisplayName;
            //                    lvi.Name = checkItem.Email1Address;
            //                    lvi.Text = name;
            //                    lvi.SubItems.Add(department);
            //                    lvi.SubItems.Add(position);
            //                    lviAllItem.Add(lvi);

            //                    int Index = GetManagerIndex(position);
            //                    if (Index != -1)
            //                    {
            //                        lstSuperiors.Add(lvi);
            //                    }
            //                    else
            //                    {
            //                        if (!commonSettingFlg)
            //                        {
            //                            // 共通設定ファイル読み込み失敗
            //                            return;
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}

            // 許可者リストにする
            SetSuperiorsList("");
        }

        /// <summary>
        /// 検索処理
        /// </summary>
        /// <param name="word">検索ワード</param>
        public void SetSuperiorsList(string word)
        {

            lviSuperiorsList.Items.Clear();
            //string[] CurrentDepartmenArray = CurrentContactItem.Department.Split(' ');

            // 部署を「○○○ ○○○ ・・・」形式にする
            string currentContact_departmenArray = departmentDisassembly(ExchangeUserItem.Department);
            
            // ListViewに並べる
            foreach (ListViewItem item in lviAllItem)
            {
                if (lviSuperiorsList.Items.ContainsKey(item.Name) == false)
                {
                    bool isSame = true;
                    if (string.IsNullOrEmpty(word) != false)
                    {
                        // 検索ワードがない場合
                        string departmen = item.SubItems[1].Text;

                        // 部署を「○○○ ○○○ ・・・」形式にする
                        string departmenStr = departmentDisassembly(departmen);

                        if (string.IsNullOrEmpty(departmenStr) == true)
                        {
                            continue;
                        }

                        string[] departmenArray = departmenStr.TrimEnd().Split(' ');

                        string CheckItem = "";
                        foreach (string dp in departmenArray)
                        {
                            if (string.IsNullOrEmpty(dp) == false)
                            {
                                CheckItem = CheckItem + " " + dp;
                                CheckItem = CheckItem.TrimStart(' ');

                                if (currentContact_departmenArray.Contains(CheckItem) == false)
                                {
                                    isSame = false;
                                }
                            }
                            else
                            {
                                isSame = false;
                            }
                        }
                    }

                    if (isSame != false)
                    {
                        string Position = item.SubItems[2].Text;
                        int Index = GetManagerIndex(Position);
                        if (Index != -1)
                        {
                            string SearchName = item.SubItems[0].Text;
                            if (SearchName.Contains(word) != false)
                            {
                                // 許可者のみを追加
                                lviSuperiorsList.Items.Add(item);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 行選択イベント
        /// </summary>
        private void SelectedCCItem()
        {
            if (lviSuperiorsList.SelectedItems.Count != 0)
            {
                if (tbxCCList.Text == "")
                {
                    string SelectItem = lviSuperiorsList.SelectedItems[0].Name;
                    if (tbxCCList.Text.Contains(SelectItem) == false)
                    {
                        tbxCCList.Text = lviSuperiorsList.SelectedItems[0].ImageKey;
                    }
                }
                else
                {
                    string SelectItem = lviSuperiorsList.SelectedItems[0].Name;
                    if (tbxCCList.Text.Contains(SelectItem) == false)
                    {
                        tbxCCList.Text = tbxCCList.Text.TrimEnd();
                        tbxCCList.Text = tbxCCList.Text.TrimEnd(';');
                        tbxCCList.Text += ADDRESS_SPACE + " " + lviSuperiorsList.SelectedItems[0].ImageKey;
                    }
                }
                lviSuperiorsList.SelectedItems.Clear();
            }
        }

        /// <summary>
        /// 一覧のダブルクリック処理
        /// </summary>
        private void lviSuperiorsList_DoubleClick(object sender, EventArgs e)
        {
            SelectedCCItem();
        }

        /// <summary>
        /// 部署を「○○○ ○○○ ・・・」形式にする
        /// <param name="departmen">部署</param>
        /// </summary>
        private string departmentDisassembly(string departmen)
        {
            departmen = departmen.Replace('　', ' ');
            string[] departmenArray = departmen.Split(' ');

            // スペースが連続であるかもしれないのでチェック
            string departmenCheck = null;
            foreach (string dp in departmenArray)
            {
                dp.Trim();
                if (string.IsNullOrEmpty(dp) == false)
                {
                    departmenCheck += dp + " ";
                }
            }

            return departmenCheck;
        }

        #endregion
    }
}
