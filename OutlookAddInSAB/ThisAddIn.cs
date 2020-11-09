using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using AddInsLibrary;
using System.Globalization;
using System.Threading;

namespace OutlookAddInSAB
{
    public partial class ThisAddIn
    {
        #region <定義>

        /// <summary>
        /// 関連会社ファイルと役職者ファイルの読み込みクラス
        /// </summary>
        ClsClassificationList data = new ClsClassificationList();

        /// <summary>
        /// 役職リストの読み込み結果格納
        /// </summary>
        List<ClsClassificationList.Manager> managerList;

        /// <summary>
        /// ファイルを入力値としたハッシュをキーにしてパスを格納
        /// </summary>
        Dictionary<int, string> dicPath = new Dictionary<int, string>();

        /// <summary>
        /// 添付ファイルを格納するリスト
        /// </summary>
        List<Outlook.Attachment> attachmentList = new List<Outlook.Attachment>();

        /// <summary>
        /// ファイルデータをリスト化
        /// </summary>
        public List<ClsFilePropertyList> m_fileList = new List<ClsFilePropertyList>();

        /// <summary>
        /// 圧縮クラス
        /// </summary>
        Zip zip = new Zip();

        /// <summary>
        /// 共通設定クラス
        /// </summary>
        ClsFilePropertyList filePropertyCls = new ClsFilePropertyList();

        /// <summary>
        /// 添付ファイルの一時フォルダ作成パス
        /// ドキュメント配下に格納
        /// </summary>
        string attachmentOutputPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);

        /// <summary>
        /// 共通設定ファイルクラス
        /// </summary>
        public CommonSettings clsCommonSettings;

        /// <summary>
        /// 役職情報判定クラス
        /// </summary>
        ClsConfidentialityMatrix clsPermitSetting = new ClsConfidentialityMatrix();

        #endregion

        #region <定数>

        /// <summary>
        /// 添付ファイルの一時フォルダ名
        /// ※ドキュメントフォルダ配下に作成される
        /// </summary>
        const string temporaryFolderName = "SAB_Attachments";

        /// <summary>
        /// 機密区分の設定値
        /// </summary>
        const string SECRECY_NONE_RANK = "0/";
        const string SECRECY_S_RANK = "1/";
        const string SECRECY_A_RANK = "2/";
        const string SECRECY_B_RANK = "3/";
        const string SECRECY_OTHER_RANK = "4/";

        /// <summary>
        /// 送信者役職区分の設定値
        /// </summary>
        const string EXECUTIVE = "executive/";
        const string MANAGER = "manager/";
        const string NOMAL = "nomal/";
        const string TRUE = "True";
        const string FALSE = "False";

        /// <summary>
        /// 区切り文字
        /// </summary>
        const string SEPARATOR = ";";

        /// <summary>
        /// メール送信時に添付ファイルか埋め込み画像かを判断する際に使用する定義
        /// </summary>
        const string PR_ATTACH_FLAGS = "http:" + "//schemas.microsoft.com/mapi/proptag/0x37140003";
        const string PR_ATTACH_CONTENT_ID = "http:" + "//schemas.microsoft.com/mapi/proptag/0x3712001E";

        #endregion

        #region <変数>

        /// <summary>
        /// 現在送信中のメールに添付されてる一時フォルダ名を保持
        /// </summary>
        string CurrentTemporaryFolderName = null;

        /// <summary>
        /// zip化処理で発行されたパスワード
        /// </summary>
        string m_zipPassword = "";

        #endregion        

        #region 起動時処理

        /// <summary>
        /// アドイン起動時処理
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (!StartupExecute())
            {
                // アドインが無効の場合
                return;
            }
        }

        /// <summary>
        /// アドイン起動
        /// </summary>
        /// <returns>true:正常、false:不正</returns>
        private bool StartupExecute()
        {
            // 共通設定クラスの作成
            if (!GetCommonSetting())
            {
                // 強制終了
                Environment.Exit(0x8020);
            }

            // 読み込んだ共通設定ファイルのチェック
            if (!CommonSettingCheck())
            {
                // 強制終了
                Environment.Exit(0x8020);
            }

            if (IsAddInEnable() == false)
            {
                // アドインが無効の場合
                return false;
            }

            // 役職判定
            clsPermitSetting.configZipLevel = clsCommonSettings.zipLevel;
            clsPermitSetting.Initialize();

            // 送信イベントの設定
            Application.ItemSend += Application_ItemSend1;

            // outlook終了イベントの設定
            ((Outlook.ApplicationEvents_11_Event)Application).Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Quit);

            managerList = data.ManagerList(clsCommonSettings);
            string strList = "";
            for (int i = 0; i < managerList.Count; i++)
            {
                strList += managerList[i].classification + "\r\n";
            }

            Outlook.ContactItem con = Application.Session.CurrentUser.AddressEntry.GetContact();

            // 言語設定読込み  // step2
            CultureInfo culture = CultureInfo.GetCultureInfo(clsCommonSettings.strCulture);
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;

            // 念のため、outlook起動時に一時フォルダを削除
            string TempPath = Path.Combine(attachmentOutputPath, temporaryFolderName);
            if (Directory.Exists(TempPath))
            {
                Directory.Delete(TempPath, true);
            }

            return true;
        }

        /// <summary>
        /// アドインの有効無効チェック
        /// </summary>
        private bool IsAddInEnable()
        {
            bool ret = false;

            // ManagerList.txtの存在チェック
            string ManagerFileServerPath = clsCommonSettings.strManagerListServerPath;
            string ManagerFileLocalPath = clsCommonSettings.strManagerListLocalPath;

            if (File.Exists(ManagerFileServerPath) == false && File.Exists(ManagerFileLocalPath) == false)
            {
                ThisRibbonCollection ribbonCollection =
                    Globals.Ribbons
                        [Globals.ThisAddIn.Application.ActiveInspector()];

                // リボン非表示
                ribbonCollection.Ribbon_Test.DisableRibonGroup();
            }
            else
            {
                ret = true;
            }

            return ret;
        }

        /// <summary>
        /// 共通設定ファイルを読み込んで、設定クラスを作成
        /// </summary>
        /// <param name=""></param>
        /// <returns>true:読み込み成功、false:読み込み失敗</returns>
        private bool GetCommonSetting()
        {
            try
            {
                CommonSettingRead read = new CommonSettingRead();
                clsCommonSettings = read.Reader();

                if (clsCommonSettings == null)
                {
                    MessageBox.Show(AddInsLibrary.Properties.Resources.msg_err_read_common_file,
                        AddInsLibrary.Properties.Resources.msgError,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Hand);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                // 共通設定ファイルを読み込めなかったのでoutlookを終了する
                return false;
            }
        }

        /// <summary>
        /// 読み込んだ共通設定ファイルのチェック
        /// </summary>
        /// <param name=""></param>
        /// <returns>true:正常、false:不正</returns>
        private bool CommonSettingCheck()
        {
            bool CommonSetttingFlg = true;          // 共通設定ファイルの不足項目有無判定フラグ
            string CommonSetttingMessage = null;    // 共通設定ファイルの不足項目メッセージ
            CommonSetttingMessage = AddInsLibrary.Properties.Resources.msgCommonSettingError + Environment.NewLine;
            
            // 共通設定ファイルに不足項目がないかチェック
            if (string.IsNullOrEmpty(clsCommonSettings.strDefaultSecrecyLevel))   // デフォルト機密区分
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += AddInsLibrary.Properties.Resources.msgDefaultSecure + Environment.NewLine;
            }
            if (string.IsNullOrEmpty(clsCommonSettings.strOfficeCode))   // 事業所コード
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += AddInsLibrary.Properties.Resources.msgOfficeCode + Environment.NewLine;
            }
            if (string.IsNullOrEmpty(clsCommonSettings.strCulture))   // 言語設定
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += AddInsLibrary.Properties.Resources.msgSettingLanguage + Environment.NewLine;
            }
            if (string.IsNullOrEmpty(clsCommonSettings.strSABListLocalPath))   // 文書のローカルパス
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += AddInsLibrary.Properties.Resources.msgLocalPath + Environment.NewLine;
            }
            if (string.IsNullOrEmpty(clsCommonSettings.strSABListServerPath))   // 文書のサーバーパス
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += AddInsLibrary.Properties.Resources.msgServerPath + Environment.NewLine;
            }
            if (string.IsNullOrEmpty(clsCommonSettings.strTempPath))   // zip一時解凍先
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += AddInsLibrary.Properties.Resources.msgTempZipPath + Environment.NewLine;
            }
            if (clsCommonSettings.lstSecureFolder.Count == 0)   // セキュアフォルダリスト
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += AddInsLibrary.Properties.Resources.msgFolderList + Environment.NewLine;
            }
            if (clsCommonSettings.lstFinal.Count == 0)   // 「最終版」を表す文字列
            {
                CommonSetttingFlg = false;
                CommonSetttingMessage += AddInsLibrary.Properties.Resources.msgFinal + Environment.NewLine;
            }

            if (!CommonSetttingFlg)
            {
                // 共通設定ファイルに不足項目あり
                MessageBox.Show(CommonSetttingMessage,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);

                // 強制終了
                return false;
            }

            return true;
        }

        #endregion

        #region 送信処理

        /// <summary>
        /// 送信ボタン押下
        /// </summary>
        /// <param name="Item">メール情報</param>
        /// <param name="Cancel">送信可否判定フラグ</param>
        private void Application_ItemSend1(object Item, ref bool Cancel)
        {
            // 送信実行
            SendExecute(Item, ref Cancel);
        }

        /// <summary>
        /// 送信実行
        /// </summary>
        /// <param name="Item">メール情報</param>
        /// <param name="Cancel">送信可否判定フラグ</param>
        private void SendExecute(object Item, ref bool Cancel)
        {
            Outlook.MailItem mailItem = Item as Outlook.MailItem;
            Outlook.Attachments attchments = mailItem.Attachments;

            // 2020/09/16
            // 添付ファイルリストは送信ボタン押下したタイミングで必ず初期化
            m_fileList.Clear();

            // 添付ファイルを一時フォルダに格納
            if (attchments.Count > 0)
            {
                if (!AttachmentTemporaryFolderStorage(attchments))
                {
                    // メッセージの二重表示を阻止
                    //MessageBox.Show(AddInsLibrary.Properties.Resources.msgTempCreateError,
                    //    AddInsLibrary.Properties.Resources.msgError,
                    //    MessageBoxButtons.OK,
                    //    MessageBoxIcon.Hand);

                    Cancel = true;
                    return;
                }
            }

            if (attchments.Count == 0)
            {
                // 添付ファイルがない場合は通常送信する
                return;
            }

            bool send = true;
            bool superiorPermit = false;
            bool zipPass = false;
            int maxSecrecyLevel = 0;
            int FileCount = 0;
            bool zipFolderFileExist = false;

            // 送信可否のチェック
            if (!SendJudge(mailItem, ref send, ref superiorPermit, ref zipPass, ref maxSecrecyLevel, ref FileCount, ref zipFolderFileExist))
            {
                // メッセージの二重表示を阻止
                //MessageBox.Show(AddInsLibrary.Properties.Resources.msgSendError,
                //    AddInsLibrary.Properties.Resources.msgError,
                //    MessageBoxButtons.OK,
                //    MessageBoxIcon.Hand);

                Cancel = true;
                return;
            }

            if (FileCount == 0)
            {
                // 添付ファイルがない場合は通常送信する
                return;
            }
            else
            {
                if (!zipFolderFileExist)
                {
                    // 何も格納されていないzipファイルが添付されているパターンの考慮を追加
                    return;
                }
            }

            Cancel = true;

            try
            {
                // 許可者要否のチェック
                if (!PermitJudge(mailItem, send, superiorPermit, maxSecrecyLevel, ref Cancel))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgPermitError, "" , MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }

            // メールを送信する場合
            if (Cancel == false)
            {

                if (zipPass == true)
                {
                    // zip化必要
                    // 添付ファイルのzip処理
                    if (!AttachmentZipCreate(mailItem, attchments))
                    {
                        MessageBox.Show(AddInsLibrary.Properties.Resources.msgTempZipError,
                            AddInsLibrary.Properties.Resources.msgError,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Hand);

                        Cancel = true;
                        return;
                    }

                    // パスワードメール下書き作成
                    if (!SetPasswordMail(mailItem, m_zipPassword))
                    {
                        Cancel = true;
                        return;
                    }
                }
                else
                {
                    // 圧縮確認ダイアログ
                    DialogResult drZipCheck = SendCheckZipMessageBox();
                    if (drZipCheck == DialogResult.Yes)
                    {
                        // YES
                        // 添付ファイルのzip処理
                        if (!AttachmentZipCreate(mailItem, attchments))
                        {
                            MessageBox.Show(AddInsLibrary.Properties.Resources.msgTempZipError,
                                AddInsLibrary.Properties.Resources.msgError,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Hand);

                            Cancel = true;
                            return;
                        }

                        // パスワードメール下書き作成
                        if (!SetPasswordMail(mailItem, m_zipPassword))
                        {
                            Cancel = true;
                            return;
                        }
                    }
                    else
                    {
                        // NO
                        // 圧縮なしでメール送信
                    }
                }
            }

            // 送信イベントの最後に一時フォルダを削除
            string TempPath = Path.Combine(attachmentOutputPath, temporaryFolderName);

            if (Directory.Exists(TempPath))
            {
                Directory.Delete(TempPath, true);

            }
        }

        #region 送信処理  要否・可否チェック処理

        /// <summary>
        /// 送信可否のチェック
        /// </summary>
        /// <param name="mailItem">メール情報</param>
        /// <param name="send">送信可否</param>
        /// <param name="superiorPermission">上長許可の要否</param>
        /// <param name="zipCompression">zip化要否</param>
        /// <param name="MaxSecrecyLevel">最高機密区分</param>
        /// <param name="FileCount">添付ファイル数</param>
        /// <param name="zipFolderFileExist">添付されているzipフォルダ内に保存されているファイルが存在するか</param>
        /// <returns>true:チェック正常、false:チェック不正</returns>
        private bool SendJudge(Outlook.MailItem mailItem, ref bool send, ref bool superiorPermission, ref bool zipCompression, ref int MaxSecrecyLevel, ref int FileCount, ref bool zipFolderFileExist)
        {
            try
            {
                // 送信者の役職、添付ファイルの最高機密区分、送信先(外部の有無)の三点から送信可否等が決まる
                string userClassification = ""; // 送信者役職区分
                int maxSecrecy = 10000; // 添付ファイル最高機密区分
                bool destinationCheck = true; // 外部への送信の有無

                ClsClassificationList classification = new ClsClassificationList();
                var associate = classification.AssociateList(clsCommonSettings);
                var managerial = classification.ManagerList(clsCommonSettings);

                // 機密区分の取得
                for (int i = 0; i < m_fileList.Count; i++)
                {
                    if (m_fileList[i].file_list == null || m_fileList[i].file_list.Count < 0)
                    {
                        // zipファイルではない場合
                        filePropertyCls.FileCheck(m_fileList[i]);

                        // zipフォルダ内が空のパターンチェック ⇒ zip以外のファイルの添付がある場合、チェック不要
                        zipFolderFileExist = true;
                    }
                    else
                    {
                        // zipファイルの場合

                        for (int j = 0; j < m_fileList[i].file_list.Count; j++)
                        {
                            HierarchyDelveInto(m_fileList[i].file_list, m_fileList[i].file_list.Count);

                            // zipフォルダ内が空のパターンチェック ⇒ zip内にファイルが１つでも存在する場合、チェック不要
                            zipFolderFileExist = true;
                        }
                    }
                }

                // 送信者の役職を取得する
                Outlook.AddressEntry entry = Application.Session.CurrentUser.AddressEntry;

                Outlook.ExchangeUser currentContactItem = GetCurrentContactItem();
                string userTitle = "";
                if (currentContactItem != null)
                {
                    userTitle = currentContactItem.JobTitle;
                }

#if DEBUG

                if (currentContactItem != null)
                {
                    MessageBox.Show("JobTitle:" + currentContactItem.JobTitle);  // 役職
                }
                else
                {
                    MessageBox.Show("Outlook.ContactItem currentContactItemがNULL");
                }
#endif

                // 送信者の役職区分を取得
                int index = managerial.FindIndex(x => x.manager == userTitle);

                if (index >= 0)
                {
                    userClassification = managerial[index].classification;
                }
                else
                {
                    userClassification = NOMAL.Substring(0, NOMAL.Length - 1);
                }

                List<string> ToAddress = new List<string>();
                List<string> CcAddress = new List<string>();
                List<string> BccAddress = new List<string>();

                // 送信者のアドインを取得
                string senderAddress = mailItem.To != null ? mailItem.To : "";
                string carbonCopy = mailItem.CC != null ? mailItem.CC : "";
                string blindCarbonCopy = mailItem.BCC != null ? mailItem.BCC : "";

                // 複数のアドレスを分割する
                //senderAddress = senderAddress.Replace(")", "");
                //carbonCopy = carbonCopy.Replace(")", "");
                //blindCarbonCopy = blindCarbonCopy.Replace(")", "");
                string[] ToAddressArray = senderAddress.Split(';');
                string[] CcAddressArray = carbonCopy.Split(';');
                string[] BccAddressArray = blindCarbonCopy.Split(';');

                ToAddressArray = GetAddress(ToAddressArray);
                CcAddressArray = GetAddress(CcAddressArray);
                BccAddressArray = GetAddress(BccAddressArray);

                // To, CC, BCCを一つの配列にまとめる
                string[] addressArray = new string[ToAddressArray.Length + CcAddressArray.Length + BccAddressArray.Length];
                ToAddressArray.CopyTo(addressArray, 0);
                CcAddressArray.CopyTo(addressArray, ToAddressArray.Length);
                BccAddressArray.CopyTo(addressArray, ToAddressArray.Length + CcAddressArray.Length);

                // 送信先に社外が含まれているかをチェック
                foreach (string address in addressArray)
                {
                    if (string.IsNullOrEmpty(address) == false)
                    {
                        if (!associate.Contains(address.Substring(address.IndexOf("@") + 1)))
                        {
                            destinationCheck = false;
                            break;
                        }
                    }
                }

                if (m_fileList.Count == 0)
                {
                    send = true;
                    maxSecrecy = 4;
                }
                else
                {
                    // 添付ファイルの最高機密区分の取得
                    for (int i = 0; i < m_fileList.Count; i++)
                    {
                        if (m_fileList[i].file_list == null || m_fileList[i].file_list.Count < 0)
                        {
                            if (m_fileList[i].fileSecrecyRank < maxSecrecy)
                            {
                                maxSecrecy = m_fileList[i].fileSecrecyRank;
                            }
                        }
                        else
                        {
                            for (int j = 0; j < m_fileList[i].file_list.Count; j++)
                            {
                                int fileSecrecy = GetfileSecrecy(m_fileList[i].file_list, m_fileList[i].file_list.Count, maxSecrecy);
                                if (fileSecrecy < maxSecrecy)
                                {
                                    maxSecrecy = fileSecrecy;
                                }
                            }
                        }
                    }
                }

                MaxSecrecyLevel = maxSecrecy;
                FileCount = m_fileList.Count;

                // 役職判定
                for (int i = 0; i < m_fileList.Count; i++)
                {
                    if (m_fileList[i].file_list == null || m_fileList[i].file_list.Count < 0)
                    {
                        string Secrecy = string.Format("{0}/", m_fileList[i].fileSecrecyRank);
                        string User = string.Format("{0}/", userClassification);
                        string Destination = string.Format("{0}", destinationCheck.ToString());

                        SendPattern sendPattern = clsPermitSetting.ResultMatrix[Secrecy][User][Destination];

                        // 判定用にセットする
                        if (sendPattern.bSend == false)
                        {
                            send = sendPattern.bSend;
                        }

                        if (sendPattern.bSuperiorPermission == true)
                        {
                            superiorPermission = sendPattern.bSuperiorPermission;
                        }

                        if (sendPattern.bZipCompression == true)
                        {
                            zipCompression = sendPattern.bZipCompression;
                        }
                    }
                    else
                    {
                        SendPattern CurrentPattern = new SendPattern(send, superiorPermission, zipCompression);

                        SendPattern sendPattern = GetSendPermit(m_fileList[i].file_list,
                            m_fileList[i].file_list.Count,
                            CurrentPattern,
                            userClassification,
                            destinationCheck.ToString()
                            );

                        // 解凍不可ZIPがあった場合
                        if (m_fileList[i].ZipError == true)
                        {
                            send = false;
                            return false;
                        }

                        // 判定用にセットする
                        if (sendPattern.bSend == false)
                        {
                            send = sendPattern.bSend;
                        }

                        if (sendPattern.bSuperiorPermission == true)
                        {
                            superiorPermission = sendPattern.bSuperiorPermission;
                        }

                        if (sendPattern.bZipCompression == true)
                        {
                            zipCompression = sendPattern.bZipCompression;
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgSendError,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);

                return false;
            }
        }

        /// <summary>
        /// 連絡帳からアドレスを取得
        /// 
        /// 表示名を使用し、アドレスを取得する
        /// </summary>
        /// <param name="addressArray">送信メールに設定されたアドレス</param>
        /// <returns>取得したアドレス</returns>
        private string[] GetAddress(string[] addressArray)
        {
            string[] emailAddress = new string[addressArray.Count()];
            int arrayCount = 0;
            string addressStr = null;
            bool addressFlg = false;

            //Outlook.ContactItem foundContact;
            Outlook.MAPIFolder contacts = (Outlook.MAPIFolder)Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

            foreach (string address in addressArray)
            {
                //foreach (var contact in contacts.Items)
                //{
                //    if (contact is Outlook.ContactItem)
                //    {
                //        foundContact = contact as Outlook.ContactItem;

                //        if (foundContact.Email1DisplayName == address.Trim())
                //        {
                //            emailAddress[arrayCount] = foundContact.Email1Address;
                //            arrayCount++;
                //            break;
                //        }
                //    }
                //}

                addressFlg = false;
                addressStr = null;

                if (Application.Session.Offline)
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
                                if (checkItem.Name == address.Trim())
                                {
                                    emailAddress[arrayCount] = checkItem.PrimarySmtpAddress;
                                    arrayCount++;
                                    addressFlg = true;
                                    break;
                                }
                            }
                        }

                        if (addressFlg == false)
                        {
                            // 手で入力したアドレスを格納
                            addressStr = address;

                            addressStr = addressStr.Replace("'", "");
                            addressStr = addressStr.Replace("\"", "");
                            emailAddress[arrayCount] = addressStr.Trim();
                            arrayCount++;
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
                            if (checkItem.Name == address.Trim())
                            {
                                emailAddress[arrayCount] = checkItem.PrimarySmtpAddress;
                                arrayCount++;
                                addressFlg = true;
                                break;
                            }
                        }
                    }

                    if (addressFlg == false)
                    {
                        // 手で入力したアドレスを格納
                        addressStr = address;

                        addressStr = addressStr.Replace("'", "");
                        addressStr = addressStr.Replace("\"", "");
                        emailAddress[arrayCount] = addressStr.Trim();
                        arrayCount++;
                    }
                }
            }

            return emailAddress;
        }

        /// <summary>
        /// 許可者要否のチェック
        /// </summary>
        /// <param name="mailItem">メール情報</param>
        /// <param name="send">送信可否</param>
        /// <param name="superiorPermission">上長許可の要否</param>
        /// <param name="MaxSecrecyLevel">最高機密区分</param>
        /// <param name="Cancel">送信可否判定フラグ</param>
        /// <returns>true:正常、false:不正</returns>
        private bool PermitJudge(Outlook.MailItem mailItem, bool send, bool superiorPermit, int maxSecrecyLevel, ref bool Cancel)
        {
            string SecrecyCheck = string.Format("{0}/", maxSecrecyLevel);
            if (send == true)
            {
                // 送信可能
                Cancel = false;
            }
            else
            {
                if (SecrecyCheck == SECRECY_NONE_RANK)
                {
                    // 機密区分登録なし送信メッセージ
                    SendNone_SecrecyMessageBox();
                    return false;
                }

                if (SecrecyCheck == SECRECY_S_RANK)
                {
                    // 機密区分登録なし送信メッセージ
                    SendS_SecrecyMessageBox();
                    return false;
                }
            }

            if (superiorPermit == true)
            {
                // 上長の許可が必要
                DialogResult dr1 = SendLicenserMessageBox1();
                if (dr1 == DialogResult.Yes)
                {
                    // はい

                    // CCの許可者確認
                    bool IsAuthorized = GetCCAuthorized(mailItem);

                    if (IsAuthorized == true)
                    {
                        // 許可者あり
                        DialogResult dr2 = SendLicenserMessageBox2();
                        if (dr2 == DialogResult.Yes)
                        {
                            // 送信
                        }
                        else if (dr2 == DialogResult.No)
                        {
                            // 許可者選択画面
                            SettingFormView(mailItem);

                            Cancel = true;
                        }
                        else
                        {
                            // キャンセル
                            Cancel = true;
                        }

                    }
                    else
                    {
                        // 許可者なし
                        DialogResult dr3 = SendLicenserMessageBox3();
                        if (dr3 == DialogResult.OK)
                        {
                            // 許可者選択画面
                            SettingFormView(mailItem);

                            Cancel = true;
                        }
                        else
                        {
                            // キャンセル
                            Cancel = true;
                        }
                    }
                }
                else
                {
                    // いいえ
                    Cancel = true;
                }
            }

            return true;
        }

        #endregion

        #region 送信処理  メッセージダイアログ

        /// <summary>
        /// 添付ファイル 機密区分登録なし送信
        /// </summary>
        private DialogResult SendNone_SecrecyMessageBox()
        {
            DialogResult ret = DialogResult.None;

            MessageBoxManager.Messagebox mb = new MessageBoxManager.Messagebox();
            ret = mb.Show_MessageBox(1);

            return ret;
        }

        /// <summary>
        /// 添付ファイル 送信不可-S秘送信
        /// </summary>
        private DialogResult SendS_SecrecyMessageBox()
        {
            DialogResult ret = DialogResult.None;

            MessageBoxManager.Messagebox mb = new MessageBoxManager.Messagebox();
            ret = mb.Show_MessageBox(2);

            return ret;
        }

        /// <summary>
        /// 添付ファイルに機密情報が含まれているメッセージダイアログを表示
        /// </summary>
        private DialogResult SendLicenserMessageBox1()
        {
            DialogResult ret = DialogResult.None;

            MessageBoxManager mbm = new MessageBoxManager();
            MessageBoxManager.InitMessageBoxManager();

            MessageBoxManager.Messagebox mb = new MessageBoxManager.Messagebox();
            ret = mb.Show_MessageBox(3);
            MessageBoxManager.ResetText();
            return ret;
        }

        /// <summary>
        /// 上長の有無判定メッセージダイアログを表示
        /// </summary>
        private DialogResult SendLicenserMessageBox2()
        {
            DialogResult ret = DialogResult.None;

            MessageBoxManager mbm = new MessageBoxManager();
            MessageBoxManager.InitMessageBoxManager();

            MessageBoxManager.Yes = AddInsLibrary.Properties.Resources.msgSend;
            MessageBoxManager.No = AddInsLibrary.Properties.Resources.msgPermitSelect;
            MessageBoxManager.Cancel = AddInsLibrary.Properties.Resources.msgCancel;
            MessageBoxManager.Register();

            MessageBoxManager.Messagebox mb = new MessageBoxManager.Messagebox();
            ret = mb.Show_MessageBox(5);
            MessageBoxManager.ResetText();

            return ret;
        }

        /// <summary>
        /// 送信許可権限者がCC欄に含まれていないメッセージダイアログを表示
        /// </summary>
        private DialogResult SendLicenserMessageBox3()
        {
            DialogResult ret = DialogResult.None;

            MessageBoxManager mbm = new MessageBoxManager();
            MessageBoxManager.InitMessageBoxManager();

            MessageBoxManager.OK = AddInsLibrary.Properties.Resources.msgPermitSelect;
            MessageBoxManager.Cancel = AddInsLibrary.Properties.Resources.msgCancel;
            MessageBoxManager.Register();

            MessageBoxManager.Messagebox mb = new MessageBoxManager.Messagebox();
            ret = mb.Show_MessageBox(4);
            MessageBoxManager.ResetText();

            return ret;
        }

        /// <summary>
        /// 圧縮の確認メッセージ表示
        /// </summary>
        private DialogResult SendCheckZipMessageBox()
        {
            DialogResult ret = DialogResult.None;

            MessageBoxManager.Messagebox mb = new MessageBoxManager.Messagebox();
            ret = mb.Show_MessageBox(6);

            return ret;
        }

        #endregion

        #region 送信処理  zip

        /// <summary>
        /// zipファイル
        /// ドキュメントが保存されている階層まで掘り下げ処理
        /// </summary>
        /// <param name="fileList">文書分類</param>
        /// <param name="count">ループする回数</param>
        private void HierarchyDelveInto(List<ClsFilePropertyList> fileList, int count)
        {
            // 参照中のフォルダに格納されているファイル数分ループ
            for (int i = 0; i < count; i++)
            {
                // ドキュメントが保存されている階層かチェック ⇒ ドキュメント階層の場合、必ずリストが存在しない
                if (fileList[i].file_list == null || fileList[i].file_list.Count < 0)
                {
                    // ドキュメントの場合、機密区分を取得
                    filePropertyCls.FileCheck(fileList[i]);
                }
                else
                {
                    // リストが存在する ⇒ 次の階層を見に行くので、この関数を再度呼ぶ
                    HierarchyDelveInto(fileList[i].file_list, fileList[i].file_list.Count);
                }
            }
        }

        /// <summary>
        /// zipファイル
        /// ドキュメントが保存されている階層まで掘り下げて、機密区分を取得
        /// </summary>
        /// <param name="fileList">文書分類</param>
        /// <param name="count">ループする回数</param>
        /// <param name="maxSecrecy">最高機密区分</param>
        private int GetfileSecrecy(List<ClsFilePropertyList> fileList, int count, int maxSecrecy)
        {
            // 参照中のフォルダに格納されているファイル数分ループ
            for (int i = 0; i < count; i++)
            {
                // ドキュメントが保存されている階層かチェック ⇒ ドキュメント階層の場合、必ずリストが存在しない
                if (fileList[i].file_list == null || fileList[i].file_list.Count < 0)
                {
                    if (fileList[i].fileSecrecyRank < maxSecrecy)
                    {
                        maxSecrecy = fileList[i].fileSecrecyRank;
                    }
                }
                else
                {
                    // リストが存在する ⇒ 次の階層を見に行くので、この関数を再度呼ぶ
                    maxSecrecy = GetfileSecrecy(fileList[i].file_list, fileList[i].file_list.Count, maxSecrecy);
                }
            }

            return maxSecrecy;
        }

        /// <summary>
        /// zipファイル
        /// ドキュメントが保存されている階層まで掘り下げて、送信判定を行う
        /// </summary>
        /// <param name="fileList">文書分類</param>
        /// <param name="count">ループする回数</param>
        /// <param name="CurrentSendPattern">送信判定</param>
        /// <param name="userClassification">自分の役職</param>
        /// <param name="destinationCheck">社内外</param>
        private SendPattern GetSendPermit(List<ClsFilePropertyList> fileList,
            int count,
            SendPattern CurrentSendPattern,
            string userClassification,
            string destinationCheck
            )
        {
            SendPattern ret = CurrentSendPattern;

            // 参照中のフォルダに格納されているファイル数分ループ
            for (int i = 0; i < count; i++)
            {
                // ドキュメントが保存されている階層かチェック ⇒ ドキュメント階層の場合、必ずリストが存在しない
                if (fileList[i].file_list == null || fileList[i].file_list.Count < 0)
                {
                    string Secrecy = string.Format("{0}/", fileList[i].fileSecrecyRank);
                    string User = string.Format("{0}/", userClassification);
                    string Destination = string.Format("{0}", destinationCheck.ToString());

                    SendPattern sendPattern = clsPermitSetting.ResultMatrix[Secrecy][User][Destination];

                    if (sendPattern.bSend == false)
                    {
                        ret.bSend = sendPattern.bSend;
                    }

                    if (sendPattern.bSuperiorPermission == true)
                    {
                        ret.bSuperiorPermission = sendPattern.bSuperiorPermission;
                    }

                    if (sendPattern.bZipCompression == true)
                    {
                        ret.bZipCompression = sendPattern.bZipCompression;
                    }
                }
                else
                {
                    // リストが存在する ⇒ 次の階層を見に行くので、この関数を再度呼ぶ
                    ret = GetSendPermit(fileList[i].file_list,
                        fileList[i].file_list.Count,
                        ret,
                        userClassification,
                        destinationCheck
                        );

                }
            }

            return ret;
        }

        /// <summary>
        /// 添付ファイルのzip化処理
        /// </summary>
        /// <param name=mailItem"">メール情報</param>
        /// <param name=attchments"">添付ファイル情報</param>
        /// <returns>true：zip化正常、false：zip化異常</returns>
        private bool AttachmentZipCreate(Outlook.MailItem mailItem, Outlook.Attachments attchments)
        {
            string[] fileList;

            // 一時フォルダと添付ファイル格納フォルダの存在チェック
            if (Directory.Exists(attachmentOutputPath + "\\" + temporaryFolderName) &&
                Directory.Exists(attachmentOutputPath + "\\" + temporaryFolderName + "\\" + CurrentTemporaryFolderName))
            {
                // 存在する場合、格納ファイルをパス形式で取得
                fileList = Directory.GetFiles(attachmentOutputPath + "\\" + temporaryFolderName + "\\" + CurrentTemporaryFolderName, "*", SearchOption.AllDirectories);
            }
            else
            {
                // 存在しない場合、エラー
                return false;
            }

            // 添付したファイル以外が存在していないかチェック
            bool checkFlg = false;
            for (int i = 0; i < fileList.Count(); i++)
            {
                // 解凍されたものかチェック
                if (fileList[i].Contains(zip.UNZIP_FOLDER))
                {
                    // outlookで添付したファイル
                    checkFlg = true;
                }

                if (checkFlg)
                {
                    // outlookに添付していない場合は削除
                    FileInfo fileInfo = new FileInfo(fileList[i]);
                    fileInfo.Delete();
                }
                checkFlg = false;
            }

            // 格納パスを再取得
            fileList = Directory.GetFiles(attachmentOutputPath + "\\" + temporaryFolderName + "\\" + CurrentTemporaryFolderName, "*", SearchOption.AllDirectories);

            // zip化実行
            try
            {
                m_zipPassword = zip.ZipCompression(fileList, attchments[1].FileName);

                // zipファイル添付前に添付したファイルを全て削除
                int attchmentCount = attchments.Count;
                for (int i = 0; i < attchmentCount; i++)
                {
                    mailItem.Attachments.Remove(1);
                }

                // zip化したファイルを添付
                mailItem.Attachments.Add(Zip.zipPath);
                File.Delete(Zip.zipPath);
            }
            catch (Exception ex)
            {
                // zip化で失敗した場合、エラーメッセージ表示
                return false;
            }

            return true;
        }

        #endregion

        #region 送信処理  一時ファイル処理

        /// <summary>
        /// 添付ファイルの一時フォルダ格納処理
        /// </summary>
        /// <param name="attchment">添付ファイル情報</param>
        /// <returns>true：正常、false：異常</returns>
        private bool AttachmentTemporaryFolderStorage(Outlook.Attachments attchments)
        {
            List<string> fileNameList = new List<string>();

            try
            {
                // 一時フォルダの存在チェック
                if (!Directory.Exists(attachmentOutputPath + "\\" + temporaryFolderName))
                {
                    // 存在しない場合、フォルダを作成
                    Directory.CreateDirectory(attachmentOutputPath + "\\" + temporaryFolderName);
                }

                // 該当ファイルを格納するパスを作成
                DateTime dt = DateTime.Now;
                string filePath = attachmentOutputPath + "\\" + temporaryFolderName + "\\" + dt.ToString("yyyyMMddHHmmssfff");

                // メール毎にフォルダ保存するので、タイムスタンプで格納フォルダを作成
                Directory.CreateDirectory(filePath);

                // 現在送信中メールに添付されているファイルが格納されているフォルダ名を保持
                CurrentTemporaryFolderName = dt.ToString("yyyyMMddHHmmssfff");

                int count = 1;
                for (int i = 1; i <= attchments.Count; i++)
                {
                    // 添付ファイルか埋め込み画像かチェック
                    bool IsAttachEmbedded = false;
                    int iAttFlags = attchments[i].PropertyAccessor.GetProperty(PR_ATTACH_FLAGS);
                    string strAttCID = attchments[i].PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID);
                    if (iAttFlags != 0)
                    {
                        IsAttachEmbedded = true;
                    }
                    if (string.IsNullOrEmpty(strAttCID) == false)
                    {
                        IsAttachEmbedded = true;
                    }
                    if (IsAttachEmbedded)
                    {
                        // フラグが0以外、ContentIDがあれば埋め込み画像
                        continue;
                    }

                    // 同じファイル名が添付されているかチェック
                    string fileName = "";
                    if (fileNameList.Contains(attchments[i].FileName))
                    {
                        // 同じファイル名の場合

                        int fileNameCount = fileNameList.Where(x => x == attchments[i].FileName).Count();

                        // ファイル名の後ろに数字を追記
                        fileName = Path.GetFileNameWithoutExtension(attchments[i].FileName) + "(" + fileNameCount.ToString() + ")" + Path.GetExtension(attchments[i].FileName); ;

                        count++;
                    }
                    else
                    {
                        // 違うファイル名の場合、そのままで保存
                        fileName = attchments[i].FileName;
                    }

                    // 添付ファイルを保存
                    attchments[i].SaveAsFile(filePath + "\\" + fileName);

                    // 同ファイル用にファイル名を格納
                    fileNameList.Add(attchments[i].FileName);

                    // ファイルデータをリスト化
                    CreateAttachmentDataList(attchments[i], filePath, fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgTempCreateError,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);

                return false;
            }

            return true;
        }

        /// <summary>
        /// ファイルデータをリスト化
        /// ※添付ファイル追加時のイベント処理を移植したもの
        /// </summary>
        /// <param name="attchment">メール情報</param>
        /// <param name="filePath">添付ファイルの格納先パス</param>
        /// <param name="fileName">添付ファイル名</param>
        private void CreateAttachmentDataList(Outlook.Attachment attchment, string filePath, string fileName)
        {
            bool zipError = false;

            // ファイルの拡張子を取得する
            string extension = Path.GetExtension(fileName);

            string filePathCombine = Path.Combine(filePath, fileName);
            if (extension != ".zip")
            {
                m_fileList.Add(new ClsFilePropertyList { attachment = attchment, fileName = fileName, filePath = filePathCombine, fileExtension = extension });
            }
            else
            {
                List<ClsFilePropertyList> unzipList = zip.UnZip(filePathCombine, filePath, m_fileList, ref zipError);
                m_fileList.Add(new ClsFilePropertyList { attachment = attchment, fileName = fileName, filePath = filePathCombine, fileExtension = extension, file_list = unzipList, ZipError = zipError });
            }
        }

        #endregion

        #region 送信処理  メソッド

        /// <summary>
        /// パスワードメール作成
        /// </summary>
        /// <param name="item">メール情報</param>
        /// <param name="_zipPassword">zip化処理で発行されたパスワード</param>
        /// <returns>true：メール作成成功、false：メール作成失敗</returns>
        private bool SetPasswordMail(Outlook.MailItem item, string _zipPassword)
        {
            try
            {
                string _Subject = AddInsLibrary.Properties.Resources.msgPasswordNotification;
                _Subject += item.Subject;

                string _BodyMessage = AddInsLibrary.Properties.Resources.msgPasswordMail;
                _BodyMessage = _BodyMessage.Replace("xxx", _zipPassword);

                var delayMail = (Outlook._MailItem)this.Application.CreateItem(Outlook.OlItemType.olMailItem);

                // デフォルトの署名を取得
                delayMail.Display(false);
                string _strSign = delayMail.Body;

                delayMail.To = item.To;
                delayMail.CC = item.CC;
                delayMail.BCC = item.BCC;
                delayMail.Subject = _Subject;
                delayMail.Body = _BodyMessage + Environment.NewLine + _strSign;
                delayMail.Save();
                delayMail.Display(false);

                // zipファイルを送信した事の通知
                SendZipMessageBox();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgPasswordMailError, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return false;
            }
        }

        /// <summary>
        /// zipファイルを送信した事の通知
        /// </summary>
        private DialogResult SendZipMessageBox()
        {
            DialogResult ret = DialogResult.None;

            MessageBoxManager.Messagebox mb = new MessageBoxManager.Messagebox();
            ret = mb.Show_MessageBox(7);

            return ret;
        }

        /// <summary>
        /// 許可者確認
        /// </summary>
        /// <param name="item">メール情報</param>
        private bool GetCCAuthorized(Outlook.MailItem item)
        {
            bool ret = false;

            if (string.IsNullOrEmpty(item.CC) == false)
            {
                SettingForm frm = new SettingForm();

                // 2020/10/06 GetCurrentContactItem()だと送信者の役職が未登録の場合の考慮がされていない ⇒ 一般扱いにならない
                //frm.CurrentContactItem = GetCurrentContactItem();
                frm.ExchangeUserItem = GetCurrentContactItem_2();
                frm.RefreshContactList();
                frm.SetSuperiorsList("");
                List<ListViewItem> SuperiorsList = frm.lstSuperiors;

                string SetCCItem = GetCCAddress(item);

                foreach (ListViewItem Superiors in SuperiorsList)
                {
                    string CheckSuperiorsUser = SEPARATOR + Superiors.Name + SEPARATOR;
                    if (SetCCItem.Contains(CheckSuperiorsUser) == true)
                    {
                        // 許可者あり
                        ret = true;
                    }
                }
            }

            return ret;
        }

        /// <summary>
        /// CCに入力されている値をアドレスだけ取得する
        /// </summary>
        /// <param name="item">メール情報</param>
        private string GetCCAddress(Outlook.MailItem item)
        {
            // 区切り文字は";"にしてアドレスの前後に挿入
            // (完全一致の検索の為）
            string ret = SEPARATOR;

            foreach (Outlook.Recipient i in item.Recipients)
            {
                if (i.Type == 2)
                {
                    if ((string.IsNullOrEmpty(i.Address) == false) &&
                       (string.IsNullOrEmpty(i.Name) == false))
                    {
                        ret += i.Address;
                        ret += SEPARATOR;
                    }
                    else if ((string.IsNullOrEmpty(i.Address) == true) &&
                            (string.IsNullOrEmpty(i.Name) == false))
                    {
                        // アドレスがnullの場合はnameから取得
                        ret += i.Name;
                        ret += SEPARATOR;
                    }
                }
            }

            return ret;
        }

        /// <summary>
        /// 許可者選択画面表示
        /// </summary>
        /// <param name="item">メール情報</param>
        public void SettingFormView(Outlook.MailItem item)
        {
            // 許可者選択
            SettingForm settingfrm = new SettingForm();

            // 2020/10/06 GetCurrentContactItem()だと送信者の役職が未登録の場合の考慮がされていない ⇒ 一般扱いにならない
            //settingfrm.CurrentContactItem = GetCurrentContactItem();
            settingfrm.ExchangeUserItem = GetCurrentContactItem_2();
            
            settingfrm.SetCCText = GetCCData(item);
            settingfrm.ShowDialog();

            if (settingfrm.DialogResult == DialogResult.OK)
            {
                string SetCCItem = GetCCData(item);

                // CCにセット
                foreach (string ResultCCItem in settingfrm.lstResultCCItem)
                {
                    if (SetCCItem.Contains(ResultCCItem) == false)
                    {
                        SetCCItem += ResultCCItem + SEPARATOR;
                    }
                }

                item.CC = SetCCItem;
            }
        }

        /// <summary>
        /// 自身のアドレス帳のオブジェクト取得
        /// </summary>
        private Outlook.ExchangeUser GetCurrentContactItem()
        {
            Outlook.ExchangeUser ret = null;
            Outlook.AddressEntry entry = Application.Session.CurrentUser.AddressEntry;

            Outlook.ExchangeUser user = entry.GetExchangeUser();

            if (string.IsNullOrEmpty(user.JobTitle) == false)
            {
                ret = user;
            }

            return ret;
        }

        /// <summary>
        /// 2020/10/06 追加
        /// 
        /// 自身のアドレス帳のオブジェクト取得
        /// 上記の「GetCurrentContactItem」だと役職が未登録の場合が考慮されていないので。。。
        /// </summary>
        private Outlook.ExchangeUser GetCurrentContactItem_2()
        {
            Outlook.ExchangeUser ret = null;
            Outlook.AddressEntry entry = Application.Session.CurrentUser.AddressEntry;

            ret = entry.GetExchangeUser();

            return ret;
        }

        /// <summary>
        /// CCに入力されている値を表示名通りに取得する
        /// </summary>
        /// <param name="item">メール情報</param>
        private string GetCCData(Outlook.MailItem item)
        {
            string ret = "";

            foreach (Outlook.Recipient i in item.Recipients)
            {
                if (i.Type == 2)
                {
                    if ((string.IsNullOrEmpty(i.Address) == false) &&
                       (string.IsNullOrEmpty(i.Name) == false))
                    {
                        if (i.Address.Length == i.Name.Length)
                        {
                            // 名称+アドレスの場合
                            ret += i.Name;
                            ret += SEPARATOR;
                        }
                        else
                        {
                            if (i.Name.Contains(i.Address) == false)
                            {
                                // Nameが名称のみだった場合
                                ret += i.Name;
                                ret += "<";
                                ret += i.Address;
                                ret += ">";
                                ret += (SEPARATOR + " ");
                            }
                            else
                            {
                                // アドレスのみの場合
                                ret += i.Name;
                                ret += SEPARATOR;
                            }
                        }
                    }
                    else if ((string.IsNullOrEmpty(i.Address) == true) &&
                            (string.IsNullOrEmpty(i.Name) == false))
                    {
                        // アドレスがnullの場合はnameから取得
                        ret += i.Name;
                        ret += SEPARATOR;
                    }
                }
            }

            return ret;
        }

        #endregion
        
        #endregion

        #region outlook終了処理

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //注: Outlook はこのイベントを発行しなくなりました。Outlook が
            //    シャットダウンする際に実行が必要なコードがある場合は、http://go.microsoft.com/fwlink/?LinkId=506785 を参照してください。
        }

        /// <summary>
        /// Outlook終了イベント
        /// </summary>
        /// <param name=""></param>
        private void ThisAddIn_Quit()
        {
            // 一時フォルダを削除
            string TampPath = Path.Combine(attachmentOutputPath, temporaryFolderName);
            if (Directory.Exists(TampPath))
            {
                Directory.Delete(TampPath, true);
            }
        }

        #endregion

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
