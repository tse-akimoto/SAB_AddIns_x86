using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.IO;

namespace OutlookAddInSAB
{
    public partial class Ribbon_Test
    {
        #region 定義

        int count = 0;

        /// <summary>
        /// 一時フォルダのパス確認用
        /// </summary>
        string m_tempPath = "";

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
        ClsFilePropertyList clsFpl = new ClsFilePropertyList();

        /// <summary>
        /// コピーフォルダのパス
        /// </summary>
        string m_copyFolderPath;

        #endregion

        private void Ribbon_Test_Load(object sender, RibbonUIEventArgs e)
        {
            // メールオブジェクトの設定
            Outlook.MailItem _mailItem = null;

            Outlook.Inspector inspector = base.Context as Microsoft.Office.Interop.Outlook.Inspector;
            _mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (_mailItem == null)
            {
                // 送信フォームを閉じたときに検知する場合がある
                return;
            }

            // 言語設定読込み  // step2
            Localizable();
        }

        /// <summary>
        /// フォーム終了時の処理
        /// </summary>
        private void Ribbon_Test_Close(object sender, EventArgs e)
        {
            if (Directory.Exists(m_copyFolderPath))
            {
                Directory.Delete(m_copyFolderPath, true);
            }

            if (Directory.Exists(m_tempPath))
            {
                Directory.Delete(m_tempPath, true);
            }

            if (Directory.Exists(Zip.zipPath))
            {
                Directory.Delete(Zip.zipPath, true);
            }
        }

        #region ボタン押下処理

        /// <summary>
        /// フォーム表示ボタン押下処理
        /// </summary>
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                OutlookAddInSAB.SettingForm sf = new SettingForm();
                sf.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("フォーム表示ボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// メッセージボックス表示ボタン押下処理
        /// </summary>
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                count++;
                MessageBoxManager.Messagebox mb = new MessageBoxManager.Messagebox();
                mb.Show_MessageBox(count);
                if (count >= 5)
                    count = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("メッセージボックス表示ボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// メールアイテム取得ボタン押下処理
        /// </summary>
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                AttachmentFile af = new AttachmentFile();
                af.file_data();
            }
            catch (Exception ex)
            {
                MessageBox.Show("メールアイテム取得ボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// プロパティ取得ボタン押下処理
        /// </summary>
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string path = m_fileList[0].filePath;
                string directory = Path.GetDirectoryName(path);
                string fileName = Path.GetFileName(path);

                Shell32.Shell shell = new Shell32.Shell();
                Shell32.Folder folder = shell.NameSpace(directory);
                Shell32.FolderItem item = folder.ParseName(fileName);

                MessageBox.Show(folder.GetDetailsOf(item, 3));
            }
            catch (Exception ex)
            {
                MessageBox.Show("プロパティ取得ボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// 機密区分取得ボタン押下処理
        /// </summary>
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                GetProperty(m_fileList);
                MessageBox.Show(m_fileList[0].fileSecrecy + "\r\n" + m_fileList[0].fileClassification + "\r\n" + m_fileList[0].fileOfficeCode);
            }
            catch (Exception ex)
            {
                MessageBox.Show("機密区分取得ボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// zip作成ボタン押下処理
        /// </summary>
        private void btnZipCreate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.Inspector inspector = base.Context as Outlook.Inspector;
                Outlook.MailItem item = inspector.CurrentItem as Outlook.MailItem;
                Outlook.Attachments attachments = item.Attachments;
                if (attachments.Count > 0)
                {
                    item.Attachments.Add(Zip.zipPath); // zip化したファイルを添付
                    File.Delete(Zip.zipPath);

                    // zip化する前の元のファイルを添付ファイル一覧から削除
                    for (int i = 0; i < m_fileList.Count; i++)
                    {
                        File.Delete(m_fileList[i].filePath);
                        item.Attachments.Remove(1);
                    }
                    // リストからも削除する
                    int zipIndex = m_fileList.FindIndex(x => x.fileName.EndsWith(".zip"));
                    if (zipIndex != -1)
                    {
                        string zipDirectoryPath = m_fileList[zipIndex].filePath.Substring(0, m_fileList[zipIndex].filePath.Length - 4);
                        Directory.Delete(zipDirectoryPath, true);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("zip作成ボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// 添付ファイルリスト化ボタン押下処理
        /// </summary>
        private void btnTempList_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.Inspector inspector = base.Context as Outlook.Inspector;
                Outlook.MailItem item = inspector.CurrentItem as Outlook.MailItem;
                m_fileList.Clear();
                attachmentList.Clear();
                GetAttachments(item.Attachments);
            }
            catch (Exception ex)
            {
                MessageBox.Show("添付ファイルリスト化ボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// 送信先ドメイン取得ボタン押下処理
        /// </summary>
        private void btnGetAddress_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.Inspector inspector = base.Context as Outlook.Inspector;
                Outlook.MailItem item = inspector.CurrentItem as Outlook.MailItem;

                // toのアドレス取得
                // MailItem.ToでToに表示されているものを取得
                // MailItem.RecipientsでOutlook.Recipient型で取得、Outlook.Recipient.Addressでアドレスだけ取得
                // MailItem.Recipientはアドレス帳に存在する連絡先のみ？入力した順に入るためTo,CC,BCCの判別できない？
                string to = "";

                // MailItemの取得データから")"を消して@から後ろを取得すれば無理矢理ドメインのみ取得可能
                to = item.To.Replace(")", "");
                string[] toList = to.Split(';');

                foreach (Outlook.Recipient recp in item.Recipients)
                {
                    to += recp.Address + "\r\n";
                }
                MessageBox.Show(toList[0].Substring(toList[0].IndexOf("@") + 1));
            }
            catch (Exception ex)
            {
                MessageBox.Show("送信先ドメイン取得ボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// 送信チェックボタン押下処理
        /// </summary>
        private void btnSendCheck_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                bool send = false; // 送信可能か
                bool permission = false; // 上長許可が必要か
                bool compression = false; // zip圧縮が必要か
                SendJudge(ref send, ref permission, ref compression);
                string sendStr = (send == true) ? "送信可" : "送信不可";
                string permissionStr = (permission == true) ? "上長許可要" : "上長許可不要";
                string compressionStr = (compression == true) ? "圧縮要" : "圧縮不要";
                string msg = string.Format("{0}\r\n{1}\r\n{2}", sendStr, permissionStr, compressionStr);
                MessageBox.Show(msg, "送信可否");
            }
            catch (Exception ex)
            {
                MessageBox.Show("送信チェックボタン押下処理でエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// 送信者役職取得ボタン押下処理
        /// 送信者の役職を取得
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetUserJobTitle_Click(object sender, RibbonControlEventArgs e)
        {

        }

        /// <summary>
        /// 許可者選択ボタン押下処理
        /// </summary>
        private void btnSettingFormView_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 許可者選択画面表示
                Outlook.MailItem _mailItem = null;
                Outlook.Inspector inspector = base.Context as Microsoft.Office.Interop.Outlook.Inspector;
                _mailItem = inspector.CurrentItem as Outlook.MailItem;
                Globals.ThisAddIn.SettingFormView(_mailItem);
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgPermitSelectButtonError,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        #endregion

        #region メソッド

        /// <summary>
        /// カスタムリボンローカライズ
        /// </summary>
        private void Localizable()
        {
            OutlookAddInSAB.SettingForm frmSet = new OutlookAddInSAB.SettingForm();
            System.Globalization.CultureInfo culture = System.Globalization.CultureInfo.GetCultureInfo(Globals.ThisAddIn.clsCommonSettings.strCulture);

            System.Threading.Thread.CurrentThread.CurrentUICulture = culture;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon_Test));

            foreach (var tab in this.Tabs)
            {
                System.Diagnostics.Debug.WriteLine(tab.Name);
                resources.ApplyResources(tab, tab.Name, culture);
                foreach (var grp in tab.Groups)
                {
                    System.Diagnostics.Debug.WriteLine(grp.Name);
                    resources.ApplyResources(grp, grp.Name, culture);
                    foreach (var item in grp.Items)
                    {
                        System.Diagnostics.Debug.WriteLine(item.Name);
                        resources.ApplyResources(item, item.Name, culture);
                    }
                }
            }
        }

        /// <summary>
        /// 添付ファイルの取得
        /// </summary>
        /// <param name="item">添付されたファイル</param>
        /// <returns>true:成功、false:失敗</returns>
        private bool GetAttachments(Outlook.Attachments item)
        {
            bool zipError = false;
            if (m_tempPath == "")
            {
                Outlook.Attachment attachment = item[1];
                Zip.zipFilePath = "";
                string path = attachment.GetTemporaryFilePath();
                string[] arr = path.Split('\\');
                for (int i = 0; i < arr.Length - 1; i++)
                {
                    m_tempPath += arr[i] + "\\";
                    if (i < arr.Length - 2)
                    {
                        Zip.zipFilePath += arr[i] + "\\";
                    }
                }
            }
            try
            {
                foreach (var AttachmentObj in item)
                {
                    Outlook.Attachment Attachment = (Outlook.Attachment)AttachmentObj;
                    string path = Attachment.PathName;
                    string attachmentFileName = Attachment.FileName;
                    path = Path.Combine(m_tempPath, attachmentFileName);
                    string extension = Path.GetExtension(path);
                    string[] tempfolderList = Directory.GetFiles(m_tempPath);

                    if (!attachmentList.Contains(Attachment))
                    {
                        attachmentList.Add(Attachment);

                        if (extension != ".zip")
                        {
                            File.Copy(path, Path.Combine(m_copyFolderPath, attachmentFileName), true);
                            m_fileList.Add(new ClsFilePropertyList { attachment = Attachment, fileName = attachmentFileName, filePath = Path.Combine(m_tempPath, attachmentFileName), fileExtension = extension });
                        }
                        // zipの場合は解凍して中身のファイルを確認する
                        else
                        {
                            m_fileList.Add(new ClsFilePropertyList { attachment = Attachment, fileName = attachmentFileName, filePath = path, fileExtension = extension, file_list = zip.UnZip(path, m_tempPath, m_fileList, ref zipError) });
                            string sourceFilePath = Path.Combine(m_tempPath, attachmentFileName.Substring(0, attachmentFileName.Length - 4));
                            string copyDirectoryPath = Path.Combine(m_copyFolderPath, attachmentFileName.Substring(0, attachmentFileName.Length - 4));

                            File.Copy(path, Path.Combine(m_copyFolderPath, attachmentFileName), true);

                            int count = m_fileList.Count();
                        }
                    }
                }
                // 正常に取得できているか確認用表示部分
                string attachmentFile = "";
                for (int i = 0; i < m_fileList.Count; i++)
                {
                    attachmentFile += m_fileList[i].fileName + "\r\n";
                }
                MessageBox.Show(attachmentFile, "添付ファイル名");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            return zipError;
        }


        /// <summary>
        /// 一時フォルダをコピーするメソッド
        /// </summary>
        /// <param name="sourceFilePath">コピーするフォルダパス</param>
        /// <param name="copyDirectoryPath">コピー先のパス</param>
        private void CopyDirectory(string sourceFilePath, string copyDirectoryPath)
        {
            // 属性情報もコピーする
            File.SetAttributes(copyDirectoryPath, File.GetAttributes(sourceFilePath));

            // ディレクトリ名の末尾に"\"を付ける
            if (copyDirectoryPath[copyDirectoryPath.Length - 1] != Path.DirectorySeparatorChar)
            {
                copyDirectoryPath = copyDirectoryPath + Path.DirectorySeparatorChar;
            }

            // ファイルをコピーする
            string[] sourceFile = Directory.GetFiles(sourceFilePath);
            string[] directory = Directory.GetDirectories(sourceFilePath);
            foreach (string dir in directory)
            {
                CopyDirectory(dir, copyDirectoryPath + Path.GetFileName(dir));
            }
        }

        /// <summary>
        /// 送信可否のチェックを行う
        /// </summary>
        /// <param name="send">送信可否</param>
        /// <param name="superiorPermission">上長許可要否</param>
        /// <param name="zipCompression">zip圧縮要否</param>
        private void SendJudge(ref bool send, ref bool superiorPermission, ref bool zipCompression)
        {

        }

        /// <summary>
        /// 機密区分などのファイルプロパティを取得する
        /// </summary>
        /// <param name="fileList">対象のファイルリスト</param>
        private void GetProperty(List<ClsFilePropertyList> fileList)
        {
            for (int i = 0; i < fileList.Count; i++)
            {
                if (fileList[i].fileExtension != ".zip")
                {
                    clsFpl.FileCheck(fileList[i]);
                    MessageBox.Show(fileList[i].fileName + ":" + fileList[i].fileSecrecy, "機密区分");
                }
                else
                {
                    GetProperty(fileList[i].file_list);
                }
            }
        }

        /// <summary>
        /// 添付ファイルリストから最も高い機密区分を取得する
        /// </summary>
        /// <param name="fileList">添付ファイルリスト</param>
        /// <returns>最も高い機密区分ランク</returns>
        private int getMaxSecrecy(List<ClsFilePropertyList> fileList)
        {
            int maxSecrecy = 10000;
            for (int i = 0; i < fileList.Count; i++)
            {
                if (fileList[i].file_list != null)
                {
                    maxSecrecy = getMaxSecrecy(fileList[i].file_list);
                }
                else
                {
                    if (fileList[i].fileSecrecyRank < maxSecrecy)
                    {
                        maxSecrecy = fileList[i].fileSecrecyRank;
                    }
                }
            }
            return maxSecrecy;
        }

        /// <summary>
        /// 新規メール作成メソッド
        /// </summary>
        private void CreateNewMail()
        {
            var application = new Outlook.Application();
            Outlook.MailItem mailItem = application.CreateItem(Outlook.OlItemType.olMailItem);

            // 現在のメールを取得
            Outlook.Inspector ins = base.Context as Outlook.Inspector;
            Outlook.MailItem item = ins.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                // Toを現在のメールからコピー、何も入力されていない場合はスルー
                if (item.To != null)
                {
                    Outlook.Recipient to = mailItem.Recipients.Add(item.To);
                    to.Type = (int)Outlook.OlMailRecipientType.olTo;
                }
                // CCを現在のメールからコピー
                if (item.CC != null)
                {
                    Outlook.Recipient cc = mailItem.Recipients.Add(item.CC);
                    cc.Type = (int)Outlook.OlMailRecipientType.olCC;
                }
                // BCCを現在のメールからコピー
                if (item.BCC != null)
                {
                    Outlook.Recipient bcc = mailItem.Recipients.Add(item.BCC);
                    bcc.Type = (int)Outlook.OlMailRecipientType.olBCC;
                }

                // 件名を設定
                mailItem.Subject = "New Mail";
                // 本文を設定
                mailItem.Body = "Mail Body";
                // 新規メール画面をモードレスダイアログで開く
                mailItem.Display(false);
                mailItem.Save();
            }
        }

        /// <summary>
        ///  リボン非表示
        /// </summary>
        public void DisableRibonGroup()
        {
            tab1.Visible = false;
        }

        #endregion
    }
}
