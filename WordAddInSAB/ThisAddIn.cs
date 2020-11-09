using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;

namespace WordAddInSAB
{
    public partial class ThisAddIn
    {
        #region <定義>

        /// <summary>
        /// 保存先リスト
        /// </summary>
        private List<string> lstDocPath = new List<string>();

        /// <summary>
        /// 保存前日時
        /// </summary>
        DateTime AfterWriteTime = DateTime.Now;

        /// <summary>
        /// 保存後日時
        /// </summary>
        DateTime BeforeWriteTime = DateTime.Now;

        #endregion

        #region <イベント>

        /// <summary>
        /// アドイン起動時処理
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 言語設定読込み  // step2
            WordAddInSAB.SettingForm frmSet = new WordAddInSAB.SettingForm();
            CultureInfo culture = CultureInfo.GetCultureInfo(frmSet.clsCommonSettting.strCulture);
            Thread.CurrentThread.CurrentUICulture = culture;

            // 保存時のイベントを登録
            this.Application.DocumentBeforeSave +=
                new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);

            // 終了直前イベント
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
        }

        /// <summary>
        /// 終了時イベント
        /// </summary>
        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            string DocPath = "";

            if (string.IsNullOrEmpty(Doc.Path) == false)
            {
                // 自分自身のパス
                DocPath = Path.Combine(Doc.Path, Doc.Name);
            }

            if (lstDocPath.Contains(DocPath) == true)
            {
                if (IsEnableStorage(DocPath) == false)
                {
                    // 自分自身が正しい位置に保存されていない場合
                    Cancel = true;

                    MessageBox.Show(AddInsLibrary.Properties.Resources.msg_warning_save_not_secure, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);

                    ExecuteSaveAs();

                    return;
                }
                else
                {
                    // S秘A秘以外
                    lstDocPath.Remove(DocPath);
                }
            }

            // 保存されていない場合
            if (Doc.Saved == false)
            {
                if (IsEnableStorage(DocPath) == false)
                {
                    Cancel = true;

                    MessageBox.Show(AddInsLibrary.Properties.Resources.msg_warning_save_not_secure, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);

                    ExecuteSaveAs();

                    return;
                }
            }

            // 削除処理
            if (ClearListPath() == false)
            {
                Cancel = true;
            }
        }

        /// <summary>
        /// 有効な保存個所チェック
        /// </summary>
        private bool IsEnableStorage(string Path)
        {
            // プロパティ情報取得
            WordAddInSAB.SettingForm frmSet = new WordAddInSAB.SettingForm();

            try
            {
                string strFilePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
                string strFilePropertyClassNo = string.Empty;      // ファイルプロパティ情報 事業所コード
                string strFilePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

                // プロパティのタグを取得
                frmSet.GetDocumentProperty(ref strFilePropertySecrecyLevel, ref strFilePropertyClassNo, ref strFilePropertyOfficeCode); // プロパティ情報取得

                // プロパティにSAB情報は未設定の場合は設定画面を表示
                if (frmSet.IsSecrecyInfoRegistered() == false)
                {
                    // 必須登録モードON
                    frmSet.MustRegistMode = true;

                    frmSet.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    frmSet.ShowDialog();

                    frmSet.GetDocumentProperty(ref strFilePropertySecrecyLevel, ref strFilePropertyClassNo, ref strFilePropertyOfficeCode); // プロパティ情報取得​
                }

                // S秘・A秘なら保存場所の確認を行う
                if ((strFilePropertySecrecyLevel == AddInsLibrary.InfomationForm.SECRECY_PROPERTY_S)
                    || (strFilePropertySecrecyLevel == AddInsLibrary.InfomationForm.SECRECY_PROPERTY_A))
                {
                    List<string> lstTarGetSecureFolder = frmSet.clsCommonSettting.lstSecureFolder;
                    string result = lstTarGetSecureFolder.FirstOrDefault(x => Path.Contains(x));

                    if (result == null)
                    {
                        // 正しい場所に保存されてない場合
                        return false;
                    }
                }
            }
            catch
            {
                // 共通設定が読み込めない場合はそもそもセキュアチェックができない為チェックを行わないでスルーする
            }

            return true;
        }

        /// <summary>
        /// アドイン終了処理
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            ClearListPath();
        }

        #endregion

        #region <メソッド>

        /// <summary>
        /// 名前を付けて保存イベント実行
        /// </summary>
        private void ExecuteSaveAs()
        {
            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ThreadStart(() =>
            {
                // 非同期で保存イベントを呼ぶ
                Application.CommandBars.ExecuteMso("FileSaveAs");
            }));

            thread.Start();
        }

        /// <summary>
        /// 正しくない保存先のリストからファイルを削除する
        /// </summary>
        private bool ClearListPath()
        {
            // 正しい保存先の場合
            foreach (string file in lstDocPath)
            {
                bool isFileDelete = false;
                while (isFileDelete == false)
                {
                    try
                    {
                        if (File.Exists(file) != false)
                        {
                            File.Delete(file);
                        }

                        isFileDelete = true;
                    }
                    catch
                    {
                        // ファイルがロックされている
                        DialogResult dr = MessageBox.Show(AddInsLibrary.Properties.Resources.msg_err_lock_file, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation);

                        if (dr == System.Windows.Forms.DialogResult.Retry)
                        {
                            // 再試行
                        }
                        else
                        {
                            // キャンセル
                            return false;
                        }

                        System.Threading.Thread.Sleep(1000);
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// 保存後処理
        /// </summary>
        private void Application_DocumentAfterSave(Microsoft.Office.Interop.Word.Document Doc, bool isCanceled)
        {
            string DocPath = Doc.Path;
            string FileName = Doc.Name;
            string FullPath = Path.Combine(DocPath, FileName);

            if (string.IsNullOrEmpty(DocPath) != false)
            {
                return;
            }

            BeforeWriteTime = System.IO.File.GetLastWriteTime(FullPath);

            if (BeforeWriteTime == AfterWriteTime)
            {
                // キャンセル
                return;
            }

            AfterWriteTime = System.IO.File.GetLastWriteTime(FullPath);

            // Word画面が表示されていない場合は設定画面を表示しない
            if (this.Application.Visible == false)
            {
                return;
            }

            // プロパティ情報取得
            WordAddInSAB.SettingForm frmSet = new WordAddInSAB.SettingForm();

            // 共通設定エラー時処理
            if (frmSet.commonFileReadCompleted == false)
            {
                return;
            }

            string strFilePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
            string strFilePropertyClassNo = string.Empty;      // ファイルプロパティ情報 事業所コード
            string strFilePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

            // プロパティのタグを取得
            frmSet.GetDocumentProperty(ref strFilePropertySecrecyLevel, ref strFilePropertyClassNo, ref strFilePropertyOfficeCode); // プロパティ情報取得

            if (lstDocPath.Contains(FullPath))
            {
                lstDocPath.Remove(FullPath);
            }

            // S秘・A秘なら保存場所の確認を行う
            if ((strFilePropertySecrecyLevel == AddInsLibrary.InfomationForm.SECRECY_PROPERTY_S)
                || (strFilePropertySecrecyLevel == AddInsLibrary.InfomationForm.SECRECY_PROPERTY_A))
            {
                List<string> lstTarGetSecureFolder = frmSet.clsCommonSettting.lstSecureFolder;

                // セキュアフォルダと同一なら保存
                string result = lstTarGetSecureFolder.FirstOrDefault(x => DocPath.Contains(x));
                if (result == null)
                {
                    // セキュアフォルダではない
                    MessageBox.Show(AddInsLibrary.Properties.Resources.msg_warning_save_not_secure, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);

                    lstDocPath.Add(FullPath);

                    ExecuteSaveAs();

                    return;
                }
            }
            ClearListPath();
        }

        /// <summary>
        /// 保存時処理
        /// </summary>
        /// <param name="Wb">Word情報</param>
        /// <param name="SaveAsUI">保存フラグ</param>
        /// <param name="Cancel">キャンセルフラグ</param>
        void Application_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            bool IsBackgroundSave = false;
            while (IsBackgroundSave == false)
            {
                if (Doc.Application.BackgroundSavingStatus != 0)
                {
                    IsBackgroundSave = false;
                }
                else
                {
                    IsBackgroundSave = true;
                }

                System.Threading.Thread.Sleep(1000);
            }

            // Word画面が表示されていない場合は設定画面を表示しない
            if (this.Application.Visible == false)
            {
                return;
            }

            // プロパティ情報取得
            WordAddInSAB.SettingForm frmSet = new WordAddInSAB.SettingForm();

            // 共通設定エラー時処理
            if (frmSet.commonFileReadCompleted == false)
            {
                return;
            }

            string strFilePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
            string strFilePropertyClassNo = string.Empty;      // ファイルプロパティ情報 事業所コード
            string strFilePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

            // プロパティにSAB情報は未設定の場合は設定画面を表示
            if (frmSet.IsSecrecyInfoRegistered() == false)
            {
                // 必須登録モードON
                frmSet.MustRegistMode = true;

                frmSet.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                frmSet.ShowDialog();
            }
            else
            {
                // プロパティのタグを取得
                frmSet.GetDocumentProperty(ref strFilePropertySecrecyLevel, ref strFilePropertyClassNo, ref strFilePropertyOfficeCode); // プロパティ情報取得

                // ファイルの事業所コードと設定値の事業所コードを比較
                if (strFilePropertyOfficeCode == frmSet.clsCommonSettting.strOfficeCode)
                {
                    // プロパティに情報を書込み
                    frmSet.SetDocumentProperty(strFilePropertySecrecyLevel);
                }
                else
                {
                    // 修正を押下された場合は、設定画面を表示する
                    frmSet.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    frmSet.ShowDialog();
                }
            }

            new System.Threading.Thread(() =>
            {
                while (true)
                {
                    try
                    {
                        var application = Doc.Application;
                        while (application.BackgroundSavingStatus > 0)
                            System.Threading.Thread.Sleep(1000);
                        break;
                    }
                    catch
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                }

                Application_DocumentAfterSave(Doc, !Doc.Saved);
            }).Start();
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
