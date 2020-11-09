using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using System.Threading;

namespace PowerPointAddInSAB
{
    public partial class ThisAddIn
    {

        #region <定義>
        /// <summary>
        /// Ribbonクラス
        /// </summary>
        private Ribbon _ribbon;

        /// <summary>
        /// 保存先リスト
        /// </summary>
        private List<string> lstPresPath = new List<string>();

        #endregion

        #region <イベント>

        /// <summary>
        /// アドイン起動時処理
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 言語設定読込み  // step2
            PowerPointAddInSAB.SettingForm frmSet = new PowerPointAddInSAB.SettingForm();
            CultureInfo culture = CultureInfo.GetCultureInfo(frmSet.clsCommonSettting.strCulture);
            Thread.CurrentThread.CurrentUICulture = culture;

            // 保存時のイベントを登録
            this.Application.PresentationBeforeSave +=
                new Microsoft.Office.Interop.PowerPoint.EApplication_PresentationBeforeSaveEventHandler(ApplicationPresentationBeforeSave);

            // 保存後イベント
            this.Application.PresentationSave += Application_PresentationSave;

            // 終了直前イベント
            this.Application.PresentationBeforeClose += Application_PresentationBeforeClose;

        }

        /// <summary>
        /// 終了時イベント
        /// </summary>
        private void Application_PresentationBeforeClose(PowerPoint.Presentation Pres, ref bool Cancel)
        {
            string PresPath = "";

            if (string.IsNullOrEmpty(Pres.Path) == false)
            {
                // 自分自身のパス
                PresPath = Path.Combine(Pres.Path, Pres.Name);
            }

            if (lstPresPath.Contains(PresPath) == true)
            {
                if (IsEnableStorage(PresPath, Pres) == false)
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
                    lstPresPath.Remove(PresPath);
                }
            }

            // 保存されていない場合
            if (Pres.Saved == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                if (IsEnableStorage(PresPath, Pres) == false)
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
        private bool IsEnableStorage(string Path, PowerPoint.Presentation Pres)
        {
            // プロパティ情報取得
            PowerPointAddInSAB.SettingForm frmSet = new PowerPointAddInSAB.SettingForm();
            frmSet.propPres = Pres;

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
            foreach (string file in lstPresPath)
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
        /// 保存後イベント
        /// </summary>
        private void Application_PresentationSave(PowerPoint.Presentation Pres)
        {
            if (Pres != this.Application.ActivePresentation)
            {
                return;
            }

            // PowerPoint画面が表示されていない場合は設定画面を表示しない
            if (this.Application.Visible == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                return;
            }

            // プロパティ情報取得
            PowerPointAddInSAB.SettingForm frmSet = new PowerPointAddInSAB.SettingForm();
            frmSet.propPres = Pres;

            // 共通設定エラー時処理
            if (frmSet.commonFileReadCompleted == false)
            {
                return;
            }

            // スライド数が0のときは登録不要
            if (frmSet.GetSlideCount() <= 0)
            {
                return;
            }

            string strFilePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
            string strFilePropertyClassNo = string.Empty;      // ファイルプロパティ情報 事業所コード
            string strFilePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

            // プロパティのタグを取得
            frmSet.GetDocumentProperty(ref strFilePropertySecrecyLevel, ref strFilePropertyClassNo, ref strFilePropertyOfficeCode); // プロパティ情報取得

            if (lstPresPath.Contains(Path.Combine(Pres.Path, Pres.Name)))
            {
                lstPresPath.Remove(Path.Combine(Pres.Path, Pres.Name));
            }

            // S秘・A秘なら保存場所の確認を行う
            if ((strFilePropertySecrecyLevel == AddInsLibrary.InfomationForm.SECRECY_PROPERTY_S)
                || (strFilePropertySecrecyLevel == AddInsLibrary.InfomationForm.SECRECY_PROPERTY_A))
            {
                string PresPath = Pres.Path;
                string FileName = Pres.Name;

                List<string> lstTarGetSecureFolder = frmSet.clsCommonSettting.lstSecureFolder;

                // セキュアフォルダと同一なら保存
                string result = lstTarGetSecureFolder.FirstOrDefault(x => PresPath.Contains(x));
                if (result == null)
                {
                    // セキュアフォルダではない
                    MessageBox.Show(AddInsLibrary.Properties.Resources.msg_warning_save_not_secure, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);

                    string FullPath = Path.Combine(PresPath, FileName);
                    lstPresPath.Add(FullPath);

                    ExecuteSaveAs();

                    return;
                }
            }

            ClearListPath();
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
        /// 保存処理本体
        /// </summary>
        /// <param name="Pres">Presentation情報</param>
        /// <param name="Cancel">保存フラグ</param>
        /// <param name="SaveAsUI">キャンセルフラグ</param>
        public void PresentationSave(Microsoft.Office.Interop.PowerPoint.Presentation Pres, ref bool Cancel, bool SaveAsUI)
        {
            // PowerPoint画面が表示されていない場合は設定画面を表示しない
            if (this.Application.Visible == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                return;
            }

            // プロパティ情報取得
            PowerPointAddInSAB.SettingForm frmSet = new PowerPointAddInSAB.SettingForm();
            frmSet.propPres = Pres;

            // 共通設定エラー時処理
            if (frmSet.commonFileReadCompleted == false)
            {
                return;
            }

            // スライド数が0のときは登録不要
            if (frmSet.GetSlideCount() <= 0)
            {
                return;
            }

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
            }
            else
            {
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
        }

        /// <summary>
        /// 「名前を付けて保存」前イベント
        /// </summary>
        /// <param name="Pres">Presentation情報</param>
        /// <param name="Cancel">キャンセルフラグ</param>
        private void ApplicationPresentationBeforeSave(Microsoft.Office.Interop.PowerPoint.Presentation Pres, ref bool Cancel)
        {
            if (Pres == this.Application.ActivePresentation)
            {
                PresentationSave(Pres, ref Cancel, true);
            }
        }

        /// <summary>
        /// カスタムリボンXMLクラスを返す様にオーバーライド
        /// </summary>
        /// <param name="serviceGuid"></param>
        /// <returns></returns>
        protected override object RequestService(Guid serviceGuid)
        {
            if (serviceGuid == typeof(Office.IRibbonExtensibility).GUID)
            {
                return _ribbon ?? (_ribbon = new Ribbon());
            }
            return base.RequestService(serviceGuid);
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
