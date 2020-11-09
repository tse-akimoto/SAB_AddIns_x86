using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddInSAB
{
    public partial class RibbonDocumentManagement
    {
        private void RibbonDocumentManagement_Load(object sender, RibbonUIEventArgs e)
        {
            // 言語設定読込み  // step2
            Localizable();
        }

        /// <summary>
        /// カスタムリボンローカライズ
        /// </summary>
        private void Localizable()
        {
            WordAddInSAB.SettingForm frmSet = new WordAddInSAB.SettingForm();
            System.Globalization.CultureInfo culture = System.Globalization.CultureInfo.GetCultureInfo(frmSet.clsCommonSettting.strCulture);

            System.Threading.Thread.CurrentThread.CurrentUICulture = culture;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonDocumentManagement));

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
        /// SABボタン押下処理
        /// </summary>
        private void buttonSAB_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 現在開いているファイルを確認
                int iOpenFileCnt = 0;

                Word.Application WordApp = (Word.Application)global::WordAddInSAB.Globals.ThisAddIn.Application;
                iOpenFileCnt = WordApp.Documents.Count;

                if (iOpenFileCnt == 0)
                {
                    return;
                }

                // 言語設定を取得
                string currentUICulture = System.Threading.Thread.CurrentThread.CurrentUICulture.ToString();

                // プロパティ情報取得
                SettingForm settingForm = new SettingForm();

                // 共通設定エラー時処理
                if (settingForm.commonFileReadCompleted == false)
                {
                    return;
                }

                // 共通設定ファイルと言語設定が異なる場合は言語設定を反映
                string strCulture = settingForm.clsCommonSettting.strCulture;
                if (currentUICulture != strCulture)
                {
                    System.Globalization.CultureInfo culture = System.Globalization.CultureInfo.GetCultureInfo(strCulture);
                    System.Threading.Thread.CurrentThread.CurrentUICulture = culture;

                    settingForm = new SettingForm();
                }

                // プロパティのセキュリティ情報が存在するか
                if (settingForm.IsSecrecyInfoRegistered() == true)
                {
                    string filePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
                    string filePropertyClassNo = string.Empty;      // ファイルプロパティ情報 文書No.
                    string filePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

                    // ファイルプロパティ情報取得
                    settingForm.GetDocumentProperty(ref filePropertySecrecyLevel, ref filePropertyClassNo, ref filePropertyOfficeCode);

                    // プロパティ情報があればインフォメーション画面表示
                    AddInsLibrary.InfomationForm infomationForm =
                        new AddInsLibrary.InfomationForm(filePropertySecrecyLevel);

                    // SAB機密区分表示画面を表示
                    System.Windows.Forms.DialogResult dialogResult = infomationForm.ShowDialog();

                    // 修正ボタンが押されたら設定画面を表示
                    if (dialogResult == System.Windows.Forms.DialogResult.OK)
                    {
                        settingForm.ShowDialog();
                    }
                }
                else
                {
                    // プロパティ情報がなければ設定画面表示
                    settingForm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }
    }
}
