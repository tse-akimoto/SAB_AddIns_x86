using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Drawing;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

// TODO:  リボン (XML) アイテムを有効にするには、次の手順に従います。

// 1: 次のコード ブロックを ThisAddin、ThisWorkbook、ThisDocument のいずれかのクラスにコピーします。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. ボタンのクリックなど、ユーザーの操作を処理するためのコールバック メソッドを、このクラスの
//    "リボンのコールバック" 領域に作成します。メモ: このリボンがリボン デザイナーからエクスポートされたものである場合は、
//    イベント ハンドラー内のコードをコールバック メソッドに移動し、リボン拡張機能 (RibbonX) のプログラミング モデルで
//    動作するように、コードを変更します。

// 3. リボン XML ファイルのコントロール タグに、コードで適切なコールバック メソッドを識別するための属性を割り当てます。  

// 詳細については、Visual Studio Tools for Office ヘルプにあるリボン XML のドキュメントを参照してください。


namespace PowerPointAddInSAB
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        Microsoft.Office.Interop.PowerPoint.Application _app;

        public Ribbon()
        {

        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            _app = Globals.ThisAddIn.Application;
            //_app = new Microsoft.Office.Interop.PowerPoint.Application();
            return GetResourceText("PowerPointAddInSAB.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック
        //ここにコールバック メソッドを作成します。コールバック メソッドの追加方法の詳細については、http://go.microsoft.com/fwlink/?LinkID=271226 にアクセスしてください。

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbonUI.InvalidateControl("tabGCPDocumentManagement");
            ribbonUI.InvalidateControl("grpDocumentManagement");
            ribbonUI.InvalidateControl("btnSAB");

            this.ribbon = ribbonUI;
        }


        public stdole.IPictureDisp getImage(Office.IRibbonControl control)
        {
            return (stdole.IPictureDisp)IPictureDispHost.GetIPictureDispFromPicture(Properties.Resources.StampB);
        }


        public string getLabel(Office.IRibbonControl control)
        {
            // ローカライズ
            string ret = "";
            if (control.Id == "tabGCPDocumentManagement")
            {
                ret = Properties.Resources.tabGCPDocumentManagement_Label;
            }
            else if (control.Id == "grpDocumentManagement")
            {
                ret = Properties.Resources.grpDocumentManagement_Label;
            }
            else if (control.Id == "btnSAB")
            {
                ret = Properties.Resources.btnSAB_Label;
            }

            return ret;
        }

        public void FileSaveOverride(Office.IRibbonControl control, ref bool cancelDefault)
        {
            if (control.Context.Presentation == _app.ActivePresentation)
            {
                if (File.Exists(Path.Combine(_app.ActivePresentation.Path, _app.ActivePresentation.Name)) == true)
                {
                    // ファイルがある場合、上書き保存
                    Globals.ThisAddIn.PresentationSave(_app.ActivePresentation, ref cancelDefault, false);
                    cancelDefault = true;
                }
                else
                {
                    // ファイルがない場合は、後続の「名前を付けて保存」の処理に任せる。
                    cancelDefault = false;
                }
            }
            else
            {
                cancelDefault = false;
            }

            Console.WriteLine("上書き保存：" + _app.ActivePresentation.FullName);
        }
        public void buttonSAB_Click(Office.IRibbonControl control)
        {
            // 現在開いているファイルを確認
            int iOpenFileCnt = 0;

            PowerPoint.Application pptApp = (PowerPoint.Application)global::PowerPointAddInSAB.Globals.ThisAddIn.Application;
            iOpenFileCnt = pptApp.Presentations.Count;

            if (iOpenFileCnt == 0)
            {
                return;
            }

            // 言語設定を取得
            string currentUICulture = System.Threading.Thread.CurrentThread.CurrentUICulture.ToString();

            // プロパティ情報取得
            SettingForm settingForm = new SettingForm();
            settingForm.propPres = pptApp.ActivePresentation;

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

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    internal sealed class IPictureDispHost : AxHost
    {
        /// <summary>
        /// Default Constructor, required by the framework.
        /// </summary>
        private IPictureDispHost() : base(string.Empty) { }
        /// <summary>
        /// Convert the image to an ipicturedisp.
        /// </summary>
        /// <param name="image">The image instance</param>
        /// <returns>The picture dispatch object.</returns>
        public new static object GetIPictureDispFromPicture(Image image)
        {
            return AxHost.GetIPictureDispFromPicture(image);
        }
        /// <summary>
        /// Convert the dispatch interface into an image object.
        /// </summary>
        /// <param name="picture">The picture interface</param>
        /// <returns>An image instance.</returns>
        public new static Image GetPictureFromIPicture(object picture)
        {
            return AxHost.GetPictureFromIPicture(picture);
        }
    }
}
