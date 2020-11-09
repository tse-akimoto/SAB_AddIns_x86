using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.Diagnostics;

namespace ExcelAddInSAB
{
    public partial class SettingForm : AddInsLibrary.SettingForm
    {
        #region <コンストラクタ>
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SettingForm()
        {
            // タイトル
            System.Diagnostics.FileVersionInfo ver = FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string AssemblyName = ver.FileVersion;
            this.Text = this.Text + " " + AssemblyName;
        }
        #endregion

        #region <メソッド>
        /// <summary>
        /// Excelのオブジェクトを削除
        /// </summary>
        /// <param name="excelWorkbooks">対象のExcel</param>
        /// <param name="shapeName">削除するオブジェクト名</param>
        private void DeleteExcelShapes(ref Excel.Workbook excelWorkbooks, string shapeName)
        {
            // すべてのExcelシートからスタンプを削除
            foreach (Excel.Worksheet sheet in excelWorkbooks.Sheets)
            {
                Excel.Shapes excelShapes = (Excel.Shapes)sheet.Shapes;

                // スタンプ画像かオブジェクト名で判定して削除
                foreach (Excel.Shape shape in excelShapes)
                {
                    if (shape.Name == shapeName) shape.Delete();
                }
            }
        }

        /// <summary>
        /// Excelの全てのシートにスタンプを貼付け
        /// </summary>
        /// <param name="ExcelWorkbooks">対象のExcel</param>
        /// <param name="imageFilePath">貼付けるスタンプの画像ファイルパス</param>
        /// <param name="stampWidth">補正する画像の横幅</param>
        /// <param name="stampHeight">補正する画像の縦幅</param>
        private void AddStampPicture(ref Excel.Workbook ExcelWorkbooks, string imageFilePath, float stampWidth, float stampHeight, string stampName)
        {
            foreach (var sheet in ExcelWorkbooks.Sheets)
            {
                Excel.Worksheet workSheet = (Excel.Worksheet)sheet;

                // 画像貼付処理
                Excel.Shapes excelShapes = (Excel.Shapes)workSheet.Shapes;
                Excel.Shape stampShape = excelShapes.AddPicture(imageFilePath,
                                                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                Microsoft.Office.Core.MsoTriState.msoTrue,
                                                                0,
                                                                0,
                                                                (float)stampWidth,
                                                                (float)stampHeight);

                // 貼付けた画像のオブジェクト名を設定
                stampShape.Name = stampName;
            }
        }
        #endregion


        #region <Overrideメソッド>
        /// <summary>
        /// スタンプ貼付け処理
        /// </summary>
        protected override Boolean SetStamp(Secrecy secrecyLevel)
        {
            // 一時ファイル名取得
            string imageFilePath = System.IO.Path.GetTempFileName();

            Bitmap bmpSrc = null;

            try
            {
                // 現在開いているExcelを取得
                Excel.Application excelApp = (Excel.Application)global::ExcelAddInSAB.Globals.ThisAddIn.Application;

                // すべてのExcelシートを取得
                Excel.Workbook excelWorkbooks = (Excel.Workbook)excelApp.ActiveWorkbook;


                // スタンプ表示OFF・区分"以外"の場合はスタンプをセットしない
                // スタンプ画像を削除して終了
                if (this.chkChange.Checked == false || this.rdoElse.Checked == true)
                {
                    // 指定した名前のオブジェクトを削除
                    this.DeleteExcelShapes(ref excelWorkbooks, STAMP_SHAPE_NAME);

                    return true;
                }


                // スタンプ画像をリソースから取得
                bmpSrc = this.GetStampImage(secrecyLevel);

                // 画像が取得できない場合は中断
                if (bmpSrc == null) return false;

                // スタンプ倍率変更
                double dStampWidth = bmpSrc.Width / STAMP_MAGNIFICATION;
                double dStampHeight = bmpSrc.Height / STAMP_MAGNIFICATION;

                // 透過処理
                float alpha = (float)(this.nudAlpha.Value * (decimal)0.01);
                bmpSrc = this.CreateAlphaImage(bmpSrc, alpha);

                // ファイルを一時保存
                bmpSrc.Save(imageFilePath, System.Drawing.Imaging.ImageFormat.Png);


                // 指定した名前のオブジェクトを削除
                this.DeleteExcelShapes(ref excelWorkbooks, STAMP_SHAPE_NAME);


                // すべてのExcelシートにスタンプを貼付け
                this.AddStampPicture(ref excelWorkbooks, imageFilePath, (float)dStampWidth, (float)dStampHeight, this.STAMP_SHAPE_NAME);
            }
            catch
            {
                // スタンプ貼り付け失敗
                return false;
            }
            finally
            {
                // 一時ファイル削除
                System.IO.File.Delete(imageFilePath);

                // 解放
                if (bmpSrc != null)
                {
                    bmpSrc.Dispose();
                }
            }

            return true;
        }

        /// <summary>
        /// ドキュメントのプロパティ取得
        /// </summary>
        /// <param name="strClassNo">文書分類番号</param>
        /// <param name="strSecrecyLevel">機密区分</param>
        /// <param name="bStamp">スタンプ有無</param>off
        public override void GetDocumentProperty(ref string strSecrecyLevel, ref string strOfficeCod, ref string strOfficeCode)
        {
            Type tBuiltProp;                                        // プロパティ情報タイプ
            Type tProperty;                                         // プロパティ値タイプ
            object oBuiltProp;                                      // プロパティ情報オブジェクト
            object oPropertyItem;                                   // プロパティアイテムオブジェクト
            object oPropertyValue;                                  // プロパティ値オブジェクト

            // 現在開いているExcelを取得
            Excel.Application ExcelApp = (Excel.Application)global::ExcelAddInSAB.Globals.ThisAddIn.Application;

            // プロパティ情報を取得
            oBuiltProp = ExcelApp.ActiveWorkbook.BuiltinDocumentProperties;

            try
            {
                /////////////////////////////
                // Category（分類項目）情報取得
                /////////////////////////////

                // プロパティ情報タイプを取得
                tBuiltProp = oBuiltProp.GetType();

                // プロパティCategory（分類項目）のアイテム情報を取得
                oPropertyItem = tBuiltProp.InvokeMember("Item", BindingFlags.GetProperty, null, oBuiltProp, new object[] { "Category" });

                // プロパティCategory（分類項目）の値を取得
                tProperty = oPropertyItem.GetType();
                oPropertyValue = tProperty.InvokeMember("Value", BindingFlags.GetProperty, null, oPropertyItem, new object[] { });


                /////////////////////////////
                // keywords（タグ）情報取得
                /////////////////////////////

                // プロパティ情報タイプを取得
                tBuiltProp = oBuiltProp.GetType();

                // プロパティkeywords（タグ）のアイテム情報を取得
                oPropertyItem = tBuiltProp.InvokeMember("Item", BindingFlags.GetProperty, null, oBuiltProp, new object[] { "keywords" });

                // プロパティkeywords（タグ）の値を取得
                tProperty = oPropertyItem.GetType();
                oPropertyValue = tProperty.InvokeMember("Value", BindingFlags.GetProperty, null, oPropertyItem, new object[] { });


                // タグ情報があった場合は機密区分、文書分類番号、スタンプの有無をセットする
                if (oPropertyValue != null)
                {
                    string[] strPropertyData = oPropertyValue.ToString().Split(';');

                    // プロパティの事業所コードを取得
                    if (strPropertyData.Count() > (int)Property.OfficeCode)
                    {
                        strOfficeCode = strPropertyData[(int)Property.OfficeCode].Trim();
                    }
                    else
                    {
                        strOfficeCode = AddInsLibrary.CommonSettings.COMMON_SETDEF_OFFICECODE;
                    }


                    // プロパティの文書番号を取得
                    if (strPropertyData.Count() > (int)Property.ClassNo)
                    {
                        strSecrecyLevel = strPropertyData[(int)Property.ClassNo].Trim();
                    }


                    // プロパティの機密区分を取得
                    if (strPropertyData.Count() > (int)Property.SecrecyLevel)
                    {
                        strSecrecyLevel = strPropertyData[(int)Property.SecrecyLevel].Trim();
                    }
                }
            }
            catch
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msg_err_read_common_file, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }


        /// <summary>
        /// ドキュメントのプロパティ設定
        /// </summary>
        /// <param name="strClassNo"></param>
        /// <param name="strSecrecyLevel"></param>
        public override bool SetDocumentProperty(string strSecrecyLevel)
        {
            Type tBuiltProp;                                        // プロパティ情報タイプ
            Type tProperty;                                         // プロパティ値タイプ
            object oBuiltProp;                                      // プロパティ情報オブジェクト
            object oPropertyItem;                                   // プロパティアイテムオブジェクト

            string strMyOfficeCode = "";    // 自事業所コード

            // 現在開いているExcelを取得
            Excel.Application ExcelApp = (Excel.Application)global::ExcelAddInSAB.Globals.ThisAddIn.Application;

            // プロパティ情報を取得
            oBuiltProp = ExcelApp.ActiveWorkbook.BuiltinDocumentProperties;

            // 自事業所コードに共通設定事業所コードを設定
            strMyOfficeCode = clsCommonSettting.strOfficeCode;

            try
            {
                /////////////////////////////
                // Category（分類項目）情報書込
                /////////////////////////////

                // プロパティ情報タイプを取得
                tBuiltProp = oBuiltProp.GetType();

                // プロパティCategory（分類項目）のアイテム情報を取得
                oPropertyItem = tBuiltProp.InvokeMember("Item", BindingFlags.GetProperty, null, oBuiltProp, new object[] { "Category" });

                // プロパティCategory（分類項目）の値をクリア
                tProperty = oPropertyItem.GetType();
                tProperty.InvokeMember("Value", BindingFlags.SetProperty, null, oPropertyItem, new object[] { "" });


                /////////////////////////////
                // Category（タグ）情報書込
                /////////////////////////////
                string strWritePropertyData = string.Format("{0}; {1}; {2};", strSecrecyLevel, string.Empty, strMyOfficeCode);

                // プロパティ情報タイプを取得
                tBuiltProp = oBuiltProp.GetType();

                // プロパティkeywords（タグ）のアイテム情報を取得
                oPropertyItem = tBuiltProp.InvokeMember("Item", BindingFlags.GetProperty, null, oBuiltProp, new object[] { "keywords" });

                // プロパティkeywords（タグ）の値を設定
                tProperty = oPropertyItem.GetType();
                tProperty.InvokeMember("Value", BindingFlags.SetProperty, null, oPropertyItem, new object[] { strWritePropertyData });
            }
            catch
            {
                // プロパティ設定の失敗
                return false;
            }

            return true;
        }
        #endregion

        private void InitializeComponent()
        {
            ((System.ComponentModel.ISupportInitialize)(this.nudAlpha)).BeginInit();
            this.SuspendLayout();
            // 
            // SettingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.ClientSize = new System.Drawing.Size(651, 356);
            this.Name = "SettingForm";
            this.Text = "SAB 1.0.0.0";
            ((System.ComponentModel.ISupportInitialize)(this.nudAlpha)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
