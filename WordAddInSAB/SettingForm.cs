using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Reflection;
using System.Diagnostics;

namespace WordAddInSAB
{
    public partial class SettingForm : AddInsLibrary.SettingForm
    {
        #region <定数>
        private int MARGIN_TOP = 10;
        private int MARGIN_RIGHT = 10;
        #endregion


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
        /// Wordのオブジェクトを削除
        /// </summary>
        /// <param name="powerPoint">対象のWord</param>
        /// <param name="shapeName">オブジェクト名</param>
        private void DeleteWordShapes(ref Word.Document wordDocument, string shapeName)
        {
            // スタンプ画像かオブジェクト名で判定して削除
            foreach (Word.Shape shape in wordDocument.Shapes)
            {
                if (shape.Name == shapeName) shape.Delete();
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
                // 現在開いているWordの取得
                Word.Application WordApp = (Word.Application)global::WordAddInSAB.Globals.ThisAddIn.Application;

                // アクティブなドキュメントを取得
                Word.Document document = WordApp.Application.ActiveDocument;

                // スタンプ表示OFF・区分"以外"の場合はスタンプをセットしない
                // スタンプ画像を削除して終了
                if (this.chkChange.Checked == false || this.rdoElse.Checked == true)
                {
                    // 指定した名前のオブジェクトを削除
                    this.DeleteWordShapes(ref document, STAMP_SHAPE_NAME);

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
                this.DeleteWordShapes(ref document, STAMP_SHAPE_NAME);


                // スタンプの右上位置を算出
                float topLocation = (float)0 - document.PageSetup.TopMargin + MARGIN_TOP;
                float leftLocation = document.PageSetup.PageWidth - document.PageSetup.RightMargin - (float)dStampWidth - MARGIN_RIGHT;

                // 画像貼付処理
                Word.Shape stampShape = document.Shapes.AddPicture(imageFilePath,
                                                                   Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                   Microsoft.Office.Core.MsoTriState.msoTrue,
                                                                      leftLocation,
                                                                   topLocation,
                                                                   dStampWidth,
                                                                   dStampHeight,
                                                                   document.Range(System.Type.Missing,
                                                                   System.Type.Missing));
                // 貼付けた画像のオブジェクト名を設定
                stampShape.Name = this.STAMP_SHAPE_NAME;
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

            // 現在開いているWordの取得
            Word.Application WordApp = (Word.Application)global::WordAddInSAB.Globals.ThisAddIn.Application;

            // プロパティ情報を取得
            oBuiltProp = WordApp.ActiveDocument.BuiltInDocumentProperties;

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

            // 現在開いているWordの取得
            Word.Application WordApp = (Word.Application)global::WordAddInSAB.Globals.ThisAddIn.Application;

            // プロパティ情報を取得
            oBuiltProp = WordApp.ActiveDocument.BuiltInDocumentProperties;

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
    }
}
