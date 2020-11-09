using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text.pdf;

namespace OutlookAddInSAB
{
    public class ClsFilePropertyList
    {
        #region 定義

        public Outlook.Attachment attachment { get; set; }
        /// <summary>
        /// ファイル名
        /// </summary>
        public string fileName { get; set; }
        /// <summary>
        /// ファイルパス
        /// </summary>
        public string filePath { get; set; }
        /// <summary>
        /// ファイル拡張子
        /// </summary>
        public string fileExtension { get; set; }
        /// <summary>
        /// 機密区分
        /// </summary>
        public string fileSecrecy { get; set; }
        private int rank;
        public int fileSecrecyRank { get { return rank; } }
        /// <summary>
        /// 事業所
        /// </summary>
        public string fileOfficeCode { get; set; }
        /// <summary>
        /// 文書分類
        /// </summary>
        public string fileClassification { get; set; }
        public List<ClsFilePropertyList> file_list { get; set; }

        /// <summary>
        /// ZIP解凍時エラーフラグ
        /// </summary>
        public bool ZipError { get; set; }


        const string NONE = "None";
        const string PDF = "PDF";
        const string EXCEL = "Excel";
        const string WORD = "Word";
        const string POWERPOINT = "PowerPoint";

        const string SECRECY_NONE = "SecrecyNone";

        #endregion

        public ClsFilePropertyList()
        {
            ZipError = false;
        }

        #region メソッド

        /// <summary>
        /// ファイルの機密区分を取得するメソッド
        /// </summary>
        /// <param name="list">対象ファイルのデータクラス</param>
        /// <returns>取得結果</returns>
        public bool FileCheck(ClsFilePropertyList list)
        {
            bool result = false;
            try
            {
                // ファイルが存在しない場合は処理を終了
                if (File.Exists(list.filePath) == false) return result;

                string fileType = NONE;

                var fileTypeDictionary = new Dictionary<string, string>(){
                    { ".pdf", PDF},
                    { ".xlsx", EXCEL}, { ".xlsm", EXCEL }, { ".xls", EXCEL},
                    { ".docx", WORD}, { ".doc", WORD},
                    { ".pptx", POWERPOINT}, { ".ppt", POWERPOINT} };
                if (fileTypeDictionary.TryGetValue(list.fileExtension, out fileType))
                {
                    fileType = fileTypeDictionary[list.fileExtension];
                }
                else
                {
                    fileType = NONE;
                }

                // 文書分類と機密区分を取得
                switch (fileType)
                {
                    case PDF:
                        ReadByPDF(list);
                        break;
                    case EXCEL:
                        if (ExtensionOpenXMLCheck(list.fileExtension) == true)
                        {
                            AccessToProperties atp = new AccessToProperties();
                            Dictionary<string, string> propertyTable = new Dictionary<string, string>();
                            string exception = "";
                            propertyTable = atp.ReadProperties(list, fileType, exception);
                        }
                        else
                        {
                            ReadByDSO(list);
                        }
                        break;
                    case WORD:
                        if (ExtensionOpenXMLCheck(list.fileExtension) == true)
                        {
                            AccessToProperties atp = new AccessToProperties();
                            Dictionary<string, string> propertyTable = new Dictionary<string, string>();
                            string exception = "";
                            propertyTable = atp.ReadProperties(list, fileType, exception);
                        }
                        else
                        {
                            ReadByDSO(list);
                        }
                        break;
                    case POWERPOINT:
                        if (ExtensionOpenXMLCheck(list.fileExtension) == true)
                        {
                            AccessToProperties atp = new AccessToProperties();
                            Dictionary<string, string> propertyTable = new Dictionary<string, string>();
                            string exception = "";
                            propertyTable = atp.ReadProperties(list, fileType, exception);
                        }
                        else
                        {
                            ReadByDSO(list);
                        }
                        break;
                    case NONE:
                        list.fileSecrecy = SECRECY_NONE;
                        break;
                    default:
                        list.fileSecrecy = "";
                        break;
                }
                list.rank = SetFileSecrecyRank(list.fileSecrecy);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
            return result;
        }

        /// <summary>
        /// PDFのプロパティ取得
        /// </summary>
        /// <param name="list">対象のファイル</param>
        /// <returns>取得結果</returns>
        private bool ReadByPDF(ClsFilePropertyList list)
        {
            bool result = false;

            PdfReader reader = new PdfReader(list.filePath);

            // プロパティ情報のリスト化
            List<string> name = new List<string>(reader.Info.Keys);
            List<string> val = new List<string>(reader.Info.Values);
            reader.Close();

            int Property_Count = name.IndexOf("Keywords");
            if (name.Contains("Keywords") == false) return result;

            string property = val[Property_Count];
            string[] propertyValue = property.Split(';');

            // propertyValue[0]が機密区分、[1]が文書分類、[2]が事業所?
            if (propertyValue.Count() != 0)
            {
                list.fileSecrecy = propertyValue[0];
            }
            if (propertyValue.Count() >= 2)
            {
                list.fileClassification = propertyValue[1];
            }
            if (propertyValue.Count() >= 3)
            {
                list.fileOfficeCode = propertyValue[2];
            }

            // 他事務所ファイル判定
            if (Globals.ThisAddIn.clsCommonSettings.strOfficeCode != list.fileOfficeCode.Trim())
            {
                // 他事務所ファイルの場合、登録なし扱にする
                list.fileSecrecy = "";
            }

            return result;
        }

        /// <summary>
        /// officeファイルがOpenXMLで書き込めるかを判定
        /// </summary>
        /// <param name="fileExtension">対象のファイルの拡張子</param>
        /// <returns>判定結果</returns>
        private bool ExtensionOpenXMLCheck(string fileExtension)
        {
            bool result = false;
            string[] openXML_Narrow = new string[] { ".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pptx", ".ppt" };
            for (int i = 0; i < openXML_Narrow.Count(); i++)
            {
                if (fileExtension == openXML_Narrow[i])
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// Officeファイルのプロパティ取得
        /// </summary>
        /// <param name="list">対象のファイル</param>
        /// <returns>取得結果</returns>
        private bool ReadByDSO(ClsFilePropertyList list)
        {
            bool result = false;
            DSOFile.OleDocumentProperties docProperty = new DSOFile.OleDocumentProperties();
            DSOFile.SummaryProperties summary;
            try
            {
                // ファイルを開く
                docProperty.Open(list.filePath, false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);
                // プロパティの取得
                summary = docProperty.SummaryProperties;

                if (summary.Keywords != null)
                {
                    string[] PropertyData = summary.Keywords.Split(';');
                    if (PropertyData[0] != "")
                    {
                        list.fileSecrecy = PropertyData[0].Trim(); // 機密区分を取得
                    }
                    else
                    {
                        list.fileSecrecy = "区分なし";
                    }
                    // プロパティが2つ以上ある場合文書分類を取得
                    if (PropertyData.Count() >= 2)
                    {
                        list.fileClassification = PropertyData[1].Trim();
                    }
                    // プロパティが3つ以上ある場合事業所を取得
                    if (PropertyData.Count() >= 3)
                    {
                        list.fileOfficeCode = PropertyData[2].Trim();
                    }

                    // 他事務所ファイル判定
                    if (Globals.ThisAddIn.clsCommonSettings.strOfficeCode != list.fileOfficeCode.Trim())
                    {
                        // 他事務所ファイルの場合、登録なし扱にする
                        list.fileSecrecy = "";
                    }
                }
                else
                {
                    list.fileSecrecy = "区分なし";
                }
                result = true;
            }
            finally
            {
                docProperty.Close();
            }
            return result;
        }

        /// <summary>
        /// 機密区分の設定値取得
        /// </summary>
        /// <param name="list">対象ファイルのデータクラス</param>
        /// <returns>機密区分の設定値</returns>
        private int SetFileSecrecyRank(string secrecy)
        {
            int result;
            switch (secrecy)
            {
                case "区分なし":
                case SECRECY_NONE:
                    result = 4;
                    break;
                case "S秘":
                case "SecrecyS":
                    result = 1;
                    break;
                case "A秘":
                case "SecrecyA":
                    result = 2;
                    break;
                case "B秘":
                case "SecrecyB":
                    result = 3;
                    break;
                default:
                    result = 0;
                    break;
            }
            return result;
        }

        #endregion

    }
}
