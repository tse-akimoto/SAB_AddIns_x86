using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;

namespace OutlookAddInSAB
{
    class PropertiesKeyList
    {
        /// <summary>
        /// プロパティ一覧
        /// </summary>
        public const string STR_TAG_DC = "dc";
        public const string STR_TAG_CP = "cp";
        public const string STR_TAG_DCTRMS = "dcterms";
        public const string STR_TAG_DCMITYPE = "dcmitype";
        public const string STR_TAG_XSI = "xsi";

        public const string STR_TITLE = "dc:title";
        public const string STR_SUBJECT = "dc:subject";
        public const string STR_CREATOR = "dc:creator";
        public const string STR_KEYWORDS = "cp:keywords";
        public const string STR_DESCRIPTION = "dc:description";
        public const string STR_LAST_MODIFIED_BY = "cp:lastModifiedBy";
        public const string STR_REVISION = "cp:revision";
        public const string STR_CREATED = "dcterms:created";
        public const string STR_MODIFIED = "dcterms:modified";
        public const string STR_CATEGORY = "cp:category";
        public const string STR_CONTENT_STATUS = "cp:contentStatus";
        public const string STR_LANGUAGE = "dc:language";
        public const string STR_VERSION = "cp:version";

        public const string STR_CORE_PROPERTIES = "//cp:coreProperties/";

        // OFFICEの保護状態が最終版の時の判定用
        public const string STR_FINAL_CONTENT = "最終版";

        /// <summary>
        /// プロパティの一覧を返す
        /// </summary>
        /// <returns></returns>
        public static List<string> getPropertiesKeyList()
        {
            List<string> list = new List<string>();

            list.Add(STR_TITLE);
            list.Add(STR_SUBJECT);
            list.Add(STR_CREATOR);
            list.Add(STR_KEYWORDS);
            list.Add(STR_DESCRIPTION);
            list.Add(STR_LAST_MODIFIED_BY);
            list.Add(STR_REVISION);
            list.Add(STR_CREATED);
            list.Add(STR_MODIFIED);
            list.Add(STR_CATEGORY);
            list.Add(STR_CONTENT_STATUS);
            list.Add(STR_LANGUAGE);
            list.Add(STR_VERSION);

            return list;
        }
    }

    class PropertiesSchemaList
    {
        // スキーマ定義
        private const string corePropertiesSchema = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private const string dcPropertiesSchema = "http://purl.org/dc/elements/1.1/";
        private const string dctermsPropertiesSchema = "http://purl.org/dc/terms/";
        private const string dcmitypePropertiesSchema = "http://purl.org/dc/dcmitype/";
        private const string xsiPropertiesSchema = "http://www.w3.org/2001/XMLSchema-instance";

        public static string CorePropertiesSchema
        {
            get
            {
                return corePropertiesSchema;
            }
        }

        public static string DcPropertiesSchema
        {
            get
            {
                return dcPropertiesSchema;
            }
        }

        public static string DctermsPropertiesSchema
        {
            get
            {
                return dctermsPropertiesSchema;
            }
        }

        public static string DcmitypePropertiesSchema
        {
            get
            {
                return dcmitypePropertiesSchema;
            }
        }

        public static string XsiPropertiesSchema
        {
            get
            {
                return xsiPropertiesSchema;
            }
        }
    }

    class AccessToProperties
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public AccessToProperties()
        {

        }

        /// <summary>
        /// プロパティを読み込みます
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="propertiesTable"></param>
        /// <returns>取得したプロパティ一覧</returns>
        public Dictionary<string, string> ReadProperties(ClsFilePropertyList list, string fileType, string str_exception)  // 20171011 修正（ファイル読込時のエラー理由を保存）
        {
            Dictionary<string, string> retTable = new Dictionary<string, string>();

            System.IO.FileInfo fi = new System.IO.FileInfo(list.filePath);
            if (fi.Length == 0)
            {
                // ファイルサイズがゼロの場合プロパティが存在しないのでエラー
                // 20171011 追加 （読取専用、ファイ存在しない以外のエラー）
                //error_reason = ListForm.LIST_VIEW_NA;
                return retTable;
            }

            SpreadsheetDocument excel = null;
            WordprocessingDocument word = null;
            PresentationDocument ppt = null;

            CoreFilePropertiesPart coreFileProperties;

            // 20171011 修正（xlsxとdocxのみファイルを開いている状態でファイルアクセスするとエラーになる為の回避対応）
            try
            {
                // ファイルのプロパティ領域を開く
                switch (fileType)
                {
                    case "Excel": // エクセルの場合
                        excel = SpreadsheetDocument.Open(list.filePath, false);
                        coreFileProperties = excel.CoreFilePropertiesPart;
                        break;
                    case "Word": // ワードの場合
                        word = WordprocessingDocument.Open(list.filePath, false);
                        coreFileProperties = word.CoreFilePropertiesPart;
                        break;
                    case "PowerPoint": // パワポの場合
                        ppt = PresentationDocument.Open(list.filePath, false);
                        coreFileProperties = ppt.CoreFilePropertiesPart;
                        break;
                    default:
                        // 異常なファイル
                        // 20171228 追加（エラー理由保存）
                        //error_reason = ListForm.LIST_VIEW_NA;
                        return retTable;
                }
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_CP, PropertiesSchemaList.CorePropertiesSchema);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DC, PropertiesSchemaList.DcPropertiesSchema);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DCTRMS, PropertiesSchemaList.DctermsPropertiesSchema);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_DCMITYPE, PropertiesSchemaList.DcmitypePropertiesSchema);
                nsManager.AddNamespace(PropertiesKeyList.STR_TAG_XSI, PropertiesSchemaList.XsiPropertiesSchema);

                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(coreFileProperties.GetStream());

                // プロパティのキーリストを作成
                List<string> propertieslist = PropertiesKeyList.getPropertiesKeyList();

                // 全キーリストを見て存在するデータを取得
                foreach (string key in propertieslist)
                {
                    // 書き込み先のキーワードを指定
                    string searchString = string.Format(PropertiesKeyList.STR_CORE_PROPERTIES + "{0}", key);
                    // 書き込み先を検索
                    XmlNode xNode = xdoc.SelectSingleNode(searchString, nsManager);

                    if (xNode != null)
                    {
                        // 読み込む
                        retTable.Add(key, xNode.InnerText);
                    }
                }

                // ファイルのプロパティ領域を閉じる
                if (excel != null) excel.Close();
                if (word != null) word.Close();
                if (ppt != null) ppt.Close();

                // プロパティ取得
                string[] propertyData = null;
                if (retTable.ContainsKey(PropertiesKeyList.STR_KEYWORDS) != false)
                {
                    propertyData = retTable[PropertiesKeyList.STR_KEYWORDS].Split(';');
                }

                if (propertyData != null)
                {
                    list.fileSecrecy = propertyData[0].TrimEnd();

                    // 機密区分が登録されているか判定
                    if (!string.IsNullOrEmpty(list.fileSecrecy))
                    {
                        // 登録あり
                        if (propertyData.Count() >= 1)
                        {
                            list.fileClassification = propertyData[1].TrimEnd();
                        }
                        if (propertyData.Count() >= 2)
                        {
                            list.fileOfficeCode = propertyData[2].TrimEnd();
                        }

                        // 他事務所ファイル判定
                        if (Globals.ThisAddIn.clsCommonSettings.strOfficeCode != list.fileOfficeCode.Trim())
                        {
                            // 他事務所ファイルの場合、登録なし扱にする
                            list.fileSecrecy = "";
                        }
                    }
                }
                else
                {
                    list.fileSecrecy = "Notting";
                }
            }
#if false
            #region HyperLink修復

            // ■ ADD TSE Kitada
            // HyperLinkが破損している場合に、そのリンクを書き直して正常にOPEN出来るようにする。
            // 但し、ドキュメントの中身を直接書き換える処理のため、見送る。
            catch (OpenXmlPackageException ope)
            {
                if (ope.ToString().Contains("Invalid Hyperlink"))
                {
                    // HyperLinkの破損が原因なので、内部のリンクを修正する
                    using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                    }

                    if (count >= 1)
                    {
                        // 2回実行してダメだったので終了
                        error_reason = ListForm.LIST_VIEW_NA;
                    }
                    else
                    {
                        // もう一度トライ
                        retTable = ReadProperties(filePath, filetype, ref error_reason, ref str_exception, 1);
                    }
                }

                if (excel != null) excel.Close();
                if (word != null) word.Close();
                if (ppt != null) ppt.Close();
            }
            #endregion
#endif
            catch (Exception e)
            {
                str_exception += "file : " + list.filePath + "\r\n\r\n";
                str_exception += "error : " + e.ToString();
                // xlsxとdocxのみファイルを開いている状態でファイルアクセスした場合
                // ファイルが開かれている場合のエラー
                //error_reason = ListForm.LIST_VIEW_NA;

                if (excel != null) excel.Close();
                if (word != null) word.Close();
                if (ppt != null) ppt.Close();

                return retTable;
            }
            return retTable;
        }

#if false
        #region HyperLink修復

        /// <summary>
        /// OFFICEドキュメント内の破損しているハイパーリンクを置き換える文字列を返します
        /// </summary>
        /// <param name="brokenUri"></param>
        /// <returns></returns>
        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }
        #endregion
#endif
    }
}
