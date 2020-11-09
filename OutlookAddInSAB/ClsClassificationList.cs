using System.Collections.Generic;
using System.Text;
using System.IO;
using AddInsLibrary;

namespace OutlookAddInSAB
{
    public class ClsClassificationList
    {
        public class Manager
        {
            enum Classification
            {
                executive, manager
            }

            /// <summary>
            /// 役職コード
            /// </summary>
            private string StrClassification;
            /// <summary>
            /// 役職名
            /// </summary>
            private string StrManager;
            public string classification
            {
                get {return StrClassification; }
                set {
                    if (value == "0")
                        StrClassification = Classification.executive.ToString();
                    else
                        StrClassification = Classification.manager.ToString();
                    }
            }
            public string manager { get { return StrManager; } set { StrManager = value; } }

        }

        #region リスト読み込み処理

        /// <summary>
        /// 関連会社リストの読み込み
        /// </summary>
        /// <param name="clsCommonSetting">共通設定ファイルの設定値クラス</param>
        /// <returns>読み込み結果</returns>
        public List<string> AssociateList(CommonSettings clsCommonSetting)
        {
            string line = "";
            string dataFileServerPath = clsCommonSetting.strGroupDomainListServerPath;
            string dataFileLocalPath = clsCommonSetting.strGroupDomainListLocalPath;
            string dataFilePath = "";

            if (File.Exists(dataFileServerPath))
            {
                // サーバにファイルが存在する場合
                dataFilePath = dataFileServerPath;
            }
            else
            {
                // サーバにファイルが存在しない場合
                dataFilePath = dataFileLocalPath;
            }

            var list = new List<string>();

            using (var reader = new StreamReader(dataFilePath, Encoding.GetEncoding("Shift_JIS")))
            {
                while ((line = reader.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            return list;
        }

        /// <summary>
        /// 役職リストの読み込み
        /// </summary>
        /// <param name="clsCommonSetting">共通設定ファイルの設定値クラス</param>
        /// <returns>読み込み結果</returns>
        public List<Manager> ManagerList(CommonSettings clsCommonSetting)
        {
            string line = "";
            string dataFileServerPath = clsCommonSetting.strManagerListServerPath;
            string dataFileLocalPath = clsCommonSetting.strManagerListLocalPath;
            string dataFilePath = "";

            if (File.Exists(dataFileServerPath))
            {
                // サーバにファイルが存在する場合
                dataFilePath = dataFileServerPath;
            }
            else
            {
                // サーバにファイルが存在しない場合
                dataFilePath = dataFileLocalPath;
            }

            var list = new List<Manager>();

            using (var reader = new StreamReader(dataFilePath, Encoding.GetEncoding("Shift_JIS")))
            {
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(',');
                    list.Add(new Manager { classification = arr[0], manager = arr[1] });
                }
            }
            return list;
        }

        #endregion
    }
}
