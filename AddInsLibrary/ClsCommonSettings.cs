using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace AddInsLibrary
{
    /// <summary>
    /// 共通設定ファイルの格納先
    /// </summary>
    public class CommonSettingStoring
    {
        /// <summary>
        /// 共通設定 格納フォルダ名
        /// </summary>
        public string COMMON_SETFOLDERNAME = "SAB";

        /// <summary>
        /// 共通設定 ファイル名
        /// </summary>
        public string COMMON_SETFILENAME = "common_setting.config";
    }

    /// <summary>
    /// 共通設定ファイルの設定値
    /// </summary>
    public class CommonSettings
    {
        #region <クラス項目定義>

        /// <summary>
        /// 事業所コード
        /// </summary>
        public string strOfficeCode { get; set; }

        /// <summary>
        /// 機密区分
        /// </summary>
        public string strDefaultSecrecyLevel { get; set; }

        /// <summary>
        /// 言語設定
        /// </summary>
        public string strCulture { get; set; }

        /// <summary>
        /// 文書のサーバーパス
        /// </summary>
        public string strSABListServerPath { get; set; }

        /// <summary>
        /// 文書のローカルパス
        /// </summary>
        public string strSABListLocalPath { get; set; }

        /// <summary>
        /// zip一時解凍先
        /// </summary>
        public string strTempPath { get; set; }

        /// <summary>
        /// セキュアフォルダリスト
        /// </summary>
        public List<string> lstSecureFolder { get; set; }

        /// <summary>
        /// 「最終版」を表す文字列
        /// </summary>
        public List<string> lstFinal { get; set; }

        /// <summary>
        /// 関連会社ファイルパス(サーバー)
        /// </summary>
        public string strGroupDomainListServerPath { get; set; }

        /// <summary>
        /// 関連会社ファイルパス(ローカル)
        /// </summary>
        public string strGroupDomainListLocalPath { get; set; }

        /// <summary>
        /// 役職者ファイルパス(サーバー)
        /// </summary>
        public string strManagerListServerPath { get; set; }

        /// <summary>
        /// 役職者ファイルパス(ローカル)
        /// </summary>
        public string strManagerListLocalPath { get; set; }

        /// <summary>
        /// zip化強制レベル
        /// </summary>
        public int zipLevel { get; set; }

        #endregion

        #region <定数定義>

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 機密区分
        /// </summary>
        public const string COMMON_SETDEF_SECLV = "SecrecyS";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 事業所コード
        /// </summary>
        public const string COMMON_SETDEF_OFFICECODE = "HLI";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 文書のサーバーパス
        /// </summary>
        public const string COMMON_SETDEF_SABLISTSERVERPATH = "C:\\SAB_TEST_SRV\\GCPList.xlsx";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 文書のローカルパス
        /// </summary>
        public const string COMMON_SETDEF_SABLISTLOCALPATH = "C:\\SAB_TEMP\\GCPList.xlsx";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// zip一時解凍先
        /// </summary>
        public const string COMMON_SETDEF_TEMPPATH = "C:\\SAB_TEMP\\WORK";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// セキュアフォルダリスト
        /// </summary>
        public const string COMMON_SETDEF_SECUREFOLDER_1 = "\\" + "\\SRV-FS001";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// セキュアフォルダリスト
        /// </summary>
        public const string COMMON_SETDEF_SECUREFOLDER_2 = "C:\\SAB_TEST_SRV";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 「最終版」を表す文字列
        /// </summary>
        public const string COMMON_SETDEF_FINAL_1 = "最終版";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 「最終版」を表す文字列
        /// </summary>
        public const string COMMON_SETDEF_FINAL_2 = "Final";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 関連会社ファイルパス(サーバー)
        /// </summary>
        public const string COMMON_SETDEF_GROUPDOMAINLISTSERVERPATH = "C:\\SAB_TEST_SRV\\DomainList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 関連会社ファイルパス(ローカル)
        /// </summary>
        public const string COMMON_SETDEF_GROUPDOMAINLISTLOCALPATH = "C:\\SAB_TEMP\\DomainList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 役職者ファイルパス(サーバー)
        /// </summary>
        public const string COMMON_SETDEF_MANAGERLISTSERVERPATH = "C:\\SAB_TEST_SRV\\ManagerList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 役職者ファイルパス(ローカル)
        /// </summary>
        public const string COMMON_SETDEF_MANAGERLISTLOCALPATH = "C:\\SAB_TEMP\\ManagerList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// zip化強制レベル
        /// </summary>
        public const int COMMON_SETDEF_ZIPLEVEL = 1;

        #endregion

        #region コンストラクタ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public CommonSettings()
        {
            // デフォルト事業所コード初期化
            strOfficeCode = COMMON_SETDEF_OFFICECODE;

            // デフォルト機密区分初期化
            strDefaultSecrecyLevel = COMMON_SETDEF_SECLV;

            // デフォルト言語コード初期化
            strCulture = System.Threading.Thread.CurrentThread.CurrentUICulture.ToString();

            // デフォルト文書のサーバーパス初期化
            strSABListServerPath = COMMON_SETDEF_SABLISTSERVERPATH;

            // デフォルト文書のローカルパス初期化
            strSABListLocalPath = COMMON_SETDEF_SABLISTLOCALPATH;

            // デフォルトzip一時解凍先初期化
            strTempPath = COMMON_SETDEF_TEMPPATH;

            // デフォルトセキュアフォルダリスト初期化
            lstSecureFolder = new List<string>() { COMMON_SETDEF_SECUREFOLDER_1, COMMON_SETDEF_SECUREFOLDER_2 };

            // デフォルト「最終版」を表す文字列初期化
            lstFinal = new List<string>() { COMMON_SETDEF_FINAL_1, COMMON_SETDEF_FINAL_2 };

            // デフォルト関連会社ファイルパス(サーバー)初期化
            strGroupDomainListServerPath = COMMON_SETDEF_GROUPDOMAINLISTSERVERPATH;

            // デフォルト関連会社ファイルパス(ローカル)初期化
            strGroupDomainListLocalPath = COMMON_SETDEF_GROUPDOMAINLISTLOCALPATH;

            // デフォルト役職者ファイルパス(サーバー)初期化
            strManagerListServerPath = COMMON_SETDEF_MANAGERLISTSERVERPATH;

            // デフォルト役職者ファイルパス(ローカル)初期化
            strManagerListLocalPath = COMMON_SETDEF_MANAGERLISTLOCALPATH;

            // デフォルトzip化強制レベル初期化
            zipLevel = COMMON_SETDEF_ZIPLEVEL;
        }

        #endregion
    }

    /// <summary>
    /// 共通設定ファイルの読み込み処理
    /// </summary>
    public class CommonSettingRead
    {
        /// <summary>
        /// 共通設定クラス
        /// </summary>
        public CommonSettings commonSettings;

        #region 共通設定ファイル読み込み

        /// <summary>
        /// 共通設定ファイル読み込み
        /// </summary>
        /// <param name=""></param>
        /// <returns>true:読み込み成功、false:読み込み失敗</returns>
        public CommonSettings Reader()
        {
            commonSettings = new CommonSettings();

            try
            {
                CommonSettingStoring commonSettingStoring = new CommonSettingStoring();
                // 共通設定ファイルパス作成
                string strCommonSettingFilePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    commonSettingStoring.COMMON_SETFOLDERNAME,
                    commonSettingStoring.COMMON_SETFILENAME
                    );

                // 共通設定ファイルが存在しない場合はデフォルト設定を書き込む
                if (File.Exists(strCommonSettingFilePath) == false)
                {
                    if (!CommonSettingWrite())
                    {
                        // エラーメッセージダイアログは各々の呼び出し先で定義
                        return null;
                    }
                }

                //XmlSerializerオブジェクトの作成
                System.Xml.Serialization.XmlSerializer serXmlCommonRead = new System.Xml.Serialization.XmlSerializer(typeof(CommonSettings));

                //ファイルを開く
                StreamReader stmCommonReader = new StreamReader(strCommonSettingFilePath, Encoding.GetEncoding("shift_jis"));

                //XMLファイルから読み込み、逆シリアル化する
                commonSettings = (CommonSettings)serXmlCommonRead.Deserialize(stmCommonReader);

                //閉じる
                stmCommonReader.Close();

                return commonSettings;
            }
            catch (Exception ex)
            {
                // 読み込み or 書き込みの失敗
                return null;
            }
        }

        #endregion

        #region 共通設定設定書き込み

        /// <summary>
        /// 共通設定設定書き込み
        /// </summary>
        /// <returns>true:書込み成功、false:書込み失敗</returns>
        private Boolean CommonSettingWrite()
        {
            try
            {
                CommonSettingStoring commonSettingStoring = new CommonSettingStoring();
                // 共通設定ファイルパス作成
                string strCommonSettingFilePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    commonSettingStoring.COMMON_SETFOLDERNAME
                    );

                if (Directory.Exists(strCommonSettingFilePath) == false)
                {
                    // フォルダ作成
                    System.IO.Directory.CreateDirectory(strCommonSettingFilePath);
                }

                strCommonSettingFilePath = Path.Combine(strCommonSettingFilePath, commonSettingStoring.COMMON_SETFILENAME);

                //XmlSerializerオブジェクトの作成
                System.Xml.Serialization.XmlSerializer serXmlCommonWrite = new System.Xml.Serialization.XmlSerializer(typeof(CommonSettings));

                //ファイルを開く
                System.IO.StreamWriter stmCommonWrite = new System.IO.StreamWriter(strCommonSettingFilePath, false, Encoding.GetEncoding("shift_jis"));

                //シリアル化し、XMLファイルに保存する
                serXmlCommonWrite.Serialize(stmCommonWrite, commonSettings);

                //閉じる
                stmCommonWrite.Close();

                return true;
            }
            catch (Exception ex)
            {
                // 共通設定ファイルの新規作成に失敗
                return false;
            }
        }

        #endregion

    }
}
