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
        public const string COMMON_SETDEF_SABLISTSERVERPATH = "C:\\ProgramData\\SAB\\GCPList.xslx";   // TODO

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 文書のローカルパス
        /// </summary>
        public const string COMMON_SETDEF_SABLISTLOCALPATH = "C:\\ProgramData\\SAB\\GCPList.xslx";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// zip一時解凍先
        /// </summary>
        public const string COMMON_SETDEF_TEMPPATH = "C:\\ProgramData\\SAB\\zipTempPath";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// セキュアフォルダリスト
        /// </summary>
        public const string COMMON_SETDEF_SECUREFOLDER = "C:\\tmp\\SecureA";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 「最終版」を表す文字列
        /// </summary>
        public const string COMMON_SETDEF_FINAL = "最終版";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 関連会社ファイルパス(サーバー)
        /// </summary>
        public const string COMMON_SETDEF_GROUPDOMAINLISTSERVERPATH = "C:\\ProgramData\\SAB\\DomainList.txt";   // TODO

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 関連会社ファイルパス(ローカル)
        /// </summary>
        public const string COMMON_SETDEF_GROUPDOMAINLISTLOCALPATH = "C:\\ProgramData\\SAB\\DomainList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 役職者ファイルパス(サーバー)
        /// </summary>
        public const string COMMON_SETDEF_MANAGERLISTSERVERPATH = "C:\\ProgramData\\SAB\\ManagerList.txt";   // TODO

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// 役職者ファイルパス(ローカル)
        /// </summary>
        public const string COMMON_SETDEF_MANAGERLISTLOCALPATH = "C:\\ProgramData\\SAB\\ManagerList.txt";

        /// <summary>
        /// 共通設定 デフォルト値
        /// 
        /// zip化強制レベル
        /// </summary>
        public const int COMMON_SETDEF_ZIPLEVEL = 1;

        #endregion

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
            lstSecureFolder = new List<string>() { COMMON_SETDEF_SECUREFOLDER };

            // デフォルト「最終版」を表す文字列初期化
            lstFinal = new List<string>() { COMMON_SETDEF_FINAL };

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

        /// <summary>
        /// 共通設定ファイル読み込み
        /// </summary>
        /// <param name=""></param>
        /// <returns>true:読み込み成功、false:読み込み失敗</returns>
        public CommonSettings Reader()
        {
            commonSettings = new CommonSettings();

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
                    string msgErrReadCommonFile = Properties.Resources.msg_err_read_common_file;
                    MessageBox.Show(msgErrReadCommonFile);
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
                return false;
            }
        }
    }
}
