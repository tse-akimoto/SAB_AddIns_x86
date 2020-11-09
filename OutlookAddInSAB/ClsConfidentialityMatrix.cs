using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddInSAB
{
    public class ClsConfidentialityMatrix
    {
        /// <summary>
        /// 機密区分の設定値 登録なし
        /// </summary>
        const string SECRECY_NONE_RANK = "0/";

        /// <summary>
        /// 機密区分の設定値 S秘
        /// </summary>
        const string SECRECY_S_RANK = "1/";

        /// <summary>
        /// 機密区分の設定値 A秘
        /// </summary>
        const string SECRECY_A_RANK = "2/";

        /// <summary>
        /// 機密区分の設定値 B秘
        /// </summary>
        const string SECRECY_B_RANK = "3/";

        /// <summary>
        /// 機密区分の設定値 以外
        /// </summary>
        const string SECRECY_OTHER_RANK = "4/";

        /// <summary>
        /// 送信者役職区分の設定値 役員
        /// </summary>
        const string EXECUTIVE = "executive/";

        /// <summary>
        /// 送信者役職区分の設定値 部門長、所属長
        /// </summary>
        const string MANAGER = "manager/";

        /// <summary>
        /// 送信者役職区分の設定値 一般
        /// </summary>
        const string NOMAL = "nomal/";

        /// <summary>
        /// 社内
        /// </summary>
        const string TRUE = "True";

        /// <summary>
        /// 社外
        /// </summary>
        const string FALSE = "False";

        /// <summary>
        /// 共通設定によるZIPパスワード設定有無
        /// </summary>
        public int configZipLevel { get; set; }

        /// <summary>
        /// 結果
        /// </summary>
        public Dictionary<string, Dictionary<string, Dictionary<string, SendPattern>>> ResultMatrix = new Dictionary<string, Dictionary<string, Dictionary<string, SendPattern>>>();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ClsConfidentialityMatrix()
        {

        }

        /// <summary>
        /// 初期化
        /// </summary>
        public void Initialize()
        {
            bool bZipLevel = false;
            if (configZipLevel == 1)
            {
                bZipLevel = true;
            }

            #region リスト

            Dictionary<string, Dictionary<string, SendPattern>> NoneDictionary = new Dictionary<string, Dictionary<string, SendPattern>>();
            Dictionary<string, SendPattern> NoneExecutivePattern = new Dictionary<string, SendPattern>();
            Dictionary<string, SendPattern> NoneManagerPattern = new Dictionary<string, SendPattern>();
            Dictionary<string, SendPattern> NoneNomalPattern = new Dictionary<string, SendPattern>();

            Dictionary<string, Dictionary<string, SendPattern>> SDictionary = new Dictionary<string, Dictionary<string, SendPattern>>();
            Dictionary<string, SendPattern> SExecutivePattern = new Dictionary<string, SendPattern>();
            Dictionary<string, SendPattern> SManagerPattern = new Dictionary<string, SendPattern>();
            Dictionary<string, SendPattern> SNomalPattern = new Dictionary<string, SendPattern>();

            Dictionary<string, Dictionary<string, SendPattern>> ABDictionary = new Dictionary<string, Dictionary<string, SendPattern>>();
            Dictionary<string, SendPattern> ABExecutivePattern = new Dictionary<string, SendPattern>();
            Dictionary<string, SendPattern> ABManagerPattern = new Dictionary<string, SendPattern>();
            Dictionary<string, SendPattern> ABNomalPattern = new Dictionary<string, SendPattern>();

            Dictionary<string, Dictionary<string, SendPattern>> OtherDictionary = new Dictionary<string, Dictionary<string, SendPattern>>();
            Dictionary<string, SendPattern> OtherExecutivePattern = new Dictionary<string, SendPattern>();
            Dictionary<string, SendPattern> OtherManagerPattern = new Dictionary<string, SendPattern>();
            Dictionary<string, SendPattern> OtherNomalPattern = new Dictionary<string, SendPattern>();

            #endregion

            #region 登録なし

            // 役員
            NoneExecutivePattern[TRUE] = new SendPattern(true, false, false);
            NoneExecutivePattern[FALSE] = new SendPattern(true, false, false);
            NoneDictionary[EXECUTIVE] = NoneExecutivePattern;

            // 部門長/所属長
            NoneManagerPattern[TRUE] = new SendPattern(false, false, false);
            NoneManagerPattern[FALSE] = new SendPattern(false, false, false);
            NoneDictionary[MANAGER] = NoneManagerPattern;

            // 一般
            NoneNomalPattern[TRUE] = new SendPattern(false, false, false);
            NoneNomalPattern[FALSE] = new SendPattern(false, false, false);
            NoneDictionary[NOMAL] = NoneNomalPattern;

            #endregion

            #region S秘

            // 役員
            SExecutivePattern[TRUE] = new SendPattern(true, false, true);
            SExecutivePattern[FALSE] = new SendPattern(true, false, true);
            SDictionary[EXECUTIVE] = SExecutivePattern;

            // 部門長/所属長
            SManagerPattern[TRUE] = new SendPattern(true, false, true);
            SManagerPattern[FALSE] = new SendPattern(false, false, false);
            SDictionary[MANAGER] = SManagerPattern;

            // 一般
            SNomalPattern[TRUE] = new SendPattern(true, true, true);
            SNomalPattern[FALSE] = new SendPattern(false, false, false);
            SDictionary[NOMAL] = SNomalPattern;

            #endregion

            #region A秘/B秘

            // 役員
            ABExecutivePattern[TRUE] = new SendPattern(true, false, bZipLevel);
            ABExecutivePattern[FALSE] = new SendPattern(true, false, true);
            ABDictionary[EXECUTIVE] = ABExecutivePattern;

            // 部門長/所属長
            ABManagerPattern[TRUE] = new SendPattern(true, false, bZipLevel);
            ABManagerPattern[FALSE] = new SendPattern(true, false, true);
            ABDictionary[MANAGER] = ABManagerPattern;

            // 一般
            ABNomalPattern[TRUE] = new SendPattern(true, false, bZipLevel);
            ABNomalPattern[FALSE] = new SendPattern(true, true, true);
            ABDictionary[NOMAL] = ABNomalPattern;

            #endregion

            #region 以外

            // 役員
            OtherExecutivePattern[TRUE] = new SendPattern(true, false, false);
            OtherExecutivePattern[FALSE] = new SendPattern(true, false, false);
            OtherDictionary[EXECUTIVE] = OtherExecutivePattern;

            // 部門長/所属長
            OtherManagerPattern[TRUE] = new SendPattern(true, false, false);
            OtherManagerPattern[FALSE] = new SendPattern(true, false, false);
            OtherDictionary[MANAGER] = OtherManagerPattern;

            // 一般
            OtherNomalPattern[TRUE] = new SendPattern(true, false, false);
            OtherNomalPattern[FALSE] = new SendPattern(true, false, false);
            OtherDictionary[NOMAL] = OtherNomalPattern;

            #endregion

            // 結果格納
            ResultMatrix[SECRECY_NONE_RANK] = NoneDictionary;
            ResultMatrix[SECRECY_S_RANK] = SDictionary;
            ResultMatrix[SECRECY_A_RANK] = ABDictionary;
            ResultMatrix[SECRECY_B_RANK] = ABDictionary;
            ResultMatrix[SECRECY_OTHER_RANK] = OtherDictionary;
        }
    }

    public class SendPattern
    {
        /// <summary>
        /// 送信可否 true: 可, false 不可
        /// </summary>
        public bool bSend = false;

        /// <summary>
        /// 上長の許可 true: 必要 false: 不要
        /// </summary>
        public bool bSuperiorPermission = false;

        /// <summary>
        /// 添付ファイル圧縮・パスワード化 ture: 必要 false:任意
        /// </summary>
        public bool bZipCompression = false;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SendPattern()
        {

        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SendPattern(bool Send, bool SuperiorPermission, bool ZipCompression)
        {
            bSend = Send;
            bSuperiorPermission = SuperiorPermission;
            bZipCompression = ZipCompression;
        }
    }
}
