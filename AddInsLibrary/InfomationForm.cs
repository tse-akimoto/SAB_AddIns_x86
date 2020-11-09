using System;
using System.Windows.Forms;

namespace AddInsLibrary
{
    public partial class InfomationForm : Form
    {
        #region コンストラクタ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="secrecyCode">SAB機密区分コード</param>
        public InfomationForm(string secrecyCode)
        {
            InitializeComponent();

            string secrecyLabelText = GetSecrecyLabelText(secrecyCode);

            this.lblSABSetting.Text = secrecyLabelText;
        }

        #endregion


        #region フィールド

        /// <summary>
        /// SAB機密区分コード
        /// 
        /// プロパティに書き込むSAB秘 S秘
        /// </summary>
        public const string SECRECY_PROPERTY_S    = "SecrecyS";

        /// <summary>
        /// SAB機密区分コード
        /// 
        /// プロパティに書き込むSAB秘 A秘
        /// </summary>
        public const string SECRECY_PROPERTY_A    = "SecrecyA";

        /// <summary>
        /// SAB機密区分コード
        /// 
        /// プロパティに書き込むSAB秘 B秘
        /// </summary>
        public const string SECRECY_PROPERTY_B    = "SecrecyB";

        /// <summary>
        /// SAB機密区分コード
        /// 
        /// プロパティに書き込むSAB秘 以外
        /// </summary>
        public const string SECRECY_PROPERTY_ELSE = "SecrecyNone";

        #endregion


        #region イベントハンドラ
        /// <summary>
        /// 登録ボタンのクリックイベント
        /// </summary>
        private void btnRegist_Click(object sender, EventArgs e)
        {
            // DialogResultで結果を取得
        }

        /// <summary>
        /// 閉じるボタンのクリックイベント
        /// </summary>
        private void btnClose_Click(object sender, EventArgs e)
        {
            // DialogResult=Cancel
        }
        #endregion


        #region メソッド

        /// <summary>
        /// 
        /// </summary>
        /// <param name="secrecyCode">SAB機密区分コード</param>
        /// <returns></returns>
        private string GetSecrecyLabelText(string secrecyCode)
        {
            if (SECRECY_PROPERTY_S == secrecyCode)
            {
                return Properties.Resources.txt_SecrecyS;
            }

            if (SECRECY_PROPERTY_A == secrecyCode)
            {
                return Properties.Resources.txt_SecrecyA;
            }

            if (SECRECY_PROPERTY_B == secrecyCode)
            {
                return Properties.Resources.txt_SecrecyB;
            }

            return Properties.Resources.txt_SecrecyNone;
        }

        #endregion
    }
}
