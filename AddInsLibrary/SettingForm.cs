using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Globalization;
using System.Threading;

namespace AddInsLibrary
{
    public partial class SettingForm : Form
    {
        #region <定数定義>
        /// <summary>
        /// 定数定義
        /// </summary>

        /// <summary>
        /// SAB区分
        /// 
        /// プロパティに書き込むSAB秘 S秘
        /// </summary>
        public const string SECRECY_PROPERTY_S    = "SecrecyS";

        /// <summary>
        /// SAB区分
        /// 
        /// プロパティに書き込むSAB秘 A秘
        /// </summary>
        public const string SECRECY_PROPERTY_A    = "SecrecyA";

        /// <summary>
        /// SAB区分
        /// 
        /// プロパティに書き込むSAB秘 B秘
        /// </summary>
        public const string SECRECY_PROPERTY_B    = "SecrecyB";

        /// <summary>
        /// SAB区分
        /// 
        /// プロパティに書き込むSAB秘 以外
        /// </summary>
        public const string SECRECY_PROPERTY_ELSE = "SecrecyNone";

        /// <summary>
        /// 共通設定関連
        /// 
        /// 共通設定 格納フォルダ名
        /// </summary>
        public const string COMMON_SETFOLDERNAME = "SAB";

        /// <summary>
        /// 共通設定関連
        /// 
        /// 共通設定 ファイル名 デフォルト機密区分
        /// </summary>
        public const string COMMON_SETFILENAME   = "common_setting.config";

        /// <summary>
        /// スタンプ倍率
        /// </summary>
        public const double STAMP_MAGNIFICATION = 1.3331;


        /// <summary>
        /// スタンプを識別するための文字列
        /// </summary>
        protected string STAMP_SHAPE_NAME = "HONDA_SECRECY_STAMP";

        /// <summary>
        /// プロパティのタグ情報位置
        /// </summary>
        protected enum Property
        {
            SecrecyLevel = 0, // 機密区分
            ClassNo,          // 文書番号
            OfficeCode,       // 事業所コード
        }

        /// <summary>
        /// SAB機密区分
        /// </summary>
        protected enum Secrecy
        {
            S,   // 機密区分
            A,   // 文書番号
            B,   // 事業所コード
            None // 事業所コード
        }
        #endregion


        #region <内部変数>

        /// <summary>
        /// 共通設定クラス
        /// </summary>
        public CommonSettings clsCommonSettting;

        /// <summary>
        /// 共通設定エラーフラグ
        /// </summary>
        public Boolean commonFileReadCompleted;

        #endregion

        #region <クラス定義>

        // 2020/09/17 共通設定ファイルを１つにまとめるため、AddInsLibrary内に専用のクラスを作成
        // XMLファイルのタグ名とクラス名が一致していないと読み込み失敗になるので、ここのCommonSettingsクラスは使用しないようにする

        #endregion


        #region <コンストラクタ>
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SettingForm()
        {
            try
            {
                CommonSettingRead read = new CommonSettingRead();
                clsCommonSettting = read.Reader();

                if (clsCommonSettting == null)
                {
                    MessageBox.Show(Properties.Resources.msg_err_read_common_file,
                        Properties.Resources.msgError,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Hand);

                    return;
                }

                // 共通設定ファイルに不足項目がないかチェック
                bool CommonSetttingFlg = true;
                string CommonSetttingMessage = null;
                CommonSetttingMessage = Properties.Resources.msgCommonSettingError + Environment.NewLine;
                if (string.IsNullOrEmpty(clsCommonSettting.strDefaultSecrecyLevel))   // デフォルト機密区分
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Properties.Resources.msgDefaultSecure + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(clsCommonSettting.strOfficeCode))   // 事業所コード
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Properties.Resources.msgOfficeCode + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(clsCommonSettting.strCulture))   // 言語設定
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Properties.Resources.msgSettingLanguage + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(clsCommonSettting.strSABListLocalPath))   // 文書のローカルパス
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Properties.Resources.msgLocalPath + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(clsCommonSettting.strSABListServerPath))   // 文書のサーバーパス
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Properties.Resources.msgServerPath + Environment.NewLine;
                }
                if (string.IsNullOrEmpty(clsCommonSettting.strTempPath))   // zip一時解凍先
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Properties.Resources.msgTempZipPath + Environment.NewLine;
                }
                if (clsCommonSettting.lstSecureFolder.Count == 0)   // セキュアフォルダリスト
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Properties.Resources.msgFolderList + Environment.NewLine;
                }
                if (clsCommonSettting.lstFinal.Count == 0)   // 「最終版」を表す文字列
                {
                    CommonSetttingFlg = false;
                    CommonSetttingMessage += Properties.Resources.msgFinal + Environment.NewLine;
                }

                if (!CommonSetttingFlg)
                {
                    // 共通設定ファイルに不足項目あり
                    MessageBox.Show(CommonSetttingMessage,
                        Properties.Resources.msgError,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Hand);

                    // 強制終了
                    Environment.Exit(0x8020);
                }
            }
            catch (Exception ex)
            {
                // 共通設定ファイルが読み込まれていないのでエラーになる
                return;
            }

            commonFileReadCompleted = true;

            // 各コンポーネント初期化
            InitializeComponent();

            // 言語設定を表示
            this.lblLanguage.Text = clsCommonSettting.strCulture;

#if DEBUG
            this.lblLanguage.Visible = true;
#endif
        }
        #endregion


        #region <フォームイベント>
        /// <summary>
        /// フォームロードイベント
        /// </summary>
        protected void FormSetting_Load(object sender, EventArgs e)
        {
            string filePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
            string filePropertyClassNo = string.Empty;      // ファイルプロパティ情報 事業所コード
            string filePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

            // ファイルプロパティ情報取得
            this.GetDocumentProperty(ref filePropertySecrecyLevel, ref filePropertyClassNo, ref filePropertyOfficeCode);

            // プロパティファイルにSAB機密区分が設定されていない場合は標準値を設定
            if(string.IsNullOrWhiteSpace(filePropertySecrecyLevel))
            {
                filePropertySecrecyLevel = clsCommonSettting.strDefaultSecrecyLevel;
            }

            // ラジオボタンをセット
            this.SetSABRadioButton(filePropertySecrecyLevel);

            // 文章の更新処理
            SetSABList();
        }

        /// <summary>
        /// 文書のサーバーパスを参照し、文書のローカルパスを更新する処理
        /// </summary>
        protected void SetSABList()
        {
            // ネットワークに接続されているか
            if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
            {
                string _serverFilePath = clsCommonSettting.strSABListServerPath;
                string _localFilePath = clsCommonSettting.strSABListLocalPath;

                try
                {
                    // 設定値のチェック
                    if (_serverFilePath == "" || _localFilePath == "")
                    {
                        MessageBox.Show(Properties.Resources.msgFailedReadDocumentPath, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);  // step2 iwasa
                        return;
                    }

                    // ファイルの有無
                    if (File.Exists(_serverFilePath) == true)
                    {
                        // ローカルの保存先があるか
                        string _localDir = Path.GetDirectoryName(_localFilePath);

                        if (Directory.Exists(_localDir) == false)
                        {
                            // フォルダ作成
                            Directory.CreateDirectory(_localDir);
                        }

                        // サーバとローカルのファイルを比較する
                        DateTime _serverDocument = File.GetLastWriteTime(_serverFilePath);
                        DateTime _localDocument = File.GetLastWriteTime(_localFilePath);

                        if (_serverDocument > _localDocument)
                        {
                            // ローカルファイルを更新する
                            File.Copy(_serverFilePath, _localFilePath, true);
                        }
                    }
                }
                catch
                {
                    // ファイル移動関連でエラーが発生する可能性があるが
                    // 処理として問題ないためスルーする
                }
            }
        }

        /// <summary>
        /// フォームキーダウン処理
        /// </summary>
        protected void FormSetting_KeyDown(object sender, KeyEventArgs e)
        {
            // ESCキーが押された場合
            if (e.KeyData == Keys.Escape)
            {
                // 後で登録ボタンクリック
                buttonNotRegist_Click(sender, e);
            }
            // Enterキーが押された場合
            else if (e.KeyData == Keys.Enter)
            {
                // 登録ボタンクリック処理
                btnRegist_Click(sender, e);
            }
        }

        /// <summary>
        /// 後で登録するボタンのクリックイベント
        /// </summary>
        protected void buttonNotRegist_Click(object sender, EventArgs e)
        {
            // フォームを閉じる
            this.Close();
        }

        /// <summary>
        /// ラジオボタンを変更したときのイベント
        /// </summary>
        protected void btnSAB_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                // SAB区分のテキスト変更
                Secrecy selectedSecrecy = this.GetSelectedSecrecyLevel();
                this.lblSABSetting.Text = this.GetSABText(selectedSecrecy);

                // SAB区分ラジオボタン 背景色変更
                this.ChangeBackColorSAB(ref this.rdoS);
                this.ChangeBackColorSAB(ref this.rdoA);
                this.ChangeBackColorSAB(ref this.rdoB);
                this.ChangeBackColorSAB(ref this.rdoElse);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Properties.Resources.msgSecrecyButtonError,
                    Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// 閉じるボタンをクリックしたときのイベント
        /// </summary>
        protected void btnClose_Click(object sender, EventArgs e)
        {
            // フォームを閉じる
            this.Close();
        }

        /// <summary>
        /// 表示切替ボタン変更時のイベント
        /// </summary>
        protected void btnChange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                string textStampON = Properties.Resources.StampDisplay_ON;
                string textStampOFF = Properties.Resources.StampDisplay_OFF;
                lblDisplay.Text = chkChange.Checked == true ? textStampON : textStampOFF;
            }
            catch (Exception ex)
            {
                MessageBox.Show(Properties.Resources.msgViewChangeError,
                    Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// フォームを閉じるときのイベント
        /// </summary>
        protected void FormSetting_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                bool canClose = this.CanClose();
                e.Cancel = canClose;
            }
            catch (Exception ex)
            {
                MessageBox.Show(Properties.Resources.msgCloseButtonError,
                    Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// 登録ボタンをクリックしたときのイベント
        /// </summary>
        protected void btnRegist_Click(object sender, EventArgs e)
        {
            Console.WriteLine(System.Threading.Thread.CurrentThread.CurrentUICulture);

            try
            {
                string strFilePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
                string strFilePropertyClassNo = string.Empty;      // ファイルプロパティ情報 事業所コード
                string strFilePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード
                this.GetDocumentProperty(ref strFilePropertySecrecyLevel, ref strFilePropertyClassNo, ref strFilePropertyOfficeCode); // プロパティ情報取得

                // プロパティ情報存在チェック
                if (!string.IsNullOrEmpty(strFilePropertySecrecyLevel))
                {
                    if (MessageBox.Show(Properties.Resources.msg_qst_stamp_orverwrite, AddInsLibrary.Properties.Resources.msgConfirm, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return;
                    }
                }

                // スタンプ貼付け処理
                Secrecy selectedSecrecyLevel = GetSelectedSecrecyLevel();
                bool setResultIsOK = this.SetStamp(selectedSecrecyLevel);

                if (setResultIsOK == false)
                {
                    MessageBox.Show(Properties.Resources.msg_err_stamp_paste, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }


                // ファイルプロパティセット
                Secrecy selectedSecrecy = this.GetSelectedSecrecyLevel();
                string SABCode = this.GetSABCode(selectedSecrecy);
                bool documentPropertyResult = SetDocumentProperty(SABCode);

                if (documentPropertyResult == false)
                {
                    MessageBox.Show(Properties.Resources.msg_err_properties_write, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }


                // ダイアログを閉じる
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(Properties.Resources.msgRegistButtonError,
                    Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        /// <summary>
        /// リンクをクリックしたときのイベント
        /// </summary>
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // リンク先に移動したことにする
            lblLink.LinkVisited = true;

            // 開きたいファイルのパス
            string Folderpath = null;

            try
            {
                Folderpath = clsCommonSettting.strSABListLocalPath;

                if (File.Exists(Folderpath) == true)
                {
                    Process.Start(Folderpath);
                }
                else
                {
                    MessageBox.Show(AddInsLibrary.Properties.Resources.msgGCPListPathCheck,
                        AddInsLibrary.Properties.Resources.msgError,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Hand);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(AddInsLibrary.Properties.Resources.msgGCPListError,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);
            }
        }

        #endregion


        #region <メソッド>
        /// <summary>
        /// 機密区分が登録済みか
        /// </summary>
        /// <returns>true:登録済</returns>
        public bool IsSecrecyInfoRegistered()
        {
            string strFilePropertySecrecyLevel = string.Empty; // ファイルプロパティ情報 機密区分
            string strFilePropertyClassNo = string.Empty;      // ファイルプロパティ情報 文書番号
            string strFilePropertyOfficeCode = string.Empty;   // ファイルプロパティ情報 事業所コード

            // プロパティ情報取得
            this.GetDocumentProperty(ref strFilePropertySecrecyLevel, ref strFilePropertyClassNo, ref strFilePropertyOfficeCode); 


            // 機密区分が空白、または事業所コードが自事業所コードではない
            if (string.IsNullOrEmpty(strFilePropertySecrecyLevel) ||
                strFilePropertyOfficeCode != this.clsCommonSettting.strOfficeCode )
            {
                return false;
            }

            // 機密区分のいずれにも該当しない
            if (SECRECY_PROPERTY_S    != strFilePropertySecrecyLevel &&
                SECRECY_PROPERTY_A    != strFilePropertySecrecyLevel &&
                SECRECY_PROPERTY_B    != strFilePropertySecrecyLevel &&
                SECRECY_PROPERTY_ELSE != strFilePropertySecrecyLevel )
            {
                return false;
            }

            return true;
        }


        private bool mustRegistMode = false;
        /// <summary>
        /// 機密区分必須登録モード切替え
        /// </summary>
        /// <returns>true:登録必須モードON、false:登録必須モードOFF</returns>
        public bool MustRegistMode
        {
            set
            {
                this.mustRegistMode = value;

                // 閉じる機能を切替え
                this.btnClose.Enabled = !value;
                this.ControlBox = !value;
            }
            get
            {
                return this.mustRegistMode;
            }
        }

        /// <summary>
        /// ラジオボタンの選択状態から背景色を更新
        /// </summary>
        /// <param name="radioSAB">対象のラジオボタン</param>
        private void ChangeBackColorSAB(ref RadioButton radioSAB)
        {
            if (radioSAB.Checked == true)
            {
                radioSAB.BackColor = Color.Green;
            }
            else
            {
                radioSAB.BackColor = Color.Gray;
            }
        }

        /// <summary>
        /// SAB機密区分コードからラジオボタンを選択
        /// </summary>
        /// <param name="secrecyCode">SAB機密区分コード</param>
        private void SetSABRadioButton(string secrecyCode)
        {
            if (SECRECY_PROPERTY_S.Equals(secrecyCode))
            {
                this.rdoS.Checked = true;

                return;
            }

            if (SECRECY_PROPERTY_A.Equals(secrecyCode))
            {
                this.rdoA.Checked = true;

                return;
            }

            if (SECRECY_PROPERTY_B.Equals(secrecyCode))
            {
                this.rdoB.Checked = true;

                return;
            }

            this.rdoElse.Checked = true;
        }

        /// <summary>
        /// ラジオボタンの選択状態からSAB機密区分を取得
        /// </summary>
        /// <returns>選択中のSAB機密区分 列挙体</returns>
        private Secrecy GetSelectedSecrecyLevel()
        {
            if (this.rdoS.Checked == true)
            {
                return Secrecy.S;
            }

            if (this.rdoA.Checked == true)
            {
                return Secrecy.A;
            }

            if (this.rdoB.Checked == true)
            {
                return Secrecy.B;
            }

            return Secrecy.None;
        }

        /// <summary>
        /// SAB機密区分列挙体からSAB機密区分のテキストを取得
        /// </summary>
        private string GetSABText(Secrecy secrecy)
        {
            if (secrecy == Secrecy.S)
            {
                return this.rdoS.Text;
            }

            if (secrecy == Secrecy.A)
            {
                return this.rdoA.Text;
            }

            if (secrecy == Secrecy.B)
            {
                return this.rdoB.Text;
            }

            return this.rdoElse.Text; ;
        }

        /// <summary>
        /// ラジオボタンの選択状態からSAB機密区分コードを取得
        /// </summary>
        /// <param name="secrecy"></param>
        /// <returns></returns>
        private string GetSABCode(Secrecy secrecy)
        {
            if (secrecy == Secrecy.S)
            {
                return SECRECY_PROPERTY_S;
            }

            if (secrecy == Secrecy.A)
            {
                return SECRECY_PROPERTY_A;
            }

            if (secrecy == Secrecy.B)
            {
                return SECRECY_PROPERTY_B;
            }

            return SECRECY_PROPERTY_ELSE;
        }

        /// <summary>
        /// SAB機密区分列挙体に対応するスタンプ画像を取得
        /// </summary>
        /// <param name="secrecy">SAB機密区分列挙体</param>
        /// <returns>スタンプ画像</returns>
        protected Bitmap GetStampImage(Secrecy secrecy)
        {
            if (secrecy == Secrecy.S)
            {
                return Properties.Resources.StampS;
            }

            if (secrecy == Secrecy.A)
            {
                return Properties.Resources.StampA;
            }

            if (secrecy == Secrecy.B)
            {
                return Properties.Resources.StampB;
            }

            return null;
        }

        /// <summary>
        /// SAB機密区分の登録状態から閉じれる状態か確認
        /// </summary>
        /// <returns>true:フォーム閉じる不可</returns>
        private bool CanClose()
        {
            bool closingCancel = false;

            // 登録必須の状態ではない場合は閉じて良い
            if (this.MustRegistMode == false)
            {
                return closingCancel;
            }

            // 機密区分が登録されている場合は閉じて良い
            if (this.IsSecrecyInfoRegistered() == true)
            {
                return closingCancel;
            }

            // 閉じて良い条件に該当しなかった場合は閉じない
            closingCancel = true;


            return closingCancel;
        }

        /// <summary>
        /// 画像透過処理
        /// </summary>
        /// <param name="bitmap"></param>
        /// <returns>透過済み画像</returns>
        protected Bitmap CreateAlphaImage(Bitmap bitmap, float alpha)
        {
            int imageWidth = bitmap.Width;
            int imageHeight = bitmap.Height;

            // 新しいビットマップを用意
            Bitmap alphaImage = new Bitmap(imageWidth, imageHeight);

            using (Graphics graphics = Graphics.FromImage(alphaImage))
            {
                // ColorMatrixオブジェクトの作成
                System.Drawing.Imaging.ColorMatrix cm = 
                    new System.Drawing.Imaging.ColorMatrix();

                // ColorMatrixの行列の値を変更して、アルファ値が0.5に変更されるようにする
                cm.Matrix00 = 1;
                cm.Matrix11 = 1;
                cm.Matrix22 = 1;
                cm.Matrix33 = (1f - alpha);
                cm.Matrix44 = 1;

                // ImageAttributesオブジェクトの作成
                System.Drawing.Imaging.ImageAttributes imageAttributes = 
                    new System.Drawing.Imaging.ImageAttributes();

                // ColorMatrixを設定
                imageAttributes.SetColorMatrix(cm);

                // ImageAttributesを使用して画像を描画
                graphics.DrawImage(bitmap,
                            new Rectangle(0, 0, imageWidth, imageHeight),
                            0,
                            0,
                            imageWidth,
                            imageHeight,
                            GraphicsUnit.Pixel,
                            imageAttributes);
            }

            return alphaImage;
        }
        #endregion


        #region <Overrideメソッド>

        /// <summary>
        /// スタンプ貼付け処理
        /// </summary>
        protected virtual Boolean SetStamp(Secrecy secrecyLevel)
        {
            // Excel・Word・PowerPointのスタンプ貼付け処理を子クラスで実装

            return true;
        }

        /// <summary>
        /// ドキュメントのプロパティ取得
        /// </summary>
        /// <param name="strClassNo">文書分類番号</param>
        /// <param name="strSecrecyLevel">機密区分</param>
        /// <param name="bStamp">スタンプ有無</param>off
        public virtual void GetDocumentProperty(ref string strSecrecyLevel, ref string strOfficeCod, ref string strOfficeCode)
        {
            // Excel・Word・PowerPointのプロパティ取得処理を子クラスで実装
        }


        /// <summary>
        /// ドキュメントのプロパティ設定
        /// オーバーライド用のメソッド
        /// </summary>
        /// <param name="strClassNo"></param>
        /// <param name="strSecrecyLevel"></param>
        public virtual bool SetDocumentProperty(string strSecrecyLevel)
        {
            // Excel・Word・PowerPointのプロパティ設定処理を子クラスで実装

            return true;
        }

        #endregion
    }
}