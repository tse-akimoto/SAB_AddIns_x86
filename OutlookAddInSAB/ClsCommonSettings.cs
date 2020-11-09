using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace AddInsLibrary
{
    /// <summary>
    /// ���ʐݒ�t�@�C���̊i�[��
    /// </summary>
    public class CommonSettingStoring
    {
        /// <summary>
        /// ���ʐݒ� �i�[�t�H���_��
        /// </summary>
        public string COMMON_SETFOLDERNAME = "SAB";

        /// <summary>
        /// ���ʐݒ� �t�@�C����
        /// </summary>
        public string COMMON_SETFILENAME = "common_setting.config";
    }

    /// <summary>
    /// ���ʐݒ�t�@�C���̐ݒ�l
    /// </summary>
    public class CommonSettings
    {
        /// <summary>
        /// ���Ə��R�[�h
        /// </summary>
        public string strOfficeCode { get; set; }

        /// <summary>
        /// �@���敪
        /// </summary>
        public string strDefaultSecrecyLevel { get; set; }

        /// <summary>
        /// ����ݒ�
        /// </summary>
        public string strCulture { get; set; }

        /// <summary>
        /// �����̃T�[�o�[�p�X
        /// </summary>
        public string strSABListServerPath { get; set; }

        /// <summary>
        /// �����̃��[�J���p�X
        /// </summary>
        public string strSABListLocalPath { get; set; }

        /// <summary>
        /// zip�ꎞ�𓀐�
        /// </summary>
        public string strTempPath { get; set; }

        /// <summary>
        /// �Z�L���A�t�H���_���X�g
        /// </summary>
        public List<string> lstSecureFolder { get; set; }

        /// <summary>
        /// �u�ŏI�Łv��\��������
        /// </summary>
        public List<string> lstFinal { get; set; }

        /// <summary>
        /// �֘A��Ѓt�@�C���p�X(�T�[�o�[)
        /// </summary>
        public string strGroupDomainListServerPath { get; set; }

        /// <summary>
        /// �֘A��Ѓt�@�C���p�X(���[�J��)
        /// </summary>
        public string strGroupDomainListLocalPath { get; set; }

        /// <summary>
        /// ��E�҃t�@�C���p�X(�T�[�o�[)
        /// </summary>
        public string strManagerListServerPath { get; set; }

        /// <summary>
        /// ��E�҃t�@�C���p�X(���[�J��)
        /// </summary>
        public string strManagerListLocalPath { get; set; }

        /// <summary>
        /// zip���������x��
        /// </summary>
        public int zipLevel { get; set; }

        #region <�萔��`>

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// �@���敪
        /// </summary>
        public const string COMMON_SETDEF_SECLV = "SecrecyS";

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// ���Ə��R�[�h
        /// </summary>
        public const string COMMON_SETDEF_OFFICECODE = "HLI";

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// �����̃T�[�o�[�p�X
        /// </summary>
        public const string COMMON_SETDEF_SABLISTSERVERPATH = "C:\\ProgramData\\SAB\\GCPList.xslx";   // TODO

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// �����̃��[�J���p�X
        /// </summary>
        public const string COMMON_SETDEF_SABLISTLOCALPATH = "C:\\ProgramData\\SAB\\GCPList.xslx";

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// zip�ꎞ�𓀐�
        /// </summary>
        public const string COMMON_SETDEF_TEMPPATH = "C:\\ProgramData\\SAB\\zipTempPath";

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// �Z�L���A�t�H���_���X�g
        /// </summary>
        public const string COMMON_SETDEF_SECUREFOLDER = "C:\\tmp\\SecureA";

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// �u�ŏI�Łv��\��������
        /// </summary>
        public const string COMMON_SETDEF_FINAL = "�ŏI��";

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// �֘A��Ѓt�@�C���p�X(�T�[�o�[)
        /// </summary>
        public const string COMMON_SETDEF_GROUPDOMAINLISTSERVERPATH = "C:\\ProgramData\\SAB\\DomainList.txt";   // TODO

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// �֘A��Ѓt�@�C���p�X(���[�J��)
        /// </summary>
        public const string COMMON_SETDEF_GROUPDOMAINLISTLOCALPATH = "C:\\ProgramData\\SAB\\DomainList.txt";

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// ��E�҃t�@�C���p�X(�T�[�o�[)
        /// </summary>
        public const string COMMON_SETDEF_MANAGERLISTSERVERPATH = "C:\\ProgramData\\SAB\\ManagerList.txt";   // TODO

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// ��E�҃t�@�C���p�X(���[�J��)
        /// </summary>
        public const string COMMON_SETDEF_MANAGERLISTLOCALPATH = "C:\\ProgramData\\SAB\\ManagerList.txt";

        /// <summary>
        /// ���ʐݒ� �f�t�H���g�l
        /// 
        /// zip���������x��
        /// </summary>
        public const int COMMON_SETDEF_ZIPLEVEL = 1;

        #endregion

        /// <summary>
        /// �R���X�g���N�^
        /// </summary>
        public CommonSettings()
        {
            // �f�t�H���g���Ə��R�[�h������
            strOfficeCode = COMMON_SETDEF_OFFICECODE;

            // �f�t�H���g�@���敪������
            strDefaultSecrecyLevel = COMMON_SETDEF_SECLV;

            // �f�t�H���g����R�[�h������
            strCulture = System.Threading.Thread.CurrentThread.CurrentUICulture.ToString();

            // �f�t�H���g�����̃T�[�o�[�p�X������
            strSABListServerPath = COMMON_SETDEF_SABLISTSERVERPATH;

            // �f�t�H���g�����̃��[�J���p�X������
            strSABListLocalPath = COMMON_SETDEF_SABLISTLOCALPATH;

            // �f�t�H���gzip�ꎞ�𓀐揉����
            strTempPath = COMMON_SETDEF_TEMPPATH;

            // �f�t�H���g�Z�L���A�t�H���_���X�g������
            lstSecureFolder = new List<string>() { COMMON_SETDEF_SECUREFOLDER };

            // �f�t�H���g�u�ŏI�Łv��\�������񏉊���
            lstFinal = new List<string>() { COMMON_SETDEF_FINAL };

            // �f�t�H���g�֘A��Ѓt�@�C���p�X(�T�[�o�[)������
            strGroupDomainListServerPath = COMMON_SETDEF_GROUPDOMAINLISTSERVERPATH;

            // �f�t�H���g�֘A��Ѓt�@�C���p�X(���[�J��)������
            strGroupDomainListLocalPath = COMMON_SETDEF_GROUPDOMAINLISTLOCALPATH;

            // �f�t�H���g��E�҃t�@�C���p�X(�T�[�o�[)������
            strManagerListServerPath = COMMON_SETDEF_MANAGERLISTSERVERPATH;

            // �f�t�H���g��E�҃t�@�C���p�X(���[�J��)������
            strManagerListLocalPath = COMMON_SETDEF_MANAGERLISTLOCALPATH;

            // �f�t�H���gzip���������x��������
            zipLevel = COMMON_SETDEF_ZIPLEVEL;
        }
    }

    /// <summary>
    /// ���ʐݒ�t�@�C���̓ǂݍ��ݏ���
    /// </summary>
    public class CommonSettingRead
    {
        /// <summary>
        /// ���ʐݒ�N���X
        /// </summary>
        public CommonSettings commonSettings;

        /// <summary>
        /// ���ʐݒ�t�@�C���ǂݍ���
        /// </summary>
        /// <param name=""></param>
        /// <returns>true:�ǂݍ��ݐ����Afalse:�ǂݍ��ݎ��s</returns>
        public CommonSettings Reader()
        {
            commonSettings = new CommonSettings();

            CommonSettingStoring commonSettingStoring = new CommonSettingStoring();
            // ���ʐݒ�t�@�C���p�X�쐬
            string strCommonSettingFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                commonSettingStoring.COMMON_SETFOLDERNAME,
                commonSettingStoring.COMMON_SETFILENAME
                );

            // ���ʐݒ�t�@�C�������݂��Ȃ��ꍇ�̓f�t�H���g�ݒ����������
            if (File.Exists(strCommonSettingFilePath) == false)
            {
                if (!CommonSettingWrite())
                {
                    string msgErrReadCommonFile = Properties.Resources.msg_err_read_common_file;
                    MessageBox.Show(msgErrReadCommonFile);
                    return null;
                }
            }

            //XmlSerializer�I�u�W�F�N�g�̍쐬
            System.Xml.Serialization.XmlSerializer serXmlCommonRead = new System.Xml.Serialization.XmlSerializer(typeof(CommonSettings));

            //�t�@�C�����J��
            StreamReader stmCommonReader = new StreamReader(strCommonSettingFilePath, Encoding.GetEncoding("shift_jis"));

            //XML�t�@�C������ǂݍ��݁A�t�V���A��������
            commonSettings = (CommonSettings)serXmlCommonRead.Deserialize(stmCommonReader);

            //����
            stmCommonReader.Close();

            return commonSettings;
        }

        /// <summary>
        /// ���ʐݒ�ݒ菑������
        /// </summary>
        /// <returns>true:�����ݐ����Afalse:�����ݎ��s</returns>
        private Boolean CommonSettingWrite()
        {
            try
            {
                CommonSettingStoring commonSettingStoring = new CommonSettingStoring();
                // ���ʐݒ�t�@�C���p�X�쐬
                string strCommonSettingFilePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    commonSettingStoring.COMMON_SETFOLDERNAME
                    );

                if (Directory.Exists(strCommonSettingFilePath) == false)
                {
                    // �t�H���_�쐬
                    System.IO.Directory.CreateDirectory(strCommonSettingFilePath);
                }

                strCommonSettingFilePath = Path.Combine(strCommonSettingFilePath, commonSettingStoring.COMMON_SETFILENAME);

                //XmlSerializer�I�u�W�F�N�g�̍쐬
                System.Xml.Serialization.XmlSerializer serXmlCommonWrite = new System.Xml.Serialization.XmlSerializer(typeof(CommonSettings));

                //�t�@�C�����J��
                System.IO.StreamWriter stmCommonWrite = new System.IO.StreamWriter(strCommonSettingFilePath, false, Encoding.GetEncoding("shift_jis"));

                //�V���A�������AXML�t�@�C���ɕۑ�����
                serXmlCommonWrite.Serialize(stmCommonWrite, commonSettings);

                //����
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
