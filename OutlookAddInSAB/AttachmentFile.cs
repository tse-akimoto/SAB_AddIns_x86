using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddInSAB
{
    class AttachmentFile
    {
        Outlook.Application app = new Outlook.Application();

        public AttachmentFile()
        {
            
        }

        /// <summary>
        /// 設定したメール情報を取得
        /// </summary>
        public void file_data()
        {
            Dictionary<int, string> dicFile = new Dictionary<int, string>();

            // Inspectorを取得、MailItemを取得、Attchmentを取得
            Outlook.Inspector ins = app.ActiveInspector();
            Outlook.MailItem item = ins.CurrentItem as Outlook.MailItem;
            Outlook.Attachments attchments = item.Attachments;

            if (attchments.Count == 0)
                MessageBox.Show("添付ファイルなし", AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            else
            {
                foreach (Outlook.Attachment attach in attchments)
                {
                    string name = attach.FileName;
                    string display_name = attach.DisplayName;
                    string path = attach.GetTemporaryFilePath();
                    string path_name = attach.PathName;
                    int hash = attach.GetHashCode();

                    MessageBox.Show(name + "\r\n" + display_name + "\r\n" + path, AddInsLibrary.Properties.Resources.msgError, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }
    }
}
