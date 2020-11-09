using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddInSAB
{
    public class FileList
    {
        public Outlook.Attachment attachment { get; set; }
        public string fileName { get; set; }
        public string filePath { get; set; }
        public string fileExtension { get; set; }
        /// <summary>
        /// ‹@–§‹æ•ª
        /// </summary>
        public string fileSecrecy { get; set; }
        /// <summary>
        /// –‹ÆŠ
        /// </summary>
        public string fileOfficeCode { get; set; }
        /// <summary>
        /// •¶‘•ª—Ş
        /// </summary>
        public string fileClassification { get; set; }
        public List<FileList> file_list { get; set; }

        public FileList()
        {

        }
    }
}
