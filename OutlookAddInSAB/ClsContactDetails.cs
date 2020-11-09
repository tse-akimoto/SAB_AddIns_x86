using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddInSAB
{
    public class ClsContactDetails
    {
        /// <summary>
        /// アドレス
        /// </summary>
        public string Address { get; set; }

        /// <summary>
        /// ドメイン
        /// </summary>
        public string Domain { get; set; }

        /// <summary>
        /// 連絡先情報の有無
        /// </summary>
        public static bool ContactDetailsExistence { get; set; }

        /// <summary>
        /// 連絡先情報
        /// </summary>
        public static List<string> ContactDetails { get; set; }

        /// <summary>
        /// 送信先
        /// </summary>
        public static string SenderAddress { get; set; }

        /// <summary>
        /// CC
        /// </summary>
        public static string CarbonCopy { get; set; }

        /// <summary>
        /// BCC
        /// </summary>
        public static string BlindCarbonCopy { get; set; }
        /// <summary>
        /// 役職
        /// </summary>
        public string JobTitle { get; set; }
        /// <summary>
        /// 役職区分
        /// </summary>
        public string jobClassification { get; set; }
    }
}
