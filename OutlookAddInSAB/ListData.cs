using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace OutlookAddInSAB
{
    public class ListData
    {
        public class Manager
        {
            enum Classification
            {
                executive, manager
            }
            private string StrClassification;
            private string StrManager;
            public string classification
            {
                get {return StrClassification; }
                set {
                        switch (value)
                        {
                            case "0":
                                StrClassification = Classification.executive.ToString();
                                break;
                            case "1":
                                StrClassification = Classification.manager.ToString();
                                break;
                        }
                    }
            }
            public string manager { get { return StrManager; } set { StrManager = value; } }
        }

        /// <summary>
        /// 関連会社リストの読み込み
        /// </summary>
        public List<string> AssociateList()
        {
            string line = "";
            string dataFilePath = "";
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
        public List<Manager> ManagerList()
        {
            string line = "";
            string dataFilePath = @"C:\Users\shiratori\Documents\SAB_Data\ManagerList.txt";
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

    }
}
