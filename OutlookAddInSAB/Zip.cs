using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Ionic.Zip;
using System.Windows.Forms;

namespace OutlookAddInSAB
{
    class Zip
    {
        #region 定義

        /// <summary>
        /// 添付されているzipファイルの解凍先フォルダ
        /// 
        /// 一時フォルダは処理の裏で使用されるものなのでルール化する
        /// 
        /// 解凍フォルダは圧縮対象外なので削除する必要あり
        ///   ⇒  削除しやすいように解凍先を指定する
        /// </summary>
        public string UNZIP_FOLDER = "unzip_folder";

        /// <summary>
        /// zipファイル名
        /// </summary>
        public static string zipPath = "";

        /// <summary>
        /// zipファイルのパス
        /// </summary>
        public static string zipFilePath { get; set; }

        #endregion

        public Zip()
        {
           
        }

        #region 解凍

        /// <summary>
        /// zipを解凍するメソッド
        /// </summary>
        /// <param name="filePath">解凍対象のzipファイル</param>
        /// <param name="tempPath">一時フォルダのパス</param>
        /// <param name="list">添付ファイルリスト</param>
        /// <returns>添付ファイルリスト</returns>
        public List<ClsFilePropertyList> UnZip(string filePath, string tempPath, List<ClsFilePropertyList> list, ref bool error)
        {
            error = false;

            var status = System.Console.Out;
            var encode = System.Text.Encoding.GetEncoding("Shift_JIS");

            string[] extractDir = filePath.Split('/');
            string dirName = filePath.Replace(".zip", "");
            string highTempPath = (tempPath.EndsWith(".zip") == true) ? tempPath.Substring(0, tempPath.Length - 4) : tempPath;
            dirName = Path.Combine(highTempPath, extractDir[extractDir.Length - 1].Replace(".zip", ""));

            dirName += "\\" + UNZIP_FOLDER;

            // 解凍先ディレクトリを作成する
            if (!Directory.Exists(dirName))
            {
                Directory.CreateDirectory(dirName);
            }
            // 同名のディレクトリが存在する場合はフォルダ名末尾に数字を追加
            else
            {
                for (int i = 1; ; i++)
                {
                    if (!Directory.Exists(dirName + i.ToString()))
                    {
                        dirName = dirName + "_" + i.ToString();
                        Directory.CreateDirectory(dirName);
                        break;
                    }
                }
            }

            List<ClsFilePropertyList> file_list = new List<ClsFilePropertyList>();
            try
            {
                var enc = new ReadOptions() { Encoding = Encoding.GetEncoding("shift_jis") };   // 2020/09/10  文字化け対応

                using (ZipFile zip = ZipFile.Read(filePath, enc))
                {
                    // zipにパスワードが設定されているかチェック
                    bool passCheck = ZipFile.CheckZipPassword(zip.Name, "");
                    if (passCheck == false)
                    {
                        string[] passFileArray = new string[zip.EntryFileNames.Count];
                        zip.EntryFileNames.CopyTo(passFileArray, 0);
                        throw new BadPasswordException();
                    }
                    // 展開して一つ一つリストに入れていく
                    foreach (ZipEntry entry in zip)
                    {
                        // zipが出てきた場合再帰呼び出しをして解凍する
                        if (entry.FileName.EndsWith(".zip") == true)
                        {
                            // zipファイルと同名のフォルダがあるか調べる
                            var entryCollection = zip.EntryFileNames;
                            List<string> entryList = new List<string>();
                            entryList.AddRange(entryCollection);

                            List<string> entryFolderName = entryList.FindAll(x => x.Contains('/'));
                            List<string> entFolderList = new List<string>();
                            for (int i = 0; i < entryFolderName.Count; i++)
                            {
                                entFolderList.Add(entryFolderName[i].Split('/')[0]);
                            }
                            string entryName = entry.FileName.Substring(0, entry.FileName.Length - 4);

                            // zipと同名のフォルダがあった場合展開先のフォルダ名の末尾に数字を追加
                            if (entFolderList.Contains(entryName))
                            {
                                string newDirName = "";
                                for (int i = 1; ; i++)
                                {
                                    if (!Directory.Exists(dirName + i.ToString()))
                                    {
                                        newDirName = Path.Combine(dirName, entryName) + "_" + i.ToString();
                                        Directory.CreateDirectory(newDirName);
                                        break;
                                    }
                                }
                                entry.Extract(newDirName);
                                string file_path = Path.Combine(newDirName, entry.FileName);

                                string deployPath = Path.Combine(newDirName, entry.FileName.Substring(0, entry.FileName.Length - 4));

                                file_list.Add(new ClsFilePropertyList { fileName = entry.FileName, filePath = Path.Combine(newDirName, entry.FileName), fileExtension = ".zip", file_list = UnZip(file_path, deployPath, file_list, ref error) });
                            }
                            else
                            {
                                // 展開先フォルダを作ってその中に展開する
                                entry.Extract(dirName);
                                string file_path = Path.Combine(dirName, entry.FileName);
                                string deployPath = Path.Combine(dirName, entry.FileName.Substring(0, entry.FileName.Length - 4));
                                file_list.Add(new ClsFilePropertyList { fileName = entry.FileName, filePath = Path.Combine(dirName, entry.FileName), fileExtension = ".zip", file_list = UnZip(file_path, deployPath, file_list, ref error) });
                            }

                        }
                        else if (entry.IsDirectory == false)
                        {
                            entry.Extract(dirName);
                            string file_path = Path.Combine(dirName, entry.FileName);
                            file_list.Add(new ClsFilePropertyList { fileName = entry.FileName, filePath = file_path, fileExtension = entry.FileName.Substring(entry.FileName.LastIndexOf(".")) });
                        }
                    }
                }
            }
            catch (BadPasswordException e)
            {
                file_list.Clear();

                MessageBox.Show(AddInsLibrary.Properties.Resources.msgPasswordZip,
                        AddInsLibrary.Properties.Resources.msgError,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Hand);

                error = true;
            }
            catch (Ionic.Zlib.ZlibException e)
            {
                string ErrorMsg = "";
                if (e.ToString().Contains("(oversubscribed dynamic bit lengths tree)") == true)
                {
                    ErrorMsg = AddInsLibrary.Properties.Resources.msgPasswordZip;
                }
                else
                {
                    ErrorMsg = AddInsLibrary.Properties.Resources.msgUnZipError;
                }

                file_list.Clear();

                MessageBox.Show(ErrorMsg,
                    AddInsLibrary.Properties.Resources.msgError,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Hand);

                error = true;
            }
            catch (Exception e)
            {
                file_list.Clear();

                MessageBox.Show(AddInsLibrary.Properties.Resources.msgUnZipError,
                        AddInsLibrary.Properties.Resources.msgError,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Hand);

                error = true;
            }

            return file_list;
        }

        #endregion

        #region 圧縮

        /// <summary>
        /// zip圧縮するメソッド
        /// </summary>
        /// <param name="filePathList">一時フォルダに格納されている添付ファイル</param>
        /// <param name="compressionFileName">圧縮する際のファイル名。添付ファイルの最初のファイル名</param>
        /// <returns>パスワード</returns>
        public string ZipCompression(string[] filePathList, string compressionFileName)
        {
            // 2020/09/11 圧縮する際のファイル名を最初に添付したファイル名に変更
            string zipName = Path.GetFileNameWithoutExtension(compressionFileName) + ".zip";

            zipPath = zipName;
            string pass = "";

            using (ZipFile zip = new ZipFile())
            {
                // エンコードの設定
                zip.ProvisionalAlternateEncoding = Encoding.GetEncoding("Shift_JIS");

                // 圧縮レベルの設定
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;

                // 必要な時はZIP64で圧縮する
                zip.UseZip64WhenSaving = Zip64Option.AsNecessary;

                // 8桁のランダムなパスワードを生成
                pass = System.Web.Security.Membership.GeneratePassword(8, 0);
                zip.Password = pass;

                zip.Encryption = EncryptionAlgorithm.PkzipWeak;        // Zip2.0暗号化

                foreach (string filePath in filePathList)
                {
                    string attachPath = Path.Combine(filePath);
                    zip.AddFile(attachPath, "");
                }

                // zipファイルの作成
                zip.Save(zipName);
            }

            return pass;
        }

        #endregion
    }
}
