using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelConvertToOkumarukunnCsv.Services;
using ExcelConvertToOkumarukunnCsv.Dto;

namespace ExcelConvertToOkumarukunnCsv.Forms
{
    public partial class OutputOkumarukunnCsvFrm : Form
    {
        public OutputOkumarukunnCsvFrm()
        {
            InitializeComponent();
        }

        private void OutputOkumarukunnCsvFrm_Load(object sender, EventArgs e)
        {
            // ドラッグアンドドロップイベントのハンドラーを登録
            OkurijyoDropArea.DragEnter += OkurijyoDropArea_DragEnter;
            OkurijyoDropArea.DragDrop += OkurijyoDropArea_DragDrop;

            JyutyuDropArea.DragEnter += JyutyuDropArea_DragEnter;
            JyutyuDropArea.DragDrop += JyutyuDropArea_DragDrop;
        }

        // 送り状ファイルを処理するためのDragEnterとDragDropイベント
        private void OkurijyoDropArea_DragEnter(object sender, DragEventArgs e)
        {
            // ファイルがドロップされているかを確認
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // ドロップ可能に設定
                e.Effect = DragDropEffects.All; 
            }
            else
            {
                // ドロップ不可
                e.Effect = DragDropEffects.None; 
            }
        }

        private void OkurijyoDropArea_DragDrop(object sender, DragEventArgs e)
        {
            // ドロップされたファイルを取得
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            // 送り状のファイルチェックと処理を行う
            HandleFileDrop(files, FileType.送り状);
        }

        // 受注一覧ファイルを処理するためのDragEnterとDragDropイベント
        private void JyutyuDropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // ドロップ可能に設定
                e.Effect = DragDropEffects.All; 
            }
            else
            {
                // ドロップ不可
                e.Effect = DragDropEffects.None;
            }
        }

        private void JyutyuDropArea_DragDrop(object sender, DragEventArgs e)
        {
            // ドロップされたファイルを取得
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            // 受注一覧のファイルチェックと処理を行う
            HandleFileDrop(files, FileType.受注一覧);
        }

        // ファイルを処理する共通メソッド
        private void HandleFileDrop(string[] files, FileType fileType)
        {
            foreach (var file in files)
            {
                // ファイルチェック
                if (!isValid(file, fileType))
                {
                    // 無効なファイルがあった場合、処理を中止
                    return;
                }

                // ファイル名の拡張子を変更して出力するCSVファイル名を生成
                string fileName = Path.GetFileNameWithoutExtension(file);
                string outputCsvFileName = fileName + ".csv"; // 拡張子だけ.csvに変更

                // 出力するCSVファイルが存在するか、単にファイル名だけで確認
                if (FileExistsInDirectory(file, outputCsvFileName))
                {
                    // ファイルが存在する場合、開かれているかを確認
                    string outputCsvPath = Path.Combine(Path.GetDirectoryName(file), outputCsvFileName);
                    if (IsFileOpen(outputCsvPath))
                    {
                        // ファイルが開かれている場合、メッセージを表示
                        MessageBox.Show($"'{outputCsvFileName}' は既に開かれています。閉じてからもう一度お試しください。", "ファイルが開いています", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                switch (fileType)
                {
                    case FileType.送り状:
                        // 送り状の処理
                        new FileService().CreateFile(file);
                        break;

                    case FileType.受注一覧:
                        // 受注一覧の処理
                        new FileServiceJ().CreateFile(file);
                        break;
                }
            }
            //    // 出力するCSVファイルのパスを生成
            //    string outputCsvPath = Path.ChangeExtension(file, ".csv");

            //    // CSVファイルが存在する場合のみ、開かれているかチェック
            //    if (File.Exists(outputCsvPath))
            //    {
            //        if (IsFileOpen(outputCsvPath))
            //        {
            //            MessageBox.Show($"'{outputCsvPath}' は既に開かれています。閉じてからもう一度お試しください。", "ファイルが開いています", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //            return;
            //        }
            //    }

            //    switch (fileType)
            //    {
            //        case FileType.送り状:
            //            // 送り状の処理
            //            new FileService().CreateFile(file);
            //            break;  // 追加：処理後にswitch文から抜ける

            //        case FileType.受注一覧:
            //            // 受注一覧の処理
            //            new FileServiceJ().CreateFile(file);
            //            break;  // 追加：処理後にswitch文から抜ける
            //    }
            //}

            ////すべてのファイルが有効だった場合のみ、メッセージを表示
            //MessageBox.Show("おくまるくん用CSVファイルの作成が完了しました。");

            // フォームを閉じて処理を終了
            this.Close();
        }

        // ファイルが有効かどうかを確認するメソッド
        private bool isValid(string file, FileType fileType)
        {
            var permitExtensions = new List<string>
            {
                ".xlsx",
                ".xls"
            };

            string fileExtension = Path.GetExtension(file)?.ToLower();
            string fileName = Path.GetFileNameWithoutExtension(file)?.ToLower();

            // 処理タイプに応じたファイル名チェック
            switch (fileType)
            {
                case FileType.送り状:
                    //if (fileName == null || !fileName.Contains("送り状"))
                    if (fileName == null || !fileName.Contains("出荷"))
                    {
                        MessageBox.Show($"ファイル名に '出荷' が含まれていません。\r\n正しいファイルを選択してください。", "無効なファイル選択", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                    break;

                case FileType.受注一覧:
                    //if (fileName == null || !fileName.Contains("bo明細"))
                    if (fileName == null || !fileName.Contains("受注"))
                    {
                        MessageBox.Show($"ファイル名に '受注' が含まれていません。\r\n正しいファイルを選択してください。", "無効なファイル選択", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                    break;
            }

            // 拡張子チェック
            if (string.IsNullOrEmpty(fileExtension) || !permitExtensions.Contains(fileExtension))
            {
                MessageBox.Show("無効なファイル形式です。Excelファイル (.xlsx または .xls) を選択してください。", "無効なファイル選択", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        // ファイル名のみで出力先ディレクトリにファイルが存在するか確認するメソッド
        private bool FileExistsInDirectory(string originalFile, string outputCsvFileName)
        {
            // 出力先ディレクトリを取得
            string directory = Path.GetDirectoryName(originalFile);

            // フルパスを組み立てて確認
            string outputCsvPath = Path.Combine(directory, outputCsvFileName);

            return File.Exists(outputCsvPath);  // 出力ファイルが存在するかチェック
        }

        // CSVファイルが開かれているかを確認するメソッド
        private bool IsFileOpen(string filePath)
        {
            try
            {
                // ファイルのアクセス権をチェックして、ファイルが開かれているかどうかを確認
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false; // ファイルが開かれていない
                }
            }
            catch (IOException)
            {
                return true; // ファイルが開かれている
            }
        }
        ////CSVファイルが開かれているかを確認するメソッド
        //private bool IsFileOpen(string filePath)
        //{
        //    try
        //    {
        //        // ファイルのアクセス権をチェックして、ファイルが開かれているかどうかを確認
        //        using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        //        {
        //            return false; // ファイルが開かれていない
        //        }
        //    }
        //    catch (IOException)
        //    {
        //        return true; // ファイルが開かれている
        //    }
        //}
    }


    // 処理するファイルの種類を区別するためのenum
    public enum FileType
    {
        送り状,
        受注一覧
    }

}
