using CSVEditor.Helpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace CSVEditor.BackgroundWork
{
    public class Work
    {
        #region Task

        public const string TASK_OPEN = "OPEN";
        public const string TASK_SAVE = "SAVE";
        public const string TASK_SAVEAS = "SAVEAS";
        public const string TASK_ERRORLOG = "ERRORLOG";

        #endregion Task

        #region Stage

        public const string STAGE_PREPARE = "PREPARE";
        public const string STAGE_START = "START";
        public const string STAGE_PROCESS = "PROCESS";
        public const string STAGE_COMPLETED = "COMPLETED";
        public const string STAGE_ERROR = "ERROR";

        #endregion Stage

        public Work(BackgroundWorker worker, DoWorkEventArgs e)
        {
            this._worker = worker;
            this._e = e;
        }

        public void Run()
        {
            if (this._e.Argument is Arguments)
            {
                Arguments arg = (Arguments)this._e.Argument;
                RunDelegate run = null;
                switch (arg.Task)
                {
                    case Work.TASK_OPEN:
                        run = this.RunOpen;
                        break;

                    case Work.TASK_SAVE:
                        run = this.RunSave;
                        break;

                    case Work.TASK_ERRORLOG:
                        run = this.RunLog;
                        break;
                }

                if (run != null)
                {
                    run(arg);
                }
            }
            else
            {
                throw new Exception("매개변수를 처리할 수 없습니다.");
            }
        }

        private void RunOpen(Arguments arg)
        {
            try
            {
                string strPath = string.Format("{0}", arg[BackgroundWorkArgumentKeys.OPEN_FILE_PATH]);
                bool blnFirstRowIsHeader = (bool)arg[BackgroundWorkArgumentKeys.FIRST_ROW_IS_HEADER];
                string strEncoding = string.Format("{0}", arg[BackgroundWorkArgumentKeys.ENCODING]);
                if (string.IsNullOrEmpty(strEncoding)) { strEncoding = ManagedEncoding.UTF8; }
                Encoding encoding = this.GetEncofing(strEncoding);

                bool blnIsHeader = false;
                string strReadline = null;  // String Read Line
                string[] arrBuffer = null;  // 메인 스레드에서 사용될 데이터
                int nRow = 0;
                int nPercentage = 0;

                if (string.IsNullOrEmpty(strPath))
                {
                    throw new Exception("빈 경로의 파일을 열수 없습니다.");
                }

                using (FileStream fileStream = new FileStream(strPath, FileMode.Open, FileAccess.Read))
                {
                    using (StreamReader reader = new StreamReader(fileStream, encoding))
                    {
                        this._worker.ReportProgress(0,
                            new UserState()
                            {
                                Task = arg.Task,
                                Stage = Work.STAGE_PREPARE,
                                ProgressMax = (int)fileStream.Length,
                                ProgressValue = 0,
                                Data = null,
                                IsHeader = false,
                                FileName = strPath
                            });

                        while (null != (strReadline = reader.ReadLine()))
                        {
                            if (nRow == 0 && blnFirstRowIsHeader)
                            {
                                // 첫번째 행을 헤더로 사용하는 경우
                                blnIsHeader = true;
                            }
                            else { blnIsHeader = false; }

                            arrBuffer = strReadline.Split(',');
                            nPercentage = (int)(((fileStream.Position * 1.0) / (fileStream.Length * 1.0)) * 100);
                            this._worker.ReportProgress(
                                nPercentage,
                                new UserState()
                                {
                                    Task = arg.Task,
                                    Stage = Work.STAGE_PROCESS,
                                    ProgressMax = (int)fileStream.Length,
                                    ProgressValue = (int)fileStream.Position,
                                    Data = arrBuffer,
                                    IsHeader = blnIsHeader,
                                    FileName = strPath
                                });

                            // 처리중 행 증가
                            nRow++;
                        }

                        reader.Close();
                    }
                    fileStream.Close();
                }
                // 완료
                this._worker.ReportProgress(
                       100,
                       new UserState()
                       {
                           Task = arg.Task,
                           Stage = Work.STAGE_COMPLETED,
                           ProgressMax = 100,
                           ProgressValue = 100,
                           Data = null,
                           IsHeader = false,
                           FileName = strPath
                       });
            }
            catch (Exception ex)
            {
                // 오류
                this._worker.ReportProgress(
                       0,
                       new UserState()
                       {
                           Task = arg.Task,
                           Stage = Work.STAGE_ERROR,
                           ProgressMax = 0,
                           ProgressValue = 0,
                           Data = null,
                           IsHeader = false,
                           FileName = null,
                           Message = ex.Message
                       });
            }
        }

        private void RunSave(Arguments arg)
        {
            try
            {
                string strSaveFilePath = string.Format("{0}", arg[BackgroundWorkArgumentKeys.SAVE_FILE_PATH]);
                string strEncoding = string.Format("{0}", arg[BackgroundWorkArgumentKeys.ENCODING]);
                if (string.IsNullOrEmpty(strEncoding)) { strEncoding = ManagedEncoding.UTF8; }
                Encoding encoding = this.GetEncofing(strEncoding);
                DataTable dataSource = (DataTable)arg[BackgroundWorkArgumentKeys.SAVE_DATASOURCE];

                if (string.IsNullOrEmpty(strSaveFilePath))
                {
                    throw new Exception("빈 경로의 파일을 저장할 수 없습니다.");
                }
                bool blnFirstRowIsHeader = (bool)arg[BackgroundWorkArgumentKeys.FIRST_ROW_IS_HEADER];
                string strReadline = string.Empty;  // String Read Line
                StringBuilder sbLine;
                int nRowCount = dataSource.Rows.Count;
                int nCurrentRow = 0;
                int nPercentage = 0;
                int nMaxWorkCount = 0;

                nMaxWorkCount = nRowCount + (blnFirstRowIsHeader ? 1 : 0);

                this._worker.ReportProgress(0,
                        new UserState()
                        {
                            Task = arg.Task,
                            Stage = Work.STAGE_PREPARE,
                            ProgressMax = nMaxWorkCount,
                            ProgressValue = nCurrentRow,
                            Data = null,
                            IsHeader = false,
                            FileName = strSaveFilePath
                        });

                using (FileStream fileStream = new FileStream(strSaveFilePath, FileMode.Create, FileAccess.Write))
                {
                    using (StreamWriter writer = new StreamWriter(fileStream, encoding))
                    {
                        if (blnFirstRowIsHeader)
                        {
                            sbLine = new StringBuilder();
                            foreach (DataColumn dc in dataSource.Columns)
                            {
                                sbLine.AppendFormat("{0},", dc.ColumnName);
                            }
                            writer.WriteLine(sbLine.ToString().Substring(0, sbLine.Length - 1));

                            nCurrentRow++;

                            nPercentage = (int)(((nCurrentRow * 1.0) / (nMaxWorkCount * 1.0)) * 100);
                            this._worker.ReportProgress(nPercentage,
                                        new UserState()
                                        {
                                            Task = arg.Task,
                                            Stage = Work.STAGE_PROCESS,
                                            ProgressMax = nMaxWorkCount,
                                            ProgressValue = nCurrentRow,
                                            Data = null,
                                            IsHeader = false,
                                            FileName = strSaveFilePath
                                        });
                        }

                        for (int i = 0; i < nRowCount; i++)
                        {
                            sbLine = new StringBuilder();
                            for (int j = 0; j < dataSource.Columns.Count; j++)
                            {
                                sbLine.AppendFormat("{0},", dataSource.Rows[i][j]);
                            }

                            writer.WriteLine(sbLine.ToString().Substring(0, sbLine.Length - 1));

                            nPercentage = (int)(((nCurrentRow * 1.0) / (nMaxWorkCount * 1.0)) * 100);
                            this._worker.ReportProgress(nPercentage,
                                        new UserState()
                                        {
                                            Task = arg.Task,
                                            Stage = Work.STAGE_PROCESS,
                                            ProgressMax = nMaxWorkCount,
                                            ProgressValue = nCurrentRow,
                                            Data = null,
                                            IsHeader = false,
                                            FileName = strSaveFilePath
                                        });
                        }

                        writer.Flush();
                        writer.Close();
                    }
                    fileStream.Close();
                }

                // 완료
                this._worker.ReportProgress(
                       100,
                       new UserState()
                       {
                           Task = arg.Task,
                           Stage = Work.STAGE_COMPLETED,
                           ProgressMax = 100,
                           ProgressValue = 100,
                           Data = null,
                           IsHeader = false,
                           FileName = strSaveFilePath
                       });
            }
            catch (Exception ex)
            {
                // 오류
                this._worker.ReportProgress(
                       0,
                       new UserState()
                       {
                           Task = arg.Task,
                           Stage = Work.STAGE_ERROR,
                           ProgressMax = 0,
                           ProgressValue = 0,
                           Data = null,
                           IsHeader = false,
                           FileName = null,
                           Message = ex.Message
                       });
            }
        }

        /// <summary>
        /// 필요 매개변수 BackgroundWorkArgumentKeys.SAVE_DATASOURCE
        /// </summary>
        /// <param name="arg"></param>
        private void RunLog(Arguments arg)
        {
            string strDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            strDirectory = string.Format("{0}\\CsvEditor\\Log\\", strDirectory);

            DirectoryInfo dir = new DirectoryInfo(strDirectory);
            if (!dir.Exists)
            {
                dir.Create();
            }

            string strSaveFilePath = string.Format("{0}Error-{1:yyyy-MM-dd}.log", dir.FullName, DateTime.Today);

            string strEncoding = string.Empty;  // UTF-8로 설정 :)
            if (string.IsNullOrEmpty(strEncoding)) { strEncoding = ManagedEncoding.UTF8; }
            Encoding encoding = this.GetEncofing(strEncoding);
            string strLogMessage = string.Format("{0}", arg[BackgroundWorkArgumentKeys.SAVE_DATASOURCE]);       // Exception.Message

            if (string.IsNullOrEmpty(strSaveFilePath))
            {
                throw new Exception("빈 경로의 파일을 저장할 수 없습니다.");
            }

            int nCurrentRow = 0;
            int nPercentage = 0;
            int nMaxWorkCount = 0;

            this._worker.ReportProgress(0,
                    new UserState()
                    {
                        Task = arg.Task,
                        Stage = Work.STAGE_PREPARE,
                        ProgressMax = nMaxWorkCount,
                        ProgressValue = nCurrentRow,
                        Data = null,
                        IsHeader = false,
                        FileName = strSaveFilePath
                    });

            using (FileStream fileStream = new FileStream(strSaveFilePath, FileMode.Append, FileAccess.Write))
            {
                using (StreamWriter writer = new StreamWriter(fileStream, encoding))
                {
                    writer.WriteLine(strLogMessage);

                    this._worker.ReportProgress(nPercentage,
                                new UserState()
                                {
                                    Task = arg.Task,
                                    Stage = Work.STAGE_PROCESS,
                                    ProgressMax = nMaxWorkCount,
                                    ProgressValue = nCurrentRow,
                                    Data = null,
                                    IsHeader = false,
                                    FileName = strSaveFilePath
                                });

                    writer.Flush();
                    writer.Close();
                }
                fileStream.Close();
            }

            // 완료
            this._worker.ReportProgress(
                   100,
                   new UserState()
                   {
                       Task = arg.Task,
                       Stage = Work.STAGE_COMPLETED,
                       ProgressMax = 100,
                       ProgressValue = 100,
                       Data = null,
                       IsHeader = false,
                       FileName = strSaveFilePath
                   });
        }

        private Encoding GetEncofing(string encodingName)
        {
            switch (encodingName)
            {
                case ManagedEncoding.ANSI:
                    // 코드페이지 번호는 http://msdn.microsoft.com/ko-kr/library/system.text.encoding.aspx 에서 확인하시면 됩니다.
                    int euckrCodepage = 51949;
                    return Encoding.GetEncoding(euckrCodepage);

                case ManagedEncoding.UNICODE:
                    return Encoding.Unicode;

                default:
                    return Encoding.UTF8;
            }
        }

        #region private Fields

        private BackgroundWorker _worker;
        private DoWorkEventArgs _e;

        private delegate void RunDelegate(Arguments arg);

        #endregion private Fields
    }

    public class UserState
    {
        public string Task { get; set; }

        public string Stage { get; set; }

        public string[] Data { get; set; }

        public bool IsHeader { get; set; }

        public int ColumnCount { get { return this.Data.Length; } }

        public int ProgressValue { get; set; }

        public int ProgressMax { get; set; }

        public string FileName { get; set; }

        public string Message { get; set; }
    }
}