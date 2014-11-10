using CSVEditor.BackgroundWork;
using CSVEditor.Helpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CSVEditor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // 초기화
            this._BackgroundWorker = new BackgroundWorker();
            this._BackgroundWorker.WorkerSupportsCancellation = true;
            this._BackgroundWorker.WorkerReportsProgress = true;

            this._BackgroundWorkerForErrorLogging = new BackgroundWorker();
            this._BackgroundWorkerForErrorLogging.WorkerSupportsCancellation = true;
            this._BackgroundWorkerForErrorLogging.WorkerReportsProgress = true;

            this.toolStripProgressBar.Visible = false;
            this.toolStripStatusLabel.Text = "";

            // 열기
            this.openToolStripMenuItem.Click += (sender, e) =>
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.AutoUpgradeEnabled = true;
                dialog.Multiselect = false;
                dialog.DefaultExt = ".csv";
                dialog.Filter = "csv 파일(*.csv)|*.csv|모든 파일(*.*)|*.*";
                if (DialogResult.OK == dialog.ShowDialog())
                {
                    this.LoadFile(dialog.FileName);
                }
            };

            // 저장
            this.saveToolStripMenuItem.Click += (sender, e) =>
            {
                bool blnNeedSave = this.HasNeedSaveData();

                string strCurrentFileName = this.grid.CurrentFilePath;

                if (blnNeedSave && !string.IsNullOrEmpty(strCurrentFileName))
                {
                    this.SaveFile(strCurrentFileName);
                }
                else
                {
                    this.toolStripStatusLabel.Text = "Please, open file.";
                }
            };

            // 다른 이름으로 저장
            this.saveAsToolStripMenuItem.Click += (sender, e) =>
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.AutoUpgradeEnabled = true;
                dialog.DefaultExt = ".csv";
                dialog.Filter = "csv 파일(*.csv)|*.csv|모든 파일(*.*)|*.*";
                if (DialogResult.OK == dialog.ShowDialog())
                {
                    bool blnNeedSave = this.HasNeedSaveData();

                    string strFileName = dialog.FileName;

                    if (blnNeedSave && !string.IsNullOrEmpty(strFileName))
                    {
                        this.SaveFile(strFileName);
                    }
                    else
                    {
                        this.toolStripStatusLabel.Text = "Please, open file.";
                    }
                }
            };

            // 닫기
            this.closeToolStripMenuItem.Click += (sender, e) =>
            {
                this.grid.DataSource = null;
                this.grid.CurrentFilePath = string.Empty;
                // 버튼 제어
                this.EnabledUsersEdit(false);
            };

            // 종료
            this.exitToolStripMenuItem.Click += (sender, e) => { this.Close(); };

            this._BackgroundWorker.DoWork += (sender, e) =>
            {
                Work work = new Work((BackgroundWorker)sender, e);
                work.Run();
            };

            this._BackgroundWorker.ProgressChanged += (sender, e) =>
            {
                UserState state = (UserState)e.UserState;

                #region Work.TASK_OPEN

                if (state.Task.Equals(Work.TASK_OPEN))
                {
                    if (state.Stage.Equals(Work.STAGE_PREPARE))
                    {
                        this.toolStripProgressBar.Visible = true;
                        this.toolStripProgressBar.Maximum = state.ProgressMax;
                        this.toolStripProgressBar.Value = state.ProgressValue;

                        this.toolStripStatusLabel.Text = "";

                        this.grid.DataSource = null;
                    }

                    if (state.Stage.Equals(Work.STAGE_PROCESS))
                    {
                        if (this.grid.DataSource == null)
                        {
                            DataTable dataTable = new DataTable();

                            if (state.IsHeader)
                            {
                                foreach (string col in state.Data)
                                    dataTable.Columns.Add(col);
                            }
                            else
                            {
                                for (int i = 0; i < state.ColumnCount; i++)
                                {
                                    dataTable.Columns.Add(string.Format("COL{0}", (i + 1)));
                                }
                            }

                            this.grid.AutoGenerateColumns = true;
                            this.grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                            this.grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
                            this.grid.AllowUserToAddRows = true;
                            this.grid.AllowUserToDeleteRows = false;

                            //this.grid.RowHeadersWidth = 50;
                            this.grid.DataSource = dataTable;
                            this.grid.CurrentFilePath = state.FileName;
                        }
                        else
                        {
                            DataTable dataSource = (DataTable)this.grid.DataSource;
                            dataSource.Rows.Add(state.Data);
                            dataSource.AcceptChanges();
                        }

                        this.toolStripProgressBar.Value = state.ProgressValue;

                        this.toolStripStatusLabel.Text = "Loading...";
                    }

                    if (state.Stage.Equals(Work.STAGE_COMPLETED))
                    {
                        this.toolStripProgressBar.Visible = false;
                        this.toolStripStatusLabel.Text = "Completed.";

                        // 버튼 제어
                        this.EnabledUsersEdit(true);
                    }
                }

                #endregion Work.TASK_OPEN

                #region Work.TASK_SAVE

                if (state.Task.Equals(Work.TASK_SAVE))
                {
                    if (state.Stage.Equals(Work.STAGE_PREPARE))
                    {
                        this.toolStripProgressBar.Visible = true;
                        this.toolStripProgressBar.Maximum = state.ProgressMax;
                        this.toolStripProgressBar.Value = state.ProgressValue;

                        this.toolStripStatusLabel.Text = "";
                    }

                    if (state.Stage.Equals(Work.STAGE_PROCESS))
                    {
                        this.toolStripProgressBar.Value = state.ProgressValue;

                        this.toolStripStatusLabel.Text = "Wait ...";
                    }

                    if (state.Stage.Equals(Work.STAGE_COMPLETED))
                    {
                        this.toolStripProgressBar.Visible = false;
                        this.toolStripStatusLabel.Text = "Saved.";
                        this.grid.CurrentFilePath = state.FileName;
                    }
                }

                #endregion Work.TASK_SAVE

                if (state.Stage.Equals(Work.STAGE_ERROR))
                {
                    Arguments arg = new Arguments();
                    arg.Task = Work.TASK_ERRORLOG;
                    arg.Add(BackgroundWorkArgumentKeys.SAVE_DATASOURCE, state.Message);

                    this._BackgroundWorkerForErrorLogging.RunWorkerAsync(arg);
                }
            };

            this._BackgroundWorker.RunWorkerCompleted += (sender, e) =>
            {
                if (e.Error != null)
                {
                    // 오류
                }
                else if (e.Cancelled)
                {
                    // 취소
                }
                else
                {
                    // 완료
                    this.grid.Enabled = true;
                }
            };

            // 종료 체크
            this.FormClosing += (sender, e) => { };

            this.Load += (sender, e) => { };

            this.grid.CellValueChanged += (sender, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    this.grid.Rows[e.RowIndex].HeaderCell.Style = new DataGridViewCellStyle() { BackColor = Color.HotPink };
                    this.grid.Rows[e.RowIndex].HeaderCell.Value = "M..";
                    this.grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.HotPink;
                    this.grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Style = new DataGridViewCellStyle() { BackColor = Color.HotPink, SelectionBackColor = Color.HotPink };

                    if (this.grid.Rows[e.RowIndex].IsNewRow)
                    {
                        DataTable dataSource = (this.grid.DataSource as DataTable);
                        dataSource.Rows.Add(((DataRowView)this.grid.Rows[e.RowIndex].DataBoundItem).Row.ItemArray);
                    }
                }
            };

            // 인코딩
            foreach (ToolStripMenuItem item in this.encodingToolStripMenuItem.DropDownItems)
            {
                item.Checked = item.Text.Equals(Properties.Settings.Default.TextFileEncodingName);

                item.Click += (sender, e) =>
                {
                    ToolStripMenuItem thisControl = (ToolStripMenuItem)sender;
                    ToolStripMenuItem paentControl = (ToolStripMenuItem)thisControl.OwnerItem;
                    string strCheckedItemText = string.Empty;

                    foreach (ToolStripMenuItem itm in paentControl.DropDownItems)
                    {
                        if (itm.Checked)
                        {
                            strCheckedItemText = itm.Text;
                            break;
                        }
                    }

                    foreach (ToolStripMenuItem itm in paentControl.DropDownItems)
                    {
                        itm.Checked = itm.Text.Equals(thisControl.Text);
                    }

                    if (!thisControl.Text.Equals(strCheckedItemText))
                    {
                        Properties.Settings.Default.TextFileEncodingName = this.GetEncoding();
                        Properties.Settings.Default.Save();
                        if (this.grid.DataSource != null)
                        {
                            this.LoadFile();
                        }
                    }
                };
            }

            int nChecedToolStripMenuItemCount = 0;

            foreach (ToolStripMenuItem item in this.encodingToolStripMenuItem.DropDownItems)
                nChecedToolStripMenuItemCount++;

            if (nChecedToolStripMenuItemCount == 0)
            {
                foreach (ToolStripMenuItem item in this.encodingToolStripMenuItem.DropDownItems)
                    item.Checked = item.Text.Equals(ManagedEncoding.UTF8);
            }

            // Copy Selected Row Data
            this.copyRowsDataToolStripMenuItem.Click += (sender, e) =>
            {
                this.SetClipBoardTextWithSelectedRowsData();
            };

            // Delete Selected Rows
            this.deleteRowsDataToolStripMenuItem.Click += (sender, e) =>
            {
                this.DeleteSelectedRows();
            };
            this.grid.KeyDown += (sender, e) =>
            {
                if (e.Modifiers == Keys.Control && e.KeyCode == Keys.C)
                {
                    // Ctrl + C :  Copy Row Data
                    this.SetClipBoardTextWithSelectedRowsData();
                }

                if (e.KeyCode == Keys.Delete)
                {
                    this.DeleteSelectedRows();
                }
            };

            // 버튼 제어
            this.EnabledUsersEdit(false);

            #region Error Log

            this._BackgroundWorkerForErrorLogging.DoWork += (sender, e) =>
            {
                Work work = new Work((BackgroundWorker)sender, e);
                work.Run();
            };

            #endregion Error Log
        }

        #region private - Method

        private bool HasNeedSaveData()
        {
            return true;
            //try
            //{
            //    DataTable dataSource = this.grid.DataSource as DataTable;
            //    if (dataSource == null) { throw new Exception("DataSource is null."); }

            //    if ((this.grid.Rows.Count - 1) != dataSource.Rows.Count) { return true; }

            //    foreach (DataGridViewRow row in this.grid.Rows)
            //    {
            //        string strHeaderValue = string.Format("{0}", row.HeaderCell.Value);
            //        if (!string.IsNullOrEmpty(strHeaderValue)) { return true; }
            //    }

            //    return false;
            //}
            //catch { return false; }
        }

        private void UpdateDataSource()
        {
            string[] arrGridColumnKeys = this.GetGridColumnDisplayOrder();
            DataTable dataSource = this.grid.DataSource as DataTable;

            for (int i = arrGridColumnKeys.Length - 1; i >= 0; i--)
            {
                dataSource.Columns[arrGridColumnKeys[i]].SetOrdinal(i);
            }

            dataSource.AcceptChanges();
        }

        private string[] GetGridColumnDisplayOrder()
        {
            List<string> lsResult = new List<string>();
            int nColumnCount = this.grid.Columns.Count;
            for (int i = 0; i < nColumnCount; i++)
            {
                foreach (DataGridViewColumn col in this.grid.Columns)
                {
                    if (col.DisplayIndex == i)
                    {
                        lsResult.Add(col.Name);
                    }
                }
            }

            return lsResult.ToArray();
        }

        private string GetEncoding()
        {
            foreach (ToolStripMenuItem item in this.encodingToolStripMenuItem.DropDownItems)
            {
                if (item.Checked)
                {
                    return item.Text;
                }
            }

            return ManagedEncoding.UTF8;
        }

        private void LoadFile(string filePath)
        {
            Arguments arg = new Arguments();
            arg.Task = Work.TASK_OPEN;
            arg.Add(BackgroundWorkArgumentKeys.OPEN_FILE_PATH, filePath);
            arg.Add(BackgroundWorkArgumentKeys.FIRST_ROW_IS_HEADER, true);
            arg.Add(BackgroundWorkArgumentKeys.ENCODING, this.GetEncoding());

            this.grid.Enabled = false;

            this._BackgroundWorker.RunWorkerAsync(arg);
        }

        private void LoadFile()
        {
            this.LoadFile(this.grid.CurrentFilePath);
        }

        private void SaveFile(string filePath)
        {
            this.UpdateDataSource();

            Arguments arg = new Arguments();
            arg.Task = Work.TASK_SAVE;
            arg.Add(BackgroundWorkArgumentKeys.SAVE_FILE_PATH, filePath);
            arg.Add(BackgroundWorkArgumentKeys.SAVE_DATASOURCE, (DataTable)this.grid.DataSource);
            arg.Add(BackgroundWorkArgumentKeys.FIRST_ROW_IS_HEADER, true);
            arg.Add(BackgroundWorkArgumentKeys.ENCODING, this.GetEncoding());

            this.grid.Enabled = false;

            this._BackgroundWorker.RunWorkerAsync(arg);
        }

        private void SetClipBoardTextWithSelectedRowsData()
        {
            if (this.grid.SelectedRows != null && this.grid.SelectedRows.Count > 0)
            {
                StringBuilder sbRowData = new StringBuilder();
                for (int i = this.grid.SelectedRows.Count - 1; i >= 0; i--)
                {
                    foreach (DataGridViewCell cell in this.grid.SelectedRows[i].Cells)
                    {
                        sbRowData.AppendFormat("{0},", cell.Value);
                    }
                    sbRowData.Remove(sbRowData.Length - 1, 1);
                    sbRowData.AppendLine();
                }
                if (sbRowData.Length > 0)
                {
                    string strRowData = sbRowData.ToString(0, sbRowData.Length - 2);
                    Clipboard.SetText(strRowData);
                }
            }
        }

        private void DeleteSelectedRows()
        {
            if (this.grid.SelectedRows != null && this.grid.SelectedRows.Count > 0)
            {
                string strMessage = string.Format("Do you want to delete selected {0:n0} Row(s)?", this.grid.SelectedRows.Count);
                DialogResult result = MessageBox.Show(strMessage, "Confirmation", MessageBoxButtons.YesNo);
                if (DialogResult.Yes == result)
                {
                    foreach (DataGridViewRow row in this.grid.SelectedRows)
                    {
                        if (row.IsNewRow) { continue; }
                        this.grid.Rows.Remove(row);
                    }
                }
            }
        }

        private void EnabledUsersEdit(bool enabled)
        {
            this.saveAsToolStripMenuItem.Enabled = enabled;
            this.saveToolStripMenuItem.Enabled = enabled;
            this.deleteRowsDataToolStripMenuItem.Enabled = enabled;
            this.copyRowsDataToolStripMenuItem.Enabled = enabled;
            this.closeToolStripMenuItem.Enabled = enabled;
        }

        #endregion private - Method

        #region private - Properties

        #endregion private - Properties

        #region EventHandlers

        #endregion EventHandlers

        #region public - Methods

        #endregion public - Methods

        #region public - Properties

        #endregion public - Properties

        #region private - Fields

        private BackgroundWorker _BackgroundWorker;
        private BackgroundWorker _BackgroundWorkerForErrorLogging;

        #endregion private - Fields
    }
}