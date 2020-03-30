using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Security.Permissions;
using Npgsql;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;

namespace TrayGuard
{
    public partial class frmTray : Form
    {
        //親フォームfrmLoginへ、イベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        // デスクトップの基本設定ファイルのパス
        string productconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\tray_guard_desktop.ini";

        //データグリッドビュー用ボタン
        DataGridViewButtonColumn openTray;

        //その他非ローカル変数
        DataTable dtTray;
        public static string userRole;
        string userId;

        // コンストラクタ
        public frmTray()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void frmTray_Load(object sender, EventArgs e)
        {
            this.Text = this.Text + " " + Assembly.GetExecutingAssembly().GetName().Version;
            // フォームの場所を指定
            this.Left = 20;
            this.Top = 10;

            // デスクトップ設定ファイルにラインの指定がある場合は、当フォームのテキストボックス・チェックボックスに反映する
            string line = readIni("MODULE NUMBERING CHECK", "LINE", productconfig);
            if (!line.Equals("XXX") && !string.IsNullOrEmpty(line))
            {
                int pos = -1;
                for (int i = 0; i < cmbLine.Items.Count; i++)
                {
                    if (line == cmbLine.GetItemText(cmbLine.Items[i])) pos = i;
                }
                cmbLine.SelectedIndex = pos;
                cbxLine.Checked = true;
            }

            // データグリットビューの更新
            dtTray = new DataTable();
            rounddownDtpHour(ref dtpRegsterDateFrom);
            rounddownDtpHour(ref dtpRegisterDateTo);
            rounddownDtpHour(ref dtpUpdateDateFrom);
            rounddownDtpHour(ref dtpUpdateDateTo);
            updateDataGridViews(dtTray, ref dgvTray, true);

            // ロールと部署による利用可能機能の制御
            // 2016/08/22 追加モードは全面禁止とする
            // 2016/12/06 追加モード復活
            if (userRole == "super") { btnAddReturn.Enabled = true; btnAddReturn.Visible = true; }
            if (userRole == "super" && txtLoginDept.Text == "QA") { btnCancelMultiTray.Enabled = true; btnCancelMultiTray.Visible = true; }
            if (userRole == "super" && txtLoginDept.Text == "QA") { btnPartiallyCancel.Enabled = true; btnPartiallyCancel.Visible = true; }
        }

        // サブプロシージャ：データグリットビューの更新。親フォームで呼び出し、親フォームの情報を引き継ぐ
        public void updateControls(string uid, string uname, string udept, string urole)
        {
            userId = uid;
            txtLoginName.Text = uname;
            txtLoginDept.Text = udept;
            cmbRegisterDept.Text = udept;
            userRole = urole;
            if (udept == "MFG") cbxRegisterDept.Checked = true;
        }

        // サブプロシージャ：データテーブルの定義
        private void defineAndReadDatatable(ref DataTable dt)
        {
            dt.Columns.Add("tray_id", typeof(string));
            dt.Columns.Add("lot", typeof(string));
            dt.Columns.Add("qty", typeof(int));
            dt.Columns.Add("register_date", typeof(DateTime));
            dt.Columns.Add("rg_dept", typeof(string));
            dt.Columns.Add("update_date", typeof(DateTime));
            dt.Columns.Add("up_dept", typeof(string));
            dt.Columns.Add("cancel_date", typeof(DateTime));
            dt.Columns.Add("cl_dept", typeof(string));
            dt.Columns.Add("multi_lot", typeof(string));
            dt.Columns.Add("pack_id", typeof(string));
        }

        // サブプロシージャ：データグリットビューの更新
        public void updateDataGridViews(DataTable dt, ref DataGridView dgv, bool load)
        {
            DateTime registerDateFrom = dtpRegsterDateFrom.Value;
            DateTime registerDateTo = dtpRegisterDateTo.Value.AddDays(1);
            string registerDept = cmbRegisterDept.Text;
            DateTime updateDateFrom = dtpUpdateDateFrom.Value;
            DateTime updateDateTo = dtpUpdateDateTo.Value.AddDays(1);
            string updateDept = cmbUpdateDept.Text;
            string trayId = txtTrayId.Text;
            string line = cmbLine.Text;
            string shift = txtShift.Text;
            string moduleId = txtModuleId.Text;
            string lot = txtLot.Text;

            bool b_registerDateFrom = cbxRegisterDateFrom.Checked;
            bool b_registerDateTo = cbxRegisterDateTo.Checked;
            bool b_registerDept = cbxRegisterDept.Checked;
            bool b_updateDateFrom = cbxUpdateDateFrom.Checked;
            bool b_updateDateTo = cbxUpdateDateTo.Checked;
            bool b_updateDept = cbxUpdateDept.Checked;
            bool b_trayId = cbxTrayId.Checked;
            bool b_line = cbxLine.Checked;
            bool b_shift = cbxShift.Checked;
            bool b_multiLot = cbxMultiLot.Checked;
            bool b_hideCancel = cbxHideCancel.Checked;
            bool b_moduleId = cbxModuleId.Checked;
            bool b_lot = cbxLot.Checked;

            // ユーザーがモジュールＩＤを検索条件として指定した場合は、個別のＳＱＬ文を使用する
            //string sqlX = "select tray_id, lot, qty, register_date, rg_dept, update_date, up_dept, cancel_date, cl_dept, multi_lot, pack_id from t_tray " + 
            //              "where tray_id in (select tray_id from t_module where module_id like '" + moduleId + "%')";
            string sqlX = "select tray_id, lot, qty, register_date, rg_dept, update_date, up_dept, cancel_date, cl_dept, multi_lot, pack_id from t_tray " +
                          "where tray_id in (select tray_id from t_module where module_id ='" + moduleId + "');";

            // ユーザーがロットＤを検索条件として指定した場合は、個別のＳＱＬ文を使用する
            string sqlY = "select tray_id, lot, qty, register_date, rg_dept, update_date, up_dept, cancel_date, cl_dept, multi_lot, pack_id from t_tray " +
                          "where tray_id in (select tray_id from t_module where lot like '" + lot + "%')";

            // モジュールＩＤ・ロットＩＤが選択されなかった場合は、その他全ての選択条件を、ＳＱＬ文に盛り込む
            string sql1 = "select tray_id, lot, qty, register_date, rg_dept, update_date, up_dept, cancel_date, cl_dept, multi_lot, pack_id from t_tray where ";

            bool[] cr = {                                      true,
                                                               true,
                          registerDept == string.Empty ? false : true,
                                                               true,
                                                               true,
                          updateDept == string.Empty ? false : true,
                          trayId     == string.Empty ? false : true,
                          line       == string.Empty ? false : true,
                          shift      == string.Empty ? false : true,
                                                               true,
                                                               true};

            bool[] ck = { b_registerDateFrom,
                          b_registerDateTo,
                          b_registerDept,
                          b_updateDateFrom,
                          b_updateDateTo,
                          b_updateDept,
                          b_trayId,
                          b_line,
                          b_shift,
                          b_multiLot,
                          b_hideCancel};

            string sql2 = (!(cr[0] && ck[0]) ? string.Empty : "register_date >= '" + registerDateFrom + "' AND ") +
                          (!(cr[1] && ck[1]) ? string.Empty : "register_date < '" + registerDateTo + "' AND ") +
                          (!(cr[2] && ck[2]) ? string.Empty : "rg_dept = '" + registerDept + "' AND ") +
                          (!(cr[3] && ck[3]) ? string.Empty : "update_date >= '" + updateDateFrom + "' AND ") +
                          (!(cr[4] && ck[4]) ? string.Empty : "update_date < '" + updateDateTo + "' AND ") +
                          (!(cr[5] && ck[5]) ? string.Empty : "up_dept = '" + updateDept + "' AND ") +
                          (!(cr[6] && ck[6]) ? string.Empty : "tray_id like '%" + trayId + "%' AND ") +
                          (!(cr[7] && ck[7]) ? string.Empty : "line = '" + line + "' AND ") +
                          (!(cr[8] && ck[8]) ? string.Empty : "shift like '%" + shift + "%' AND ") +
                          (!(cr[9] && ck[9]) ? string.Empty : "multi_lot = 'T' AND ") +
                          (!(cr[10] && ck[10]) ? string.Empty : "cancel_date is null AND ");

            bool b_all = (cr[0] && ck[0]) || (cr[1] && ck[1]) || (cr[2] && ck[2]) || (cr[3] && ck[3]) || (cr[4] && ck[4]) || (cr[5] && ck[5]) ||
                         (cr[6] && ck[6]) || (cr[7] && ck[7]) || (cr[8] && ck[8]) ||
                         (moduleId != string.Empty && b_moduleId) || (lot != string.Empty && b_lot);

            if (!b_all)
            {
                MessageBox.Show("Please select at least one check box and fill the criteria.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 5) + " order by tray_id";
            string sql4 = string.Empty;
            if (moduleId != string.Empty && b_moduleId) sql4 = sqlX;
            else if (lot != string.Empty && b_lot) sql4 = sqlY;
            else sql4 = sql3;
            System.Diagnostics.Debug.Print(sql4);

            // ＳＱＬ結果を、ＤＴＡＡＴＡＢＬＥへ格納
            dt.Clear();
            TfSQL tf = new TfSQL();
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql4, ref dt);

            // データグリットビューへＤＴＡＡＴＡＢＬＥを格納
            dgv.DataSource = dt;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // グリットビュー右端にボタンを追加（初回のみ）
            if (load) addButtonsToDataGridView(dgv);

            //行ヘッダーに行番号を表示する
            for (int i = 0; i < dgv.Rows.Count; i++) dgv.Rows[i].HeaderCell.Value = (i + 1).ToString();

            //行ヘッダーの幅を自動調節する
            dgv.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            // 一番下の行を表示する
            if (dgv.Rows.Count != 0) dgv.FirstDisplayedScrollingRowIndex = dgv.Rows.Count - 1;
         }
        
        // サブサブプロシージャ：グリットビュー右端にボタンを追加
        private void addButtonsToDataGridView(DataGridView dgv)
        {
            openTray = new DataGridViewButtonColumn();
            openTray.HeaderText = string.Empty;
            openTray.Text = "Open";
            openTray.UseColumnTextForButtonValue = true;
            openTray.Width = 80;
            dgv.Columns.Add(openTray);
        }

        // 検索ボタン押下、実際はグリットビューの更新をするだけ
        private void btnSearchTray_Click(object sender, EventArgs e)
        {
            updateDataGridViews(dtTray, ref dgvTray, false);
        }

        // グリッドビュー上のボタン押下時、モジュールフォームを閲覧モードで開く、デレゲートあり
        private void dgvBoxId_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int currentRow = int.Parse(e.RowIndex.ToString());

            if (dgvTray.Columns[e.ColumnIndex] == openTray && currentRow >= 0)
            {
                //既にfrmModuleInTray が開かれている場合は、それを閉じるよう促す
                if (TfGeneral.checkOpenFormExists("frmModuleInTray"))
                {
                    MessageBox.Show("Please close the currently open form.", "Notice",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    return;
                }

                string trayId = dgvTray["tray_id", currentRow].Value.ToString();
                DateTime trayDate = (DateTime)dgvTray["register_date", currentRow].Value;
                string shift = VBS.Right(trayId, 3);
                bool canceled = !String.IsNullOrEmpty(dgvTray["cl_dept", currentRow].Value.ToString());
                bool packed = !String.IsNullOrEmpty(dgvTray["pack_id", currentRow].Value.ToString());

                // モジュールテキストボックスが空でない、かつチェックボックスがオン、かつ検索結果が１行の場合のみ、再プリントモードを有効
                bool reprintMode = (txtModuleId.Text.Length != 0 && cbxModuleId.Checked && dtTray.Rows.Count == 1);

                frmModuleInTray fM = new frmModuleInTray();
                //子イベントをキャッチして、データグリッドを更新する
                fM.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(dtTray, ref dgvTray, false);
                    this.Focus();
                };

                fM.updateControls(trayId, trayDate, userId, txtLoginName.Text, txtLoginDept.Text, userRole, shift, false, false, canceled, packed, 1, false, reprintMode);
                fM.Show();
            }
        }

        // モジュールフォームを追加モードで開く（ＢＩＮ：Ａ）
        private void btnAddBoxId_Click(object sender, EventArgs e)
        {
            if (TfGeneral.checkOpenFormExists("frmModuleInTray")) 
            {
                MessageBox.Show("Please close the currently open form.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            frmModuleInTray fM = new frmModuleInTray();
            //子イベントをキャッチして、データグリッドを更新する
            fM.RefreshEvent += delegate(object sndr, EventArgs excp) 
            {
                updateDataGridViews(dtTray, ref dgvTray, false);
                this.Focus(); 
            };

            fM.updateControls(String.Empty, DateTime.Now, userId, txtLoginName.Text, txtLoginDept.Text, userRole, string.Empty, true, false, false, false, 1, false, false);
            fM.Show();
        }
        
        // 返品の登録：フォーム３を編集モードで開く、デレゲートあり
        private void btnAddReturn_Click(object sender, EventArgs e)
        {
            // 2016/08/22 追加モードは全面禁止とする
            // return;
            // 2016/12/06 追加モード復活
            if (TfGeneral.checkOpenFormExists("frmModuleInTray"))
            {
                MessageBox.Show("Please close the currently open form.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            frmModuleInTray fM = new frmModuleInTray();
            //子イベントをキャッチして、データグリッドを更新する
            fM.RefreshEvent += delegate (object sndr, EventArgs excp)
            {
                updateDataGridViews(dtTray, ref dgvTray, false);
                this.Focus();
            };

            fM.updateControls(String.Empty, DateTime.Now, userId, txtLoginName.Text, txtLoginDept.Text, userRole, string.Empty, true, true, false, false, 1, false, false);
            fM.Show();
        }

        //frmTrayを閉じる際、非表示になっている親フォームfrmLoginを表示する
        private void frmTray_FormClosed(object sender, FormClosedEventArgs e)
        {
            //親フォームfrmLoginを閉じるよう、デレゲートイベントを発生させる
            this.RefreshEvent(this, new EventArgs());
        }

        // サブサブプロシージャ：ＤＡＴＥＴＩＭＥＰＩＣＫＥＲの時間以下を切り下げる
        private void rounddownDtpHour(ref DateTimePicker dtp)
        {
            DateTime dval = dtp.Value;
            int hour = dval.Hour;
            int minute = dval.Minute;
            int second = dval.Second;
            int millisecond = dval.Millisecond;
            dtp.Value = dval.AddHours(-hour).AddMinutes(-minute).AddSeconds(-second).AddMilliseconds(-millisecond);
        }

        // データグリッドビューのダブルクリック時、データをエクセルへエクスポート
        private void dgvTray_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)dgvTray.DataSource;
            ExcelClass xl = new ExcelClass();
            xl.ExportToExcel(dt);
        }

        // フォームfrmModuleInTrayが開かれていないことを確認してから、閉じる
        private void btnClose_Click(object sender, EventArgs e)
        {
            if (TfGeneral.checkOpenFormExists("frmModuleInTray"))
            {
                MessageBox.Show("Please close the currently open form.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }
            Close();
        }

        // ＱＡのスーパーユーザーに限って、複数トレーの一括キャンセルができる（パック後は不可）
        private void btnCancelMultiTray_Click(object sender, EventArgs e)
        {
            if (dgvTray.Rows.Count <= 0) return;

            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
            if (dgvTray.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2)
            {
                MessageBox.Show("Please select only tray id column.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }
            // 列名が正しい場合のみ、処理を行う
            if (dgvTray.Columns[dgvTray.CurrentCell.ColumnIndex].Name != "tray_id")
            {
                MessageBox.Show("Please select only tray id column.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            // パック済みトレーがないか、確認する
            foreach (DataGridViewCell cell in dgvTray.SelectedCells)
            {
                if (dgvTray["pack_id", cell.RowIndex].Value.ToString() != string.Empty ||
                    dgvTray["cl_dept", cell.RowIndex].Value.ToString() != string.Empty)
                {
                    MessageBox.Show(cell.Value.ToString() + " has been packed or canceled already.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    return;
                }
            }

            // 本当に削除してよいか、２重で確認する。
            DialogResult result1 = MessageBox.Show("Do you really want to cancel this tray?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result1 == DialogResult.No) return;
            DialogResult result2 = MessageBox.Show("Are you really sure? Please select NO if you are not sure.", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) return;

            // 選択セルを配列に格納、およびメッセージ用文字列を生成し、一括キャンセルＳＱＬコマンドを実行
            string[] traylist = { };
            string message = string.Empty;
            int i = 0;
            foreach (DataGridViewCell cell in dgvTray.SelectedCells)
            {
                if (dgvTray["cl_dept", cell.RowIndex].Value.ToString() == string.Empty)
                {
                    i += 1;
                    Array.Resize(ref traylist, i);
                    traylist[i -1] = cell.Value.ToString();
                    message = message + Environment.NewLine + cell.Value.ToString();
                }
            }
            if (message == string.Empty)
            {
                MessageBox.Show("No tray ID was selected.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            TfSQL tf = new TfSQL();
            bool res = tf.sqlMultipleCancelModuleInTray(traylist, txtLoginDept.Text, txtLoginName.Text);

            if (res)
            {
                //本フォームのデータグリットビュー更新
                dtTray.Clear();
                updateDataGridViews(dtTray, ref dgvTray, false);
                MessageBox.Show("The following " + i + " tray IDs were canceled: " + message, "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Cancel process was not successful.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ＱＡのスーパーユーザーは、トレーの一部モジュールのキャンセルができる
        private void btnPartiallyCancel_Click(object sender, EventArgs e)
        {
            if (dtTray.Rows.Count == 0) return;
            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
            if (dgvTray.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2 ||
                dgvTray.Rows.GetRowCount(DataGridViewElementStates.Selected) >= 2 || dgvTray.CurrentCell.ColumnIndex != 0)
            {
                MessageBox.Show("Please select only one tray id.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }
            //既にfrmModuleInTray が開かれている場合は、それを閉じるよう促す
            if (TfGeneral.checkOpenFormExists("frmModuleInTray"))
            {
                MessageBox.Show("Please close the currently open form.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int currentRow = dgvTray.CurrentCell.RowIndex;
            string trayId = dgvTray["tray_id", currentRow].Value.ToString();
            DateTime trayDate = (DateTime)dgvTray["register_date", currentRow].Value;
            string shift = VBS.Right(trayId, 3);
            bool canceled = !String.IsNullOrEmpty(dgvTray["cl_dept", currentRow].Value.ToString());
            bool packed = !String.IsNullOrEmpty(dgvTray["pack_id", currentRow].Value.ToString());

            if (canceled || packed)
            {
                MessageBox.Show("The tray id is already canceled or packed.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            frmModuleInTray fM = new frmModuleInTray();
            //子イベントをキャッチして、データグリッドを更新する
            fM.RefreshEvent += delegate (object sndr, EventArgs excp)
            {
                updateDataGridViews(dtTray, ref dgvTray, false);
                this.Focus();
            };

            fM.updateControls(trayId, trayDate, userId, txtLoginName.Text, txtLoginDept.Text, userRole, shift, false, false, canceled, packed, 1, true, false);
            fM.Show();
        }

        // ユーザーマスタ登録フォームの起動
        private void txtLoginDept_DoubleClick(object sender, EventArgs e)
        {
            // スーパーユーザーが使用中の場合のみ、ユーザーマスター変更フォームを開く
            if (userRole != "super") return;
            if (TfGeneral.checkOpenFormExists("frmMasterUser")) return;
            frmMasterUser fU = new frmMasterUser(txtLoginDept.Text);
            fU.Show();
        }

        // ユーザーマスタ登録フォームの起動
        private void txtLoginName_DoubleClick(object sender, EventArgs e)
        {
            // スーパーユーザーが使用中の場合のみ、ユーザーマスター変更フォームを開く
            if (userRole != "super") return;
            if (TfGeneral.checkOpenFormExists("frmMasterUser")) return;
            frmMasterUser fU = new frmMasterUser(txtLoginDept.Text);
            fU.Show();
        }

        // 設定テキストファイルの読み込み
        private string readIni(string s, string k, string cfs)
        {
            StringBuilder retVal = new StringBuilder(255);
            string section = s;
            string key = k;
            string def = String.Empty;
            int size = 255;
            int strref = GetPrivateProfileString(section, key, def, retVal, size, cfs);
            return retVal.ToString();
        }
        // Windows API をインポート
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filepath);


        //閉じるボタンやショートカットでの終了を許さない
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x112;
            const long SC_CLOSE = 0xF060L;
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & 0xFFF0L) == SC_CLOSE) { return; }
            base.WndProc(ref m);
        }
    }
}