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
using System.Reflection;

namespace TrayGuard
{
    public partial class frmPallet : Form
    {
        //親フォームfrmLoginへ、イベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        //データグリッドビュー用ボタン
        DataGridViewButtonColumn openCarton;

        //その他非ローカル変数
        DataTable dtPallet;
        string userRole;
        string userId;

        // コンストラクタ
        public frmPallet()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void frmPallet_Load(object sender, EventArgs e)
        {
            this.Text = this.Text + " " + Assembly.GetExecutingAssembly().GetName().Version;
            // フォームの場所を指定
            this.Left = 20;
            this.Top = 10;

            dtPallet = new DataTable();
            rounddownDtpHour(ref dtpRegsterDateFrom);
            rounddownDtpHour(ref dtpRegisterDateTo);
            updateDataGridViews(dtPallet, ref dgvPallet, true);

            // インボイス番号を編集できるのは、管理ユーザーのみ
            btnUpdateInvoice.Enabled = userRole == "super" ? true : false;
        }

        // サブプロシージャ：データグリットビューの更新。親フォームで呼び出し、親フォームの情報を引き継ぐ
        public void updateControls(string uid, string uname, string udept, string urole)
        {
            userId = uid;
            txtLoginName.Text = uname;
            txtLoginDept.Text = udept;
            userRole = urole;
        }

        // サブプロシージャ：データテーブルの定義
        private void defineAndReadDatatable(ref DataTable dt)
        {
            dt.Columns.Add("pallet_id", typeof(string));
            dt.Columns.Add("lot", typeof(string));
            dt.Columns.Add("l_cnt", typeof(int));
            dt.Columns.Add("m_qty", typeof(int));
            dt.Columns.Add("batch", typeof(string));
            dt.Columns.Add("register_date", typeof(DateTime));
            dt.Columns.Add("rg_user", typeof(string));
            dt.Columns.Add("cancel_date", typeof(DateTime));
            dt.Columns.Add("cl_user", typeof(string));
            dt.Columns.Add("invoice_no", typeof(string));
        }

        // サブプロシージャ：データグリットビューの更新
        public void updateDataGridViews(DataTable dt, ref DataGridView dgv, bool load)
        {
            DateTime registerDateFrom = dtpRegsterDateFrom.Value;
            DateTime registerDateTo = dtpRegisterDateTo.Value.AddDays(1);
            string invoiceNo = txtInvoiceNo.Text;
            string palletId = txtPalletId.Text;
            string lot = txtLot.Text;
            string cartonId = txtCartonId.Text;
            string batch= txtBatch.Text;

            bool b_registerDateFrom = cbxRegisterDateFrom.Checked;
            bool b_registerDateTo = cbxRegisterDateTo.Checked;
            bool b_invoiceNo = cbxInvoiceNo.Checked;
            bool b_palletId = cbxPalletId.Checked;
            bool b_lot = cbxLot.Checked;
            bool b_cartonId = cbxCartonId.Checked;
            bool b_batch = cbxBatch.Checked;
            bool b_multi_lot = cbxMultiLot.Checked;
            bool b_hideCancel = cbxHideCancel.Checked;

            // ユーザーがパックＩＤを検索条件として指定した場合は、個別のＳＱＬ文を使用する
            string sqlX = "select pallet_id, lot, l_cnt, m_qty, batch, register_date, rg_user, cancel_date, cl_user, invoice_no from t_pallet " +
                          "where pallet_id in (select pallet_id from t_carton where carton_id like '" + cartonId + "%')";

            // ユーザーが選択した検索条件を、ＳＱＬ文に反映する
            string sql1 = "select pallet_id, lot, l_cnt, m_qty, batch, register_date, rg_user, cancel_date, cl_user, invoice_no from t_pallet where ";

            bool[] cr = {                                     true,
                                                              true,
                          invoiceNo == string.Empty ? false : true,
                          palletId  == string.Empty ? false : true,
                          lot       == string.Empty ? false : true,
                          cartonId    == string.Empty ? false : true,
                          batch     == string.Empty ? false : true,
                                                              true,
                                                              true};

            bool[] ck = { b_registerDateFrom,
                          b_registerDateTo,
                          b_invoiceNo,
                          b_palletId,
                          b_lot,
                          b_cartonId,
                          b_batch,
                          b_multi_lot,
                          b_hideCancel};

            string sql2 = (!(cr[0] && ck[0]) ? string.Empty : "register_date >= '" + registerDateFrom + "' AND ") +
                          (!(cr[1] && ck[1]) ? string.Empty : "register_date < '" + registerDateTo + "' AND ") +
                          (!(cr[2] && ck[2]) ? string.Empty : "invoice_no like '%" + invoiceNo + "%' AND ") +
                          (!(cr[3] && ck[3]) ? string.Empty : "pallet_id like '%" + palletId + "%' AND ") +
                          (!(cr[4] && ck[4]) ? string.Empty : "lot like '%" + lot + "%' AND ") +
                          (!(cr[5] && ck[5]) ? string.Empty : "carton_id like '%" + cartonId + "%' AND ") +
                          (!(cr[6] && ck[6]) ? string.Empty : "batch = '" + batch + "' AND ") +
                          (!(cr[7] && ck[7]) ? string.Empty : "l_cnt >= 2 AND ") +
                          (!(cr[8] && ck[8]) ? string.Empty : "cancel_date is null AND ");

            bool b_all = (cr[0] && ck[0]) || (cr[1] && ck[1]) || (cr[2] && ck[2]) || (cr[3] && ck[3]) || (cr[4] && ck[4]) || (cr[5] && ck[5]) || (cr[6] && ck[6]);

            if (!b_all)
            {
                MessageBox.Show("Please select at least one check box and fill the criteria.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 5) + " order by pallet_id";
            string sql4 = string.Empty;
            if (cartonId != string.Empty && b_cartonId) sql4 = sqlX;
            else sql4 = sql3;
            System.Diagnostics.Debug.Print(sql4);

            // ＳＱＬ結果を、ＤＴＡＡＴＡＢＬＥへ格納
            dt.Clear();
            TfSQL tf = new TfSQL();
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql4, ref dt);

            // データグリットビューへＤＴＡＡＴＡＢＬＥを格納
            dgv.DataSource = dt;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

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
            openCarton = new DataGridViewButtonColumn();
            openCarton.HeaderText = string.Empty;
            openCarton.Text = "Open";
            openCarton.UseColumnTextForButtonValue = true;
            openCarton.Width = 80;
            dgv.Columns.Add(openCarton);
        }

        // 検索ボタン押下、実際はグリットビューの更新をするだけ
        private void btnSearchPack_Click(object sender, EventArgs e)
        {
            updateDataGridViews(dtPallet, ref dgvPallet, false);
        }

        // グリッドビュー上のボタン押下時、モジュールフォームを閲覧モードで開く、デレゲートあり
        private void dgvBoxId_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int currentRow = int.Parse(e.RowIndex.ToString());

            if (dgvPallet.Columns[e.ColumnIndex] == openCarton && currentRow >= 0)
            {
                //既にfrmModuleInTray が開かれている場合は、それを閉じるよう促す
                if (TfGeneral.checkOpenFormExists("frmCartonOnPallet"))
                {
                    MessageBox.Show("Please close the currently open form.", "Notice",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    return;
                }

                string palletId = dgvPallet["pallet_id", currentRow].Value.ToString();
                DateTime palletDate = (DateTime)dgvPallet["register_date", currentRow].Value;
                string batch = txtBatch.Text;
                bool canceled = !String.IsNullOrEmpty(dgvPallet["cl_user", currentRow].Value.ToString());
                bool invoiced = !String.IsNullOrEmpty(dgvPallet["invoice_no", currentRow].Value.ToString());

                // モジュールテキストボックスが空でない、かつチェックボックスがオン、かつ検索結果が１行の場合のみ、再プリントモードを有効
                bool reprintMode = (txtCartonId.Text.Length != 0 && cbxCartonId.Checked && dtPallet.Rows.Count == 1);

                frmCartonOnPallet fP = new frmCartonOnPallet();
                //子イベントをキャッチして、データグリッドを更新する
                fP.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(dtPallet, ref dgvPallet, false);
                    this.Focus();
                };

                fP.updateControls(palletId, palletDate, userId, txtLoginName.Text, txtLoginDept.Text, userRole, batch, false, false, canceled, invoiced, reprintMode);
                fP.Show();
            }
        }

        // frmTrayInPack を追加モードで開く、デレゲートあり
        private void btnAddBoxId_Click(object sender, EventArgs e)
        {
            if (TfGeneral.checkOpenFormExists("frmCartonOnPallet")) 
            {
                MessageBox.Show("Please close the currently open form.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            frmCartonOnPallet fP = new frmCartonOnPallet();
            //子イベントをキャッチして、データグリッドを更新する
            fP.RefreshEvent += delegate(object sndr, EventArgs excp) 
            {
                updateDataGridViews(dtPallet, ref dgvPallet, false);
                this.Focus(); 
            };

            fP.updateControls(String.Empty, DateTime.Now, userId, txtLoginName.Text, txtLoginDept.Text, userRole, string.Empty, true, false, false, false, false);
            fP.Show();
        }

        //frmPalletを閉じる際、非表示になっている親フォームfrmLoginを表示する
        private void frmPallet_FormClosed(object sender, FormClosedEventArgs e)
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
        private void dgvCarton_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)dgvPallet.DataSource;
            ExcelClass xl = new ExcelClass();
            xl.ExportToExcel(dt);
        }

        // フォームfrmTrayInPackが開かれていないことを確認してから、閉じる
        private void btnClose_Click(object sender, EventArgs e)
        {
            if (TfGeneral.checkOpenFormExists("frmCartonOnPallet"))
            {
                MessageBox.Show("Please close the currently open form.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }
            Close();
        }

        // インボイスフィールドの登録とキャンセル
        private void btnUpdateInvoice_Click(object sender, EventArgs e)
        {
            if (dgvPallet.Rows.Count <= 0) return;

            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
            if (dgvPallet.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2)
            {
                MessageBox.Show("Please select only pallet id column.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }
            // 列名が正しい場合のみ、処理を行う
            if (dgvPallet.Columns[dgvPallet.CurrentCell.ColumnIndex].Name != "pallet_id")
            {
                MessageBox.Show("Please select only pallet id column.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }
       
            // キャンセル済みカートンがないか、確認する
            foreach (DataGridViewCell cell in dgvPallet.SelectedCells)
            {
                if (dgvPallet["cl_user", cell.RowIndex].Value.ToString() != string.Empty)
                {
                    MessageBox.Show(cell.Value.ToString() + " is canceled already.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    return;
                }
            }

            // 選択セルを配列に格納、およびメッセージ用文字列を生成し、インボイス一括登録ＳＱＬコマンドを実行
            string[] palletlist = { };
            string message = string.Empty;
            int i = 0;
            int topRow = 0;

            foreach (DataGridViewCell cell in dgvPallet.SelectedCells)
            {
                i += 1;
                Array.Resize(ref palletlist, i);
                palletlist[i - 1] = cell.Value.ToString();
                message = message + Environment.NewLine + cell.Value.ToString();
                if (i == 1) topRow = cell.RowIndex;
            }
            
            frmInvoiceNo fI = new frmInvoiceNo();
            //子イベントをキャッチして、データグリッドを更新する
            fI.RefreshEvent += delegate (object sndr, EventArgs excp)
            {
                updateDataGridViews(dtPallet, ref dgvPallet, false);
            };

            fI.updateControls(palletlist, message, dgvPallet["cl_user", topRow].Value.ToString());
            fI.Show();
        }

        // カートンラベルの一括プリントアウト（クリック版）
        private void btnPrintCarton_Click(object sender, EventArgs e)
        {
            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
            if (dgvPallet.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2)
            {
                MessageBox.Show("Please select only carton id column.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }
            // 列名が正しい場合のみ、処理を行う
            if (dgvPallet.Columns[dgvPallet.CurrentCell.ColumnIndex].Name != "pallet_id")
            {
                MessageBox.Show("Please select only pallet id column.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            // キャンセル済みトレーを選択している場合は、印刷しない
            foreach (DataGridViewCell cell in dgvPallet.SelectedCells)
            {
                if (dgvPallet["cl_user", cell.RowIndex].Value.ToString() != string.Empty)
                {
                    MessageBox.Show(cell.Value.ToString() + " has been canceled already.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    return;
                }
            }

            // 選択セルを配列に格納、およびメッセージ用文字列を生成し、一括プリントを行う
            string[] palletlist = { };
            string message = string.Empty;
            int i = 0;
            foreach (DataGridViewCell cell in dgvPallet.SelectedCells)
            {
                if (dgvPallet["cl_user", cell.RowIndex].Value.ToString() == string.Empty)
                {
                    i += 1;
                    Array.Resize(ref palletlist, i);
                    palletlist[i - 1] = cell.Value.ToString();
                    message = message + Environment.NewLine + cell.Value.ToString();
                }
            }
            if (message == string.Empty)
            {
                MessageBox.Show("No carton id was selected.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            TfSato tfs = new TfSato();
            for (int j = 0; j < palletlist.Length; j++)
            {
                //tfs.printStart("palletPega", palletlist[j], 0);
            }
        }

        // カートンラベルの一括プリントアウト（エンターキー版）
        private void btnPrintPallet_KeyDown(object sender, KeyEventArgs e)
        {

        }

        // 閉じるボタンやショートカットでの終了を許さない
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