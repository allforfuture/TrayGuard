using System;
using System.Data;
using System.Windows.Forms;
using System.Security.Permissions;
using System.Drawing;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;

namespace TrayGuard
{
    public partial class frmCartonAdjust : Form
    {
        //親フォームfrmTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        //その他非ローカル変数
        DataTable dtCarton;
        string palletId;
        string user;
        string mode;
        bool cartonNg;
        bool sound;
        
        
        // コンストラクタ
        public frmCartonAdjust()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void frmCartonAdjust_Load(object sender, EventArgs e)
        {
            //フォームの場所を指定
            this.Left = 450;
            this.Top = 100;

            changeFormatByMode();
            dtCarton = new DataTable();
            defineDatatable(ref dtCarton);
            updateDataGridViews(dtCarton, ref dgvCarton);
        }

        // サブプロシージャ: ＤＴの定義
        private void defineDatatable(ref DataTable dt)
        {
            dt.Columns.Add("carton_id", typeof(string));
            dt.Columns.Add("lot", typeof(string));
            dt.Columns.Add("m_qty", typeof(int));
            dt.Columns.Add("batch", typeof(string));
            dt.Columns.Add("register_date", typeof(DateTime));
            dt.Columns.Add("rg_user", typeof(string));
            dt.Columns.Add("multi_lot", typeof(string));
            dt.Columns.Add("check", typeof(string));
        }

        // サブプロシージャ：親フォームで呼び出し、親フォームの情報を、テキストボックスへ格納して引き継ぐ
        public void updateControls(string pallet, string carton, int row, string userInfo, string modeSelected)
        {
            palletId = pallet;
            txtBefore.Text = carton;
            txtRow.Text = row.ToString();
            user = userInfo;
            mode = modeSelected;
        }

        // サブプロシージャ： 置換モード、削除モード、追加モードの切り替え
        private void changeFormatByMode()
        {
            // 置換モードの場合
            if (mode == "replace")
            {
                this.Text = "Replace Carton";
                btnReplaceCarton.Visible = true;
                btnDeleteCarton.Visible = false;
                btnAddCarton.Visible = false;
                txtAfter.Enabled = true;
            }
            // 削除モードの場合
            else if (mode == "delete")
            {
                this.Text = "Delete Carton";
                btnReplaceCarton.Visible = false;
                btnDeleteCarton.Visible = true;
                btnAddCarton.Visible = false;
                txtAfter.Enabled = false;
            }
            // 追加モードの場合
            else if (mode == "add")
            {
                this.Text = "Delete Carton";
                btnReplaceCarton.Visible = false;
                btnDeleteCarton.Visible = false;
                btnAddCarton.Visible = true;
                txtAfter.Enabled = true;
            }
        }

        // サブプロシージャ：データグリットビューの更新
        private void updateDataGridViews(DataTable dt, ref DataGridView dgv)
        {
            // データグリットビューへＤＴＡＡＴＡＢＬＥを格納
            dgv.DataSource = dt;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            // テスト結果がＦＡＩＬまたはレコードなしのシリアルをマーキングする
            if (dt.Rows.Count <= 0) return;
            if (dgv["check", 0].Value.ToString() != "OK" || dgv["lot", 0].Value.ToString() == string.Empty)
            {
                for (int i = 0; i < dgv.ColumnCount; i++)
                {
                    dgv[i, 0].Style.BackColor = Color.Red;
                }
                soundAlarm();
                cartonNg = true;
            }
            else
            {
                for (int i = 0; i < dgv.ColumnCount; i++)
                {
                    dgv[i, 0].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                }
                cartonNg = false;
            }
        }

        // 変更後モジュールがスキャンされたときの処理
        private void txtAfter_KeyDown(object sender, KeyEventArgs e)
        {
            // エンターキーの場合、テキストボックスの桁数が１５桁の場合のみ、処理を行う
            if (e.KeyCode != Keys.Enter || txtAfter.Text.Length != 15) return;

            // 置換モード、または、追加モードの場合のみ、処理を行う
            if (mode == "replace" || mode == "add")
            {
                string carton = txtAfter.Text;
                string sql = "select carton_id, lot, m_qty, batch, register_date, rg_user, " +
                    "case when l_cnt >= 2 then 'T' else 'F' end as multi_lot, " +
                    "case when cancel_date is not null then to_char(cancel_date,'YYYY/MM/DD') " +
                         "when pallet_id is not null then pallet_id else 'OK' end as check " +
                    "from t_carton where carton_id='" + carton + "'";
                System.Diagnostics.Debug.Print(sql);
                DataTable dt = new DataTable();
                TfSQL tf = new TfSQL();
                tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql, ref dt);

                // テスターデータに該当がない場合でも、ユーザー認識用に表示するための処理
                dtCarton.Rows.Clear();
                DataRow dr = dtCarton.NewRow();
                dr["carton_id"] = carton;
                // テスターデータに該当がある場合の処理
                if (dt.Rows.Count != 0)
                {
                    dr["lot"] = (string)dt.Rows[0]["lot"];
                    dr["m_qty"] = (int)dt.Rows[0]["m_qty"];
                    dr["batch"] = (string)dt.Rows[0]["batch"];
                    dr["register_date"] = (DateTime)dt.Rows[0]["register_date"];
                    dr["rg_user"] = (string)dt.Rows[0]["rg_user"];
                    dr["multi_lot"] = (string)dt.Rows[0]["multi_lot"];
                    dr["check"] = (string)dt.Rows[0]["check"];
                }

                // メモリ上のテーブルにレコードを追加
                dtCarton.Rows.Add(dr);

                // データグリットビューの更新
                updateDataGridViews(dtCarton, ref dgvCarton);
            }
        }

        // 登録済みのカートンおよびその付帯情報を、ＵＰＤＡＴＥ文で置き換える
        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (cartonNg || dtCarton.Rows.Count <= 0) return;  

            string mdlBefore = txtBefore.Text;
            string mdlAfter = dtCarton.Rows[0]["carton_id"].ToString();

            // 更新処理
            TfSQL tf = new TfSQL();
            bool res = tf.sqlReplaceCartonOnPallet(palletId, txtBefore.Text, txtAfter.Text, user);
        
            if (res)
            {
                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                btnReplaceCarton.Enabled = false;
                txtAfter.Enabled = false;
                // txtRow.Text = string.Empty;
                this.Focus();
                MessageBox.Show("The replacement was successful." + Environment.NewLine + 
                    "Please re-print the pallet label.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
        }

        // 登録済みのパレットに対し、カートンを追加する
        private void btnAddCarton_Click(object sender, EventArgs e)
        {
            if (cartonNg || dtCarton.Rows.Count <= 0) return;

            string mdlAfter = dtCarton.Rows[0]["carton_id"].ToString();
            string sql = "update t_carton set pallet_id = '" + palletId + "' where carton_id = '" + mdlAfter + "'";
            System.Diagnostics.Debug.Print(sql);
            TfSQL tf = new TfSQL();
            bool res = tf.sqlExecuteNonQuery(sql, false);

            if (res)
            {
                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                btnReplaceCarton.Enabled = false;
                txtAfter.Enabled = false;
                // txtRow.Text = string.Empty;
                this.Focus();
                MessageBox.Show("The addition was successful." + Environment.NewLine +
                    "Please re-print the pallet label.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
        }

        // 登録済みのパレットに対し、選択されたカートンを削除する
        private void btnDeleteCarton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtBefore.Text)) return;

            string mdlBefore = txtBefore.Text;
            string sql = "update t_carton set pallet_id = null where carton_id = '" + mdlBefore + "'";
            System.Diagnostics.Debug.Print(sql);
            TfSQL tf = new TfSQL();
            bool res = tf.sqlExecuteNonQuery(sql, false);

            if (res)
            {
                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                btnReplaceCarton.Enabled = false;
                btnDeleteCarton.Enabled = false;
                btnAddCarton.Enabled = false;
                txtAfter.Enabled = false;
                // txtRow.Text = string.Empty;
                this.Focus();
                MessageBox.Show("The deletion was successful." + Environment.NewLine +
                    "Please re-print the pallet label.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
        }

        //閉じるボタンやショートカットでの終了を許さない
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x112;
            const long SC_CLOSE = 0xF060L;
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & 0xFFF0L) == SC_CLOSE) { return; }
            base.WndProc(ref m);
        }

        // 他のフォームとの整合性を取るため、キャンセルボタンを設ける
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //MP3ファイル（今回は警告音）を再生する
        [System.Runtime.InteropServices.DllImport("winmm.dll")]
        private static extern int mciSendString(String command,
           StringBuilder buffer, int bufferSize, IntPtr hwndCallback);

        private string aliasName = "MediaFile";

        private void soundAlarm()
        {
            string currentDir = System.Environment.CurrentDirectory;
            string fileName = currentDir + @"\warning.mp3";
            string cmd;

            if (sound)
            {
                cmd = "stop " + aliasName;
                mciSendString(cmd, null, 0, IntPtr.Zero);
                cmd = "close " + aliasName;
                mciSendString(cmd, null, 0, IntPtr.Zero);
                sound = false;
            }

            cmd = "open \"" + fileName + "\" type mpegvideo alias " + aliasName;
            if (mciSendString(cmd, null, 0, IntPtr.Zero) != 0) return;
            cmd = "play " + aliasName;
            mciSendString(cmd, null, 0, IntPtr.Zero);
            sound = true;
        }
    }
}