using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Security.Permissions;
using System.Runtime.InteropServices;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace TrayGuard
{
    public partial class frmPackInCarton : Form
    {
        //親フォームfrmTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        //データグリッドビュー用ボタン
        DataGridViewButtonColumn openPack;

        // プリント用テキストファイルの保存用フォルダを、基本設定ファイルで設定する
        string appconfig = System.Environment.CurrentDirectory + "\\info.ini";
        

        //その他、非ローカル変数
        DataTable dtPack;
        DataTable dtLot;
        bool formAddMode;
        bool formReprintMode;
        bool cartonIdCanceled;
        bool cartonIdPalleted;
        string userRole;
        int okCount;
        int capacity = 2;
        string maxLot;
        DateTime registerDate;
        bool sound;
        int position;

        // コンストラクタ
        public frmPackInCarton()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void frmPackInCarton_Load(object sender, EventArgs e)
        {
            this.Text = this.Text + " " + Assembly.GetExecutingAssembly().GetName().Version;
            // 当フォームの表示場所を指定
            this.Left = 250;
            this.Top = 20;
            if (position == 2) { this.Left = 400; this.Top = 30; }
            else if (position == 3) { this.Left = 450; this.Top = 40; }
            else if (position == 4) { this.Left = 500; this.Top = 50; }

            // 各種処理用のテーブルを生成し、データを読み込む
            dtPack = new DataTable();
            definePackTable(ref dtPack);
            if (!formAddMode) readPackInfo(ref dtPack);

            // グリットビューの更新
            updateDataGridViews(dtPack, ref dgvPack, true);

            // 追加モード、閲覧モードの切り替え
            changeMode();
        }

        // サブプロシージャ： 追加モード、閲覧モードの切り替え
        private void changeMode()
        {
            // 追加モードの場合
            if (formAddMode)
            {
                txtPack.Enabled = true;
                btnRegisterCarton.Text = "Register Carton";
                btnRegisterCarton.Enabled = false;
                btnCancelCarton.Enabled = false;
                btnClose.Enabled = true;
                btnDeleteSelection.Enabled = true;
            }
            // 閲覧モードの場合
            else
            {
                txtPack.Enabled = false;
                btnRegisterCarton.Text = "Re-Print Label";
                btnRegisterCarton.Enabled = cartonIdCanceled ? false : true;
                // 2016.09.24 FUJII  再プリントモードを追加（frmPackにて、トレーＩＤをキーとして、パックＩＤを特定した場合のみ、再印刷可能）
                //btnRegisterCarton.Enabled = cartonIdCanceled ? false : ((formReprintMode && userRole == "super") ? true : false);
                // 2017.01.18 FUJII  再プリントモードを変更（ユーザー制限、再印刷モード制限を廃止）
                btnRegisterCarton.Enabled = cartonIdCanceled ? false : true;

                btnCancelCarton.Enabled = false;
                btnClose.Enabled = true;
                btnDeleteSelection.Enabled = false;
                if (userRole == "super") btnCancelCarton.Enabled = (cartonIdCanceled || cartonIdPalleted) ? false : true;
            }
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

        // サブプロシージャ：親フォームで呼び出し、親フォームの情報を、テキストボックスへ格納して引き継ぐ
        public void updateControls(string pid, DateTime cdate, string uid, string uname, string udept, string urole, string shift, 
            bool addMode, bool returnMode, bool canceled, bool palleted, int pos, bool reprintMode)
        {
            txtCarton.Text = pid;
            dtpRegisterDate.Value = cdate;
            txtLoginName.Text = uname;
            txtLoginDept.Text = udept;
            userRole = urole;
            formAddMode = addMode;
            cartonIdCanceled = canceled;
            cartonIdPalleted = palleted;
            position = pos;
            formReprintMode = reprintMode;
        }

        // サブプロシージャ：データテーブルの定義
        private void definePackTable(ref DataTable dt)
        {
            dt.Columns.Add("pack_id", typeof(string));
            dt.Columns.Add("lot", typeof(string));
            dt.Columns.Add("m_qty", typeof(int));
            dt.Columns.Add("batch", typeof(string));
            dt.Columns.Add("register_date", typeof(DateTime));
            dt.Columns.Add("rg_user", typeof(string));
            dt.Columns.Add("multi_lot", typeof(string));
            dt.Columns.Add("check", typeof(string));
        }

        // サブプロシージャ：ＤＢからデータテーブルへの読み込み
        private void readPackInfo(ref DataTable dt)
        {
            dt.Rows.Clear();
            string sql = "select pack_id, lot, m_qty, batch, register_date, rg_user, " +
                "case when l_cnt >= 2 then 'T' else 'F' end as multi_lot, " +
                "'OK' as check from t_pack where carton_id='" + txtCarton.Text + "'";
            TfSQL tf = new TfSQL();
            System.Diagnostics.Debug.Print(sql);
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql, ref dt);
        }


        // サブプロシージャ：データグリットビューの更新
        private void updateDataGridViews(DataTable dt1, ref DataGridView dgv1, bool load)
        {
            // データグリットビューへＤＴＡＡＴＡＢＬＥを格納
            dgv1.DataSource = dt1;
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            // テスト結果がＦＡＩＬまたはレコードなしのシリアルをマーキングする 
            if(formAddMode) colorViewForCanceledAndCartoned(ref dgv1);

            // 重複レコード、および１セル２重入力をマーキングする
            if (formAddMode) colorViewForDuplicatePack(ref dgv1);

            // グリットビュー右端にボタンを追加（初回のみ）
            if (load && !formAddMode) addButtonsToDataGridView(dgv1);

            //行ヘッダーに行番号を表示する
            for (int i = 0; i < dgv1.Rows.Count; i++) dgv1.Rows[i].HeaderCell.Value = (i + 1).ToString();

            //行ヘッダーの幅を自動調節する
            dgv1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            // 現在の一時登録件数を変数へ保持する
            okCount = getOkCount(dt1);
            txtOkCount.Text = okCount.ToString() + "/" + capacity.ToString();

            // ロット集計グリッドビューを更新し、数量の最も多いロットを保持する
            maxLot = updateLotSummary(dt1);

            // レコード数とトレイ容量による、ボタンとコントロールの制御、（追加モードの場合）
            if (formAddMode)
            {
                if (okCount == capacity)
                {
                    txtPack.Enabled = false;
                    btnRegisterCarton.Enabled = true;
                    btnDeleteSelection.Enabled = true;
                }
                else
                {
                    txtPack.Enabled = true;
                    btnRegisterCarton.Enabled = false;
                    btnDeleteSelection.Enabled = true;
                    txtPack.SelectAll(); // 連続スキャン用
                }
            }

            // バッチテキストボックスへ、グリッドビューのバッチを表示する（複数バッチ混入の場合は、警告する）
            if(dt1.Rows.Count >= 1) txtBatch.Text = dt1.Rows[0]["batch"].ToString();
            if (dt1.Rows.Count == 2 && dt1.Rows[0]["batch"].ToString() != dt1.Rows[dt1.Rows.Count-1]["batch"].ToString())
                { txtBatch.Text = "Error"; soundAlarm();} 
        }

        // サブサブプロシージャ：グリットビュー右端にボタンを追加
        private void addButtonsToDataGridView(DataGridView dgv)
        {
            openPack = new DataGridViewButtonColumn();
            openPack.HeaderText = string.Empty;
            openPack.Text = "Open";
            openPack.UseColumnTextForButtonValue = true;
            openPack.Width = 80;
            dgv.Columns.Add(openPack);
        }

        // サブプロシージャ：シリアル番号重複なしのＰＡＳＳ個数を取得する
        private int getOkCount(DataTable dt)
        {
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "pack_id", "check" });
            DataRow[] dr = distinct.Select("check = 'OK'");
            return dr.Length;
        }

        // サブプロシージャ：ロット集計グリッドビューを更新し、数量の最も多いロット番号を返す
        public string updateLotSummary(DataTable dt)
        {
            if (dt.Rows.Count <= 0)
            {
                dtLot = new DataTable();
                dgvLotSummary.DataSource = dtLot;
                return string.Empty;
            }

            // 各パックに含まれる、トレーＩＤリストを取得する
            string sql1 = "select tray_id from t_tray where pack_id in (";
            string sql2 = string.Empty;
            var query1 = dt.AsEnumerable().Select(row => new { pack_id = row.Field<string>("pack_id") });
            foreach (var q in query1) sql2 += "'" + q.pack_id + "', ";
            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
            System.Diagnostics.Debug.Print(sql3);
            DataTable dtTemp1 = new DataTable();
            TfSQL tf = new TfSQL();
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql3, ref dtTemp1);

            // 上記で取得したトレーＩＤが持つモジュールから、ロット集計を作成する
            string sql4 = "select lot, count(lot) as qty from t_module where tray_id in (";
            string sql5 = string.Empty;
            var query2 = dtTemp1.AsEnumerable().Select(row => new { tray_id = row.Field<string>("tray_id") });
            foreach (var q in query2) sql5 += "'" + q.tray_id + "', ";
            string sql6 = sql4 + VBS.Left(sql5, sql5.Length - 2) + ") group by lot order by qty desc, lot ";
            System.Diagnostics.Debug.Print(sql6);
            DataTable dtTemp2 = new DataTable();
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql6, ref dtTemp2);

            // ロット集計表示テーブルに、一時テーブルの情報を移す
            dgvLotSummary.DataSource = null;
            dgvLotSummary.Refresh();
            dtLot = new DataTable();
            var query3 = dtTemp2.AsEnumerable().Select(r => new {lot = r.Field<string>("lot"), qty = r.Field<Int64>("qty") });
            // 列の追加
            foreach (var q in query3) dtLot.Columns.Add(q.lot, typeof(int));
            dtLot.Columns.Add("total", typeof(int));
            // 行の追加
            dtLot.Rows.Add(); 
            foreach (var q in query3) dtLot.Rows[0][q.lot] = q.qty;
            dtLot.Rows[0]["total"] = query3.Sum(a => a.qty);
 
            dgvLotSummary.DataSource = dtLot;
            return query3.First().lot;
        }

        // シリアルがスキャンされた時の処理            
        private void txtModuleId_KeyDown(object sender, KeyEventArgs e)
        {
            // エンターキーの場合、テキストボックスの桁数が１５桁の場合のみ、処理を行う
            if (e.KeyCode != Keys.Enter || txtPack.Text.Length != 15) return;

            string pack = txtPack.Text;
            string sql = "select pack_id, lot, m_qty, batch, register_date, rg_user, " +
                "case when l_cnt >= 2 then 'T' else 'F' end as multi_lot, " +
                "case when cancel_date is not null then to_char(cancel_date,'YYYY/MM/DD') " +
                     "when carton_id is not null then carton_id else 'OK' end as check " +
                "from t_pack where pack_id='" + pack + "'";
            System.Diagnostics.Debug.Print(sql);
            DataTable dt = new DataTable();
            TfSQL tf = new TfSQL();
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql, ref dt);

            // テスターデータに該当がない場合でも、ユーザー認識用に表示するための処理
            DataRow dr = dtPack.NewRow();
            dr["pack_id"] = pack;
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
            dtPack.Rows.Add(dr);

            // データグリットビューの更新
            updateDataGridViews(dtPack, ref dgvPack, false);
        }

        // 登録ボタン押下時、各種確認、ボックスＩＤの発行、シリアルの登録、バーコードラベルのプリントを行う
        private void btnRegisterTray_Click(object sender, EventArgs e)
        {
            //閲覧モードの場合は、プリントアウト
            if (!formAddMode)
            {
                string carton = txtCarton.Text;
                TfSato tfs = new TfSato();

                // ビンＡか、ビンＢか、ユーザーに選択させる
                //DialogResult binResult = MessageBox.Show("Please click YES for Bin A, NO for Bin B.", "Print Option",
                //    MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                //string userBin = (binResult == DialogResult.Yes) ? "A" : "B";
                string userBin = txtBatch.Text;

                // ペガトロン用か否か、ユーザーに選択させる
                DialogResult result = MessageBox.Show("Do you print Pegatoron label also?", "Print Option", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes)
                {
                //tfs.printStart("packCartonInternal", carton, dtLot, txtBatch.Text, dtpRegisterDate.Value, "Carton", "Fox", 2, userBin);
                    tfs.printStart("packCartonPega", carton, dtLot, txtBatch.Text, dtpRegisterDate.Value, "Carton", "Pega", 1, userBin);
                }
                else
                {
                    tfs.printStart("packCartonInternal", carton, dtLot, txtBatch.Text, dtpRegisterDate.Value, "Carton", "Fox", 2, userBin);
                }
                return;
            }

            //複数のバッチが混入している場合は、警告する
            if (txtBatch.Text == "Error")
            {
                MessageBox.Show("You can not register 2 batchs in a carton.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 登録処理中は、閉じるボタンを除き、すべてのコントロールを無効にする
            txtPack.Enabled = false;
            btnRegisterCarton.Enabled = false;
            btnDeleteSelection.Enabled = false;

            //カートンＩＤの新規採番
            // 2016.08.22 FUJII  カートンＩＤの新規採番プロシージャ「getNewCartonId」を、トランザクション処理バージョンへ変更
            //string cartonNew = getNewCartonId(txtBatch.Text, txtLoginName.Text);
            TfSQL tf = new TfSQL();
            string cartonNew = tf.sqlGetNewCartonId(maxLot, txtBatch.Text, txtLoginName.Text, dtLot, ref registerDate);
            if (cartonNew == string.Empty)
            {
                MessageBox.Show("An error happened in the carton id issuing process.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //パックテーブルのフィールドＣＡＲＴＯＮＩＤ、その他を更新する
            bool res = tf.sqlMultipleUpdatePackInCarton(dtPack, cartonNew);

            if (res)
            {
                //登録済みの状態を表示
                txtCarton.Text = cartonNew;
                dtpRegisterDate.Value = registerDate;

                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                this.Focus();
                MessageBox.Show("Carton ID: " + cartonNew + Environment.NewLine +
                    "and its packs were registered.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //ユーザーによるメッセージボックス確認後の処理
                txtCarton.Text = String.Empty;
                txtPack.Text = String.Empty;
                dtPack.Clear();
                updateDataGridViews(dtPack, ref dgvPack, false);
            }
        }

        // スーパーユーザーに限り、登録済カートンをキャンセルできる（インボイス処理後は不可）
        private void btnCancelCarton_Click(object sender, EventArgs e)
        {
            // 本当に削除してよいか、２重で確認する。
            DialogResult result1 = MessageBox.Show("Do you really want to cancel this tray?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result1 == DialogResult.No) return;

            DialogResult result2 = MessageBox.Show("Are you really sure? Please select NO if you are not sure.", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) return;

            // キャンセルの実行
            string cartonId = txtCarton.Text;
            TfSQL tf = new TfSQL();
            bool res = tf.sqlCancelPackInCarton(dtPack, cartonId, txtLoginName.Text);
            if (res)
            {
                //本フォームのデータグリットビュー更新
                dtPack.Clear();
                updateDataGridViews(dtPack, ref dgvPack, false);

                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                this.Focus();
                MessageBox.Show("Carton ID " + cartonId + " and its packs were canceled.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // メッセージボックスの確認後、閉じる
                Close();
            }
            else
            {
                MessageBox.Show("Cancel process was not successful.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 当フォームのクローズ（他のフォームと形式をあわせているだけで、意味はない）
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        // サブプロシージャ：キャンセル済みまたはカートン済みのパックをマーキングする
        private void colorViewForCanceledAndCartoned(ref DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; ++i)
            {
                if (dgv["check", i].Value.ToString() != "OK")
                {
                    dgv["check", i].Style.BackColor = Color.Red;
                    soundAlarm();
                }
                else
                {
                    dgv["check", i].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                }
            }
        }

        // サブプロシージャ：重複レコード、または１セル２重入力をマーキングする
        private void colorViewForDuplicatePack(ref DataGridView dgv)
        {
            if (dgv.Rows.Count <= 0) return;

            DataTable dt = ((DataTable)dgv.DataSource).Copy();
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                string pack = dgv["pack_id", i].Value.ToString();
                DataRow[] dr = dt.Select("pack_id = '" + pack + "'");
                if (dr.Length >= 2)
                {
                    dgv["pack_id", i].Style.BackColor = Color.Red;
                    soundAlarm();
                }
                else
                {
                    dgv["pack_id", i].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                }
            }
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

        // データグリッドビューのダブルクリック時、データをエクセルへエクスポート
        private void dgvPack_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)dgvPack.DataSource;
            ExcelClass xl = new ExcelClass();
            xl.ExportToExcel(dt);
        }

        // ＬＯＴ集計グリッドビュー左のボタンを押下時、対象のモジュールをエクセルファイルへ出力する
        private void btnExportModule_Click(object sender, EventArgs e)
        {
            // 2016.08.29 FUJII 別スレッドで処理（処理スピード対策）
            var task = Task.Factory.StartNew(() =>
            {
                DataTable dt = new DataTable();
                string sql = "select a.*, b.pack_id, '" + txtCarton.Text + "' as carton_id from (" +
                             "select * from t_module where tray_id in (" +
                             "select tray_id from t_tray where pack_id in (" +
                             "select pack_id from t_pack where carton_id = '" + txtCarton.Text + "'))) a " +
                             "inner join t_tray b on a.tray_id = b.tray_id";
                TfSQL tf = new TfSQL();
                System.Diagnostics.Debug.Print(sql);
                tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql, ref dt);
                ExcelClass xl = new ExcelClass();
                // 2016.08.29 FUJII エクセルへの出力から、デスクトップＣＳＶへの出力へ変更
                xl.ExportToCsv(dt, System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\carton.csv");
            });
        }

        // 一時テーブルの選択された複数レコードを、一括消去させる
        private void btnDeleteSelection_Click(object sender, EventArgs e)
        {
            if (dtPack.Rows.Count <= 0) return;

            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
            if (dgvPack.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2)
            {
                MessageBox.Show("Please select range with only one column.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            foreach (DataGridViewCell cell in dgvPack.SelectedCells)
            {
                dtPack.Rows[cell.RowIndex].Delete();
            }
            dtPack.AcceptChanges();
            updateDataGridViews(dtPack, ref dgvPack, false);
            txtPack.Focus();
            txtPack.SelectAll();
        }

        //パックグリッドビュー上のボタンのクリックで、formTrayInPackを呼び出し、対応するトレーを表示する（デレゲートなし）
        private void dgvPack_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int currentRow = int.Parse(e.RowIndex.ToString());

            if (dgvPack.Columns[e.ColumnIndex] == openPack && currentRow >= 0)
            {
                //既にfrmTrayInPack が開かれている場合は、それを閉じるよう促す
                if (TfGeneral.checkOpenFormExists("frmTrayInPack"))
                {
                    MessageBox.Show("Please close the currently open form.", "Notice",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    return;
                }

                string packId = dgvPack["pack_id", currentRow].Value.ToString();
                string batch = dgvPack["batch", currentRow].Value.ToString();
                DateTime packDate = (DateTime)dgvPack["register_date", currentRow].Value;

                //デレゲートなし
                frmTrayInPack fT = new frmTrayInPack();
                fT.updateControls(packId, packDate, "", txtLoginName.Text, txtLoginDept.Text, userRole, batch, false, false, true, true, 3, false);
                fT.Show();
            }
        }

        // １パレットあたりのカートン数を変更する、テキストボックスのダブルクリックから起動
        private void txtOkCount_DoubleClick(object sender, EventArgs e)
        {
            //if (userRole != "super") return;

            // 既に同フォームが開かれている場合は、処理を行わない
            if (TfGeneral.checkOpenFormExists("frmCapacity"))
            {
                MessageBox.Show("Please close or complete another form.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            frmCapacity fC = new frmCapacity();
            //子イベントをキャッチして、データグリッドを更新する
            fC.RefreshEvent += delegate (object sndr, EventArgs excp)
            {
                capacity = fC.returnCapacity();
                updateDataGridViews(dtPack, ref dgvPack, false);
            };

            fC.updateControls(capacity);
            fC.Show();
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