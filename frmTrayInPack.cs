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
using System.Collections.Generic;

namespace TrayGuard
{
    public partial class frmTrayInPack : Form
    {
        //親フォームfrmTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        //データグリッドビュー用ボタン
        DataGridViewButtonColumn openTray;

        // プリント用テキストファイルの保存用フォルダを、基本設定ファイルで設定する
        string appconfig = System.Environment.CurrentDirectory + "\\info.ini";
        string productconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\tray_guard_desktop_for_pack.ini";

        //その他、非ローカル変数
        DataTable dtTray;
        DataTable dtLot;
        bool formAddMode;
        bool formReprintMode;
        bool packIdCanceled;
        bool packIdCartoned;
        string userRole;
        int okCount;
        int capacity = 13;
        string maxLot;
        DateTime registerDate;
        bool sound;
        int position;

        // 製品シリアル構成要素チェック用変数
        List<string> plant1;
        List<string> year1;
        List<string> week1;
        List<string> day1;
        List<string> line1;
        List<string> eeee1;
        List<string> revision1;
        List<string> mass1;
        List<string> flexure1;
        List<string> cover_base1;
        List<string> dframe1;
        List<string> fpc1;
        List<string> shift1;

        // コンストラクタ
        public frmTrayInPack()
        {
            InitializeComponent();

            // 製品シリアル構成要素の、変数への格納
            plant1 = readIni("MODULE NUMBERING CHECK FOR PACK", "PLANT", productconfig).Split(',').ToList();
            year1 = readIni("MODULE NUMBERING CHECK FOR PACK", "YEAR", productconfig).Split(',').ToList();
            week1 = readIni("MODULE NUMBERING CHECK FOR PACK", "WEEK", productconfig).Split(',').ToList();
            day1 = readIni("MODULE NUMBERING CHECK FOR PACK", "DAY", productconfig).Split(',').ToList();
            line1 = readIni("MODULE NUMBERING CHECK FOR PACK", "LINE", productconfig).Split(',').ToList();
            eeee1 = readIni("MODULE NUMBERING CHECK FOR PACK", "EEEE", productconfig).Split(',').ToList();
            revision1 = readIni("MODULE NUMBERING CHECK FOR PACK", "REVISION", productconfig).Split(',').ToList();
            mass1 = readIni("MODULE NUMBERING CHECK FOR PACK", "MASS", productconfig).Split(',').ToList();
            flexure1 = readIni("MODULE NUMBERING CHECK FOR PACK", "FLEXURE", productconfig).Split(',').ToList();
            cover_base1 = readIni("MODULE NUMBERING CHECK FOR PACK", "COVER/BASE", productconfig).Split(',').ToList();
            dframe1 = readIni("MODULE NUMBERING CHECK FOR PACK", "D-FRAME", productconfig).Split(',').ToList();
            fpc1 = readIni("MODULE NUMBERING CHECK FOR PACK", "FPC", productconfig).Split(',').ToList();
            shift1 = readIni("MODULE NUMBERING CHECK FOR PACK", "SHIFT", productconfig).Split(',').ToList();
        }

        // ロード時の処理
        private void frmTrayInPack_Load(object sender, EventArgs e)
        {
            this.Text = this.Text + " " + Assembly.GetExecutingAssembly().GetName().Version;
            // 当フォームの表示場所を指定
            this.Left = 250;
            this.Top = 20;
            if (position == 2) { this.Left = 400; this.Top = 30; }
            else if (position == 3) { this.Left = 450; this.Top = 40; }
            else if (position == 4) { this.Left = 500; this.Top = 50; }

            // 各種処理用のテーブルを生成し、データを読み込む
            dtTray = new DataTable();
            defineTrayTable(ref dtTray);
            if (!formAddMode) readTrayInfo(ref dtTray);

            // グリットビューの更新
            updateDataGridViews(dtTray, ref dgvTray, true);

            // 追加モード、閲覧モードの切り替え
            changeMode();

            // サブプロシージャ： バッチコンボボックスの設定
            setBatchComboBox();
        }

        // サブプロシージャ： バッチコンボボックスの設定
        private void setBatchComboBox()
        {
            string sql = "select content from t_criteria where criteria = 'BATCH' order by content"; 
            TfSQL tf = new TfSQL();
            tf.getComboBoxData(sql, ref cmbBatch);
        }

        // サブプロシージャ： 追加モード、閲覧モードの切り替え
        private void changeMode()
        {
            // 追加モードの場合
            if (formAddMode)
            {
                cmbBatch.Enabled = false; 
                txtTrayId.Enabled = true;
                btnRegisterPack.Text = "Register Pack";
                btnRegisterPack.Enabled = false;
                btnCancelPack.Enabled = false;
                btnClose.Enabled = true;
                btnDeleteSelection.Enabled = true;
            }
            // 閲覧モードの場合
            else
            {
                cmbBatch.Enabled = false;
                txtTrayId.Enabled = false;
                btnRegisterPack.Text = "Re-Print Label";
                // 2016.09.24 FUJII  再プリントモードを追加（frmPackにて、トレーＩＤをキーとして、パックＩＤを特定した場合のみ、再印刷可能）
                //btnRegisterPack.Enabled = packIdCanceled ? false : ((formReprintMode && userRole == "super") ? true : false);
                //btnRegisterPack.Enabled = packIdCanceled ? false : (userRole == "super" ? true : false);
                // 2017.01.18 FUJII  ユーザー制限解除
                btnRegisterPack.Enabled = packIdCanceled ? false : true;

                btnCancelPack.Enabled = false;
                btnClose.Enabled = true;
                btnDeleteSelection.Enabled = false;
                if (userRole == "super") btnCancelPack.Enabled = (packIdCanceled || packIdCartoned) ? false : true;
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
        public void updateControls(string tid, DateTime tdate, string uid, string uname, string udept, string urole, string shift, 
            bool addMode, bool returnMode, bool canceled, bool cartoned, int pos, bool reprintMode)
        {
            txtPackId.Text = tid;
            dtpRegisterDate.Value = tdate;
            txtLoginName.Text = uname;
            txtLoginDept.Text = udept;
            userRole = urole;
            if (!addMode) cmbBatch.Text = shift;
            formAddMode = addMode;
            packIdCanceled = canceled;
            packIdCartoned = cartoned;
            position = pos;
            formReprintMode = reprintMode;
        }

        // サブプロシージャ：データテーブルの定義
        private void defineTrayTable(ref DataTable dt)
        {
            dt.Columns.Add("tray_id", typeof(string));
            dt.Columns.Add("lot", typeof(string));
            dt.Columns.Add("qty", typeof(int));
            dt.Columns.Add("register_date", typeof(DateTime));
            dt.Columns.Add("rg_dept", typeof(string));
            dt.Columns.Add("multi_lot", typeof(string));
            dt.Columns.Add("check", typeof(string));
        }

        // サブプロシージャ：ＤＢからデータテーブルへの読み込み
        private void readTrayInfo(ref DataTable dt)
        {
            dt.Rows.Clear();
            string sql = "select tray_id, lot, qty, register_date, rg_dept, multi_lot, " +
                "'OK' as check from t_tray where pack_id='" + txtPackId.Text + "'";
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
            if(formAddMode) colorViewForCanceledAndPacked(ref dgv1);

            // 重複レコード、および１セル２重入力をマーキングする
            if (formAddMode) colorViewForDuplicateTray(ref dgv1);

            // グリットビュー右端にボタンを追加（初回のみ）
            if (load && !formAddMode) addButtonsToDataGridView(dgv1);

            //行ヘッダーに行番号を表示する
            for (int i = 0; i < dgv1.Rows.Count; i++) dgv1.Rows[i].HeaderCell.Value = (i + 1).ToString();

            //行ヘッダーの幅を自動調節する
            dgv1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            // 一番下の行を表示する（インライン）
            if (dgv1.Rows.Count >= 1) dgv1.FirstDisplayedScrollingRowIndex = dgv1.Rows.Count - 1;

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
                    cmbBatch.Enabled = false;
                    txtTrayId.Enabled = false;
                    btnRegisterPack.Enabled = true;
                    btnDeleteSelection.Enabled = true;
                }
                else
                {
                    cmbBatch.Enabled = false;
                    txtTrayId.Enabled = true;
                    btnRegisterPack.Enabled = false;
                    btnDeleteSelection.Enabled = true;
                    txtTrayId.SelectAll(); // 連続スキャン用
                }
            }
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

        // サブプロシージャ：シリアル番号重複なしのＰＡＳＳ個数を取得する
        private int getOkCount(DataTable dt)
        {
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "tray_id", "check" });
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

            // 一時テーブルに、ロット集計を格納する
            string sql1 = "select lot, count(lot) as qty from t_module where tray_id in (";
            string sql2 = string.Empty;
            var query1 = dt.AsEnumerable()
                        .Select(row => new { tray_id = row.Field<string>("tray_id") });
            foreach (var q in query1) sql2 += "'" + q.tray_id + "', ";
            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ") group by lot order by qty desc, lot";
            System.Diagnostics.Debug.Print(sql3);
            DataTable dtTemp = new DataTable();
            TfSQL tf = new TfSQL();
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql3, ref dtTemp);

            // ロット集計表示テーブルに、一時テーブルの情報を移す
            dgvLotSummary.DataSource = null;
            dgvLotSummary.Refresh();
            dtLot = new DataTable();
            var query2 = dtTemp.AsEnumerable().Select(r => new {lot = r.Field<string>("lot"), qty = r.Field<Int64>("qty") });
            // 列の追加
            foreach (var q in query2) dtLot.Columns.Add(q.lot, typeof(int));
            dtLot.Columns.Add("total", typeof(int));
            // 行の追加
            dtLot.Rows.Add(); 
            foreach (var q in query2) dtLot.Rows[0][q.lot] = q.qty;
            dtLot.Rows[0]["total"] = query2.Sum(a => a.qty);
 
            dgvLotSummary.DataSource = dtLot;
            return query2.First().lot;
        }

        // シリアルがスキャンされた時の処理            
        private void txtModuleId_KeyDown(object sender, KeyEventArgs e)
        {
            // エンターキーの場合、テキストボックスの桁数が１５桁の場合のみ、処理を行う
            if (e.KeyCode != Keys.Enter || txtTrayId.Text.Length != 21) return;

            string tray = txtTrayId.Text;
            string sql = "select tray_id, lot, qty, register_date, rg_dept, multi_lot, " +
                "case when cancel_date is not null then to_char(cancel_date,'YYYY/MM/DD') " +
                "when pack_id is not null then pack_id else 'OK' end as check " +
                "from t_tray where tray_id='" + tray + "'";
            System.Diagnostics.Debug.Print(sql);
            DataTable dt = new DataTable();
            TfSQL tf = new TfSQL();
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql, ref dt);

            if (dt.Rows.Count != 0)
            {
                bool warnedAlready = false;
                string sql2 = " select sum(case when tm1.test_result in ('n/a', 'PASS') then 1 else 0 end) as okcount" +
                          "      , count(tm1.test_result) as totalcount" +
                          " from t_tray " +
                          " left join t_module tm1 on tm1.tray_id = t_tray.tray_id " +
                          " where t_tray.tray_id = '" + tray + "'";
                DataTable dt2 = new DataTable();
                TfSQL tf2 = new TfSQL();
                tf2.sqlDataAdapterFillDatatableFromTrayGuardDb(sql2, ref dt2);
            
                if ((long)dt2.Rows[0]["okcount"] != 24)
                {
                    MessageBox.Show("OK module is not 24", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    warnedAlready = true;
                    //if (txtLoginDept.Text != "PC") return;
                }
                if ((long)dt2.Rows[0]["totalcount"] != 24 && !warnedAlready)
                {
                    MessageBox.Show("Module data is not 24", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    warnedAlready = true;
                    //if (txtLoginDept.Text != "PC") return;
                }

                if ((int)dt.Rows[0]["qty"] != 24 && !warnedAlready)
                {
                    MessageBox.Show("Module in tray is not 24.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    warnedAlready = true;
                    //if (txtLoginDept.Text != "PC") return;
                }
            }

            // 先ずは、シリアルの構成要素のパターンが適正か、ユーザーデスクトップの設定ファイルを使用して確認する
            // 2017/03/07 Fujii 構成要素チェックのオフ
            //string matchResult = matchSerialNumberingPattern(tray);
            string matchResult = string.Empty;
            if (matchResult != string.Empty)
            {
                MessageBox.Show(matchResult + " does not match with desktop file's setting.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // テスターデータに該当がない場合でも、ユーザー認識用に表示するための処理
            DataRow dr = dtTray.NewRow();
            dr["tray_id"] = tray;
            // テスターデータに該当がある場合の処理
            if (dt.Rows.Count != 0)
            {
                dr["lot"] = (string)dt.Rows[0]["lot"];
                dr["qty"] = (int)dt.Rows[0]["qty"];
                dr["register_date"] = (DateTime)dt.Rows[0]["register_date"];
                dr["rg_dept"] = (string)dt.Rows[0]["rg_dept"];
                dr["multi_lot"] = (string)dt.Rows[0]["multi_lot"];
                dr["check"] = (string)dt.Rows[0]["check"];
            }

            // メモリ上のテーブルにレコードを追加
            dtTray.Rows.Add(dr);

            // データグリットビューの更新
            updateDataGridViews(dtTray, ref dgvTray, false);
        }

        // サブプロシージャ：シリアルの構成要素のパターンが適正か、ユーザーデスクトップの設定ファイルを使用して確認する
        private string matchSerialNumberingPattern(string tray)
        {
            string result = string.Empty;

            string sql = "select module_id, substr(module_id, 1, 3) as plant, substr(module_id, 4, 1) as year, substr(module_id, 5, 2) as week, " +
                "substr(module_id, 7, 1) as day, substr(module_id, 8, 1) as line, substr(module_id, 12, 4) as eeee, " +
                "substr(module_id, 16, 1) as revision, substr(module_id, 19, 1) as mass, substr(module_id, 20, 1) as flexure, " +
                "substr(module_id, 21, 1) as cover_base, substr(module_id, 22, 2) as dframe, substr(module_id, 23, 1) as fpc, " +
                "substr(module_id, 24, 1) as shift from t_module where tray_id = '" + tray + "'";
            System.Diagnostics.Debug.Print(sql);
            DataTable dt = new DataTable();
            TfSQL tf = new TfSQL();
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql, ref dt);

            // ＤＢから取得したモジュール構成要素のリストを、デスクトップで指定されたリストでフィルターを掛け、該当文字列を警告メッセージに追加する
            if (!plant1.Any(s => s == "XXX"))
            {
                List<string> plant2 = dt.AsEnumerable().Select(r => r.Field<string>("plant")).Except(plant1).ToList();
                foreach(var e in plant2) result += "Plant '" + e + "'" + Environment.NewLine;
            }
            if (!year1.Any(s => s == "XXX"))
            {
                List<string> year2 = dt.AsEnumerable().Select(r => r.Field<string>("year")).Except(year1).ToList();
                foreach (var e in year2) result += "Year '" + e + "'" + Environment.NewLine;
            }
            if (!week1.Any(s => s == "XXX"))
            {
                List<string> week2 = dt.AsEnumerable().Select(r => r.Field<string>("week")).Except(week1).ToList();
                foreach (var e in week2) result += "Week '" + e + "'" + Environment.NewLine;
            }
            if (!day1.Any(s => s == "XXX"))
            {
                List<string> day2 = dt.AsEnumerable().Select(r => r.Field<string>("day")).Except(day1).ToList();
                foreach (var e in day2) result += "Day '" + e + "'" + Environment.NewLine;
            }
            if (!line1.Any(s => s == "XXX"))
            {
                List<string> line2 = dt.AsEnumerable().Select(r => r.Field<string>("line")).Except(line1).ToList();
                foreach (var e in line2) result += "Line '" + e + "'" + Environment.NewLine;
            }
            if (!eeee1.Any(s => s == "XXX"))
            {
                List<string> eeee2 = dt.AsEnumerable().Select(r => r.Field<string>("eeee")).Except(eeee1).ToList();
                foreach (var e in eeee2) result += "4E '" + e + "'" + Environment.NewLine;
            }
            if (!revision1.Any(s => s == "XXX"))
            {
                List<string> revision2 = dt.AsEnumerable().Select(r => r.Field<string>("revision")).Except(revision1).ToList();
                foreach (var e in revision2) result += "Revision '" + e + "'" + Environment.NewLine;
            }
            if (!mass1.Any(s => s == "XXX"))
            {
                List<string> mass2 = dt.AsEnumerable().Select(r => r.Field<string>("mass")).Except(mass1).ToList();
                foreach (var e in mass2) result += "Mass '" + e + "'" + Environment.NewLine;
            }
            if (!flexure1.Any(s => s == "XXX"))
            {
                List<string> flexure2 = dt.AsEnumerable().Select(r => r.Field<string>("flexure")).Except(flexure1).ToList();
                foreach (var e in flexure2) result += "Flexure '" + e + "'" + Environment.NewLine;
            }
            if (!cover_base1.Any(s => s == "XXX"))
            {
                List<string> cover_base2 = dt.AsEnumerable().Select(r => r.Field<string>("cover_base")).Except(cover_base1).ToList();
                foreach (var e in cover_base2) result += "Cover/base '" + e + "'" + Environment.NewLine;
            }
            if (!dframe1.Any(s => s == "XXX"))
            {
                List<string> dframe2 = dt.AsEnumerable().Select(r => r.Field<string>("dframe")).Except(dframe1).ToList();
                foreach (var e in dframe2) result += "Dframe '" + e + "'" + Environment.NewLine;
            }
            if (!fpc1.Any(s => s == "XXX"))
            {
                List<string> fpc2 = dt.AsEnumerable().Select(r => r.Field<string>("fpc")).Except(fpc1).ToList();
                foreach (var e in fpc2) result += "Fpc '" + e + "'" + Environment.NewLine;
            }
            if (!shift1.Any(s => s == "XXX"))
            {
                List<string> shift2 = dt.AsEnumerable().Select(r => r.Field<string>("shift")).Except(shift1).ToList();
                foreach (var e in shift2) result += "Shift '" + e + "'" + Environment.NewLine;
            }

            return result;
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    if (!plant1.Any(s=> s=="XXX") && dt.Rows[i]["plant"].ToString() != plant) result += module + " Plant '" + dt.Rows[i]["plant"].ToString() + "'" + Environment.NewLine;
            //    else if (!year1.Any(s=> s=="XXX") && dt.Rows[i]["year"].ToString() != year) result += module + " Year '" + dt.Rows[i]["year"].ToString() + "'" + Environment.NewLine;
            //    else if (!week1.Any(s=> s=="XXX") && dt.Rows[i]["week"].ToString() != week) result += module + " Week '" + dt.Rows[i]["week"].ToString() + "'" + Environment.NewLine;
            //    else if (!day1.Any(s=> s=="XXX") && dt.Rows[i]["day"].ToString() != day) result += module + " Day '" + dt.Rows[i]["day"].ToString() + "'" + Environment.NewLine;
            //    else if (!line1.Any(s=> s=="XXX") && dt.Rows[i]["line"].ToString() != line) result += module + " Line '" + dt.Rows[i]["line"].ToString() + "'" + Environment.NewLine;
            //    else if (!eeee1.Any(s=> s=="XXX") && dt.Rows[i]["eeee"].ToString() != eeee) result += module + " 4E '" + dt.Rows[i]["eeee"].ToString() + "'" + Environment.NewLine;
            //    else if (!revision1.Any(s=> s=="XXX") && dt.Rows[i]["revision"].ToString() != revision) result += module + " Revision '" + dt.Rows[i]["revision"].ToString() + "'" + Environment.NewLine;
            //    else if (!mass1.Any(s=> s=="XXX") && dt.Rows[i]["mass"].ToString() != mass) result += module + " Mass '" + dt.Rows[i]["mass"].ToString() + "'" + Environment.NewLine;
            //    else if (!flexure1.Any(s=> s=="XXX") && dt.Rows[i]["flexure"].ToString() != flexure) result += module + " Flexure '" + dt.Rows[i]["flexure"].ToString() + "'" + Environment.NewLine;
            //    else if (!cover_base1.Any(s=> s=="XXX") && dt.Rows[i]["cover_base"].ToString() != cover_base) result += module + " Cover/base '" + dt.Rows[i]["cover_base"].ToString() + "'" + Environment.NewLine;
            //    else if (!dframe1.Any(s=> s=="XXX") && dt.Rows[i]["dframe"].ToString() != dframe) result += module + " D-Frame '" + dt.Rows[i]["dframe"].ToString() + "'" + Environment.NewLine;
            //    else if (!fpc1.Any(s=> s=="XXX") && dt.Rows[i]["fpc"].ToString() != fpc) result += module + " FPC '" + dt.Rows[i]["fpc"].ToString() + "'" + Environment.NewLine;
            //    else if (!shift1.Any(s=> s=="XXX") && dt.Rows[i]["shift"].ToString() != shift) result += module + " Shift '" + dt.Rows[i]["shift"].ToString() + "'" + Environment.NewLine;
            //}
        }

        // 登録ボタン押下時、各種確認、ボックスＩＤの発行、シリアルの登録、バーコードラベルのプリントを行う
        private void btnRegisterTray_Click(object sender, EventArgs e)
        {
            if (getOkCount(dtTray) != dtTray.Rows.Count)
            {
                MessageBox.Show("Tray is not 13.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (txtLoginDept.Text != "MFG")
            {
                //閲覧モードの場合は、プリントアウト
                if (!formAddMode)
                {
                    string pack = txtPackId.Text;
                    TfSato tfs = new TfSato();

                    // ビンＡか、ビンＢか、ユーザーに選択させる
                    //DialogResult binResult = MessageBox.Show("Please click YES for Bin A, NO for Bin B.", "Print Option",
                    //    MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    //string userBin = (binResult == DialogResult.Yes) ? "A" : "B";
                    string userBin = cmbBatch.Text;

                    // ペガトロン用か否か、ユーザーに選択させる
                    DialogResult result = MessageBox.Show("Do you print Pegatoron label also?", "Print Option", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    if (result == DialogResult.Yes)
                    {
                        //tfs.printStart("packCartonInternal", pack, dtLot, cmbBatch.Text, dtpRegisterDate.Value, "Pack", "Fox", 1, userBin);
                        tfs.printStart("packCartonPega", pack, dtLot, cmbBatch.Text, dtpRegisterDate.Value, "Pack", "Pega", 1, userBin);
                    }
                    else
                    {
                        tfs.printStart("packCartonInternal", pack, dtLot, cmbBatch.Text, dtpRegisterDate.Value, "Pack", "Fox", 1, userBin);
                    }
                    return;
                }
            }

            //以下追加モードの場合の処理

            // 2016.08.22 FUJII  パックＩＤの新規採番プロシージャ「getNewPackId」を、トランザクション処理バージョンへ変更
            //if (cmbBatch.Text == string.Empty)
            //{
            //    MessageBox.Show("Please select Batch.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            //    return;
            //}

            // 登録処理中は、閉じるボタンを除き、すべてのコントロールを無効にする
            cmbBatch.Enabled = false;
            txtTrayId.Enabled = false;
            btnRegisterPack.Enabled = false;
            btnDeleteSelection.Enabled = false;

            //パックＩＤの新規採番
            string packNew = txtPackId.Text; // パック内容変更の処理も同ボタンで行うため、テキストボックスの既存ＩＤを保持する
            if (formAddMode)
            {
                TfSQL tf0 = new TfSQL();
                packNew = tf0.sqlGetNewPackId(maxLot, cmbBatch.Text, txtLoginName.Text, dtLot, ref registerDate);
                if (packNew == string.Empty)
                {
                    MessageBox.Show("An error happened in the pack id issuing process.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                txtPackId.Text = packNew;
            }

            if (txtLoginDept.Text == "MFG")
            {
                // ビンＡか、ビンＢか、ユーザーに選択させる
                string pack = txtPackId.Text;
                TfSato tfs = new TfSato();
                //DialogResult binResult = MessageBox.Show("Please click YES for Bin A, NO for Bin B.", "Print Option",
                //    MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                //string userBin = (binResult == DialogResult.Yes) ? "A" : "B";
                string userBin = cmbBatch.Text;
                tfs.printStart("packCartonInternal", pack, dtLot, cmbBatch.Text, dtpRegisterDate.Value, "Pack", "Non", 1, userBin);
                if (!formAddMode)
                {
                    return;
                }
            }

            //トレーテーブルのフィールドＰＡＣＫＩＤ、その他を更新する
            TfSQL tf1 = new TfSQL();
            bool res = tf1.sqlMultipleUpdateTrayInPack(dtTray, packNew);

            if (res)
            {
                //登録済みの状態を表示
                txtPackId.Text = packNew;
                dtpRegisterDate.Value = registerDate;

                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                this.Focus();
                MessageBox.Show("Pack ID: " + packNew + Environment.NewLine +
                    "and its trays were registered.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //ユーザーによるメッセージボックス確認後の処理
                txtPackId.Text = String.Empty;
                txtTrayId.Text = String.Empty;
                dtTray.Clear();
                capacity = 13;
                updateDataGridViews(dtTray, ref dgvTray, false);
            }
        }

        // スーパーユーザーに限り、登録済パックをキャンセルできる（カートン処理後は不可）
        private void btnCancelPack_Click(object sender, EventArgs e)
        {
            // 本当に削除してよいか、２重で確認する。
            DialogResult result1 = MessageBox.Show("Do you really want to cancel this tray?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result1 == DialogResult.No) return;

            DialogResult result2 = MessageBox.Show("Are you really sure? Please select NO if you are not sure.", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) return;

            // キャンセルの実行
            string packId = txtPackId.Text;
            TfSQL tf = new TfSQL();
            bool res = tf.sqlCancelTrayInPack(dtTray,packId, txtLoginName.Text);
            if (res)
            {
                //本フォームのデータグリットビュー更新
                dtTray.Clear();
                updateDataGridViews(dtTray, ref dgvTray, false);

                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                this.Focus();
                MessageBox.Show("Pack ID " + packId + " and its trays were canceled.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // メッセージボックスの確認後、閉じる
                Close();
            }
            else
            {
                MessageBox.Show("Cancel process was not successful.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // １パックあたりのトレー数を変更する、テキストボックスのダブルクリックから起動
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
                updateDataGridViews(dtTray, ref dgvTray, false);
            };

            fC.updateControls(capacity);
            fC.Show();
        }

        // 当フォームのクローズ時、他の子フォームが開いていないことを確認する
        private void btnCancel_Click(object sender, EventArgs e)
        {
            // frmCapacity を閉じていない場合は、先に閉じるよう通知する
            if (TfGeneral.checkOpenFormExists("frmCapacity"))
            {
                MessageBox.Show("You need to close Capacity form before canceling.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            Close();
        }

        // サブプロシージャ：キャンセル済みまたはパック済みのトレーをマーキングする
        private void colorViewForCanceledAndPacked(ref DataGridView dgv)
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
        private void colorViewForDuplicateTray(ref DataGridView dgv)
        {
            if (dgv.Rows.Count <= 0) return;

            DataTable dt = ((DataTable)dgv.DataSource).Copy();
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                string module = dgv["tray_id", i].Value.ToString();
                DataRow[] dr = dt.Select("tray_id = '" + module + "'");
                if (dr.Length >= 2)
                {
                    dgv["tray_id", i].Style.BackColor = Color.Red;
                    soundAlarm();
                }
                else
                {
                    dgv["tray_id", i].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
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
        private void dgvTray_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)dgvTray.DataSource;
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
                string sql = "select *, '" + txtPackId.Text + "' as pack_id from t_module where tray_id in (" +
                             "select tray_id from t_tray where pack_id = '" + txtPackId.Text + "')";
                TfSQL tf = new TfSQL();
                System.Diagnostics.Debug.Print(sql);
                tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql, ref dt);
                ExcelClass xl = new ExcelClass();
                // 2016.08.29 FUJII エクセルへの出力から、デスクトップＣＳＶへの出力へ変更
                xl.ExportToCsv(dt, System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\pack.csv");
            });
        }

        // 一時テーブルの選択された複数レコードを、一括消去させる
        private void btnDeleteSelection_Click(object sender, EventArgs e)
        {
            if (dtTray.Rows.Count <= 0) return;

            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
            if (dgvTray.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2)
            {
                MessageBox.Show("Please select range with only one column.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            DialogResult result = MessageBox.Show("Do you really want to delete the selected rows?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.No) return;

            foreach (DataGridViewCell cell in dgvTray.SelectedCells)
            {
                dtTray.Rows[cell.RowIndex].Delete();
            }
            dtTray.AcceptChanges();
            updateDataGridViews(dtTray, ref dgvTray, false);
            txtTrayId.Focus();
            txtTrayId.SelectAll();
        }

        //トレーグリッドビュー上のボタンのクリックで、formModuleInTrayを呼び出し、対応するモジュールを表示する（デレゲートなし）
        private void dgvTray_CellContentClick(object sender, DataGridViewCellEventArgs e)
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
                string line = VBS.Right(trayId, 1);
                DateTime trayDate = (DateTime)dgvTray["register_date", currentRow].Value;

                //デレゲートなし。当トレーフォームのフォーム位置が２の場合は、モジュールフォームの位置は３。それ以外は２。
                frmModuleInTray fM = new frmModuleInTray();
                fM.updateControls(trayId, trayDate, "", txtLoginName.Text, txtLoginDept.Text, userRole, line, false, false, true, true, 4, false, false);
                fM.Show();
            }
        }

        // 項目マスタ登録フォームの起動
        private void cmbBatch_KeyDown(object sender, KeyEventArgs e)
        {
            // スーパーユーザーが使用中の場合のみ、ユーザーマスター変更フォームを開く
            if (e.KeyCode != Keys.Enter || userRole != "super") return;
            if (TfGeneral.checkOpenFormExists("frmMasterCriteria")) return;
            frmMasterCriteria fI = new frmMasterCriteria("BATCH");
            //子イベントをキャッチして、データグリッドを更新する
            fI.RefreshEvent += delegate (object sndr, EventArgs excp)
            {
                //バッチコンボボックスの設定
                setBatchComboBox();
            };
            fI.Show();
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
    }
}