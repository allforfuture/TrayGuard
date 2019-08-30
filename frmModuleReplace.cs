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
    public partial class frmModuleReplace : Form
    {
        //親フォームfrmTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        // プリント用テキストファイルの保存用フォルダを、基本設定ファイルで設定する
        string appconfig = System.Environment.CurrentDirectory + "\\info.ini";
        string productconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\tray_guard_desktop.ini";
        string outPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\NTRS Log\";

        //その他非ローカル変数
        DataTable dtModule;
        string trayId;
        bool formReturnMode;
        bool sound;
        bool testerNg;
        string totalSwitch;
        string OK2ShipCheckSwitch;
        string forcedNGSwitch;

        // 製品シリアル構成要素チェック用変数
        string plant;
        string year;
        string week;
        string day;
        string line;
        string eeee;
        string revision;
        string mass;
        string flexure;
        string cover_base;
        string dframe;
        string fpc;
        string shift;

        // コンストラクタ
        public frmModuleReplace()
        {
            InitializeComponent();

            // 製品シリアル構成要素の、変数への格納
            plant = readIni("MODULE NUMBERING CHECK", "PLANT", productconfig);
            year = readIni("MODULE NUMBERING CHECK", "YEAR", productconfig);
            week = readIni("MODULE NUMBERING CHECK", "WEEK", productconfig);
            day = readIni("MODULE NUMBERING CHECK", "DAY", productconfig);
            line = readIni("MODULE NUMBERING CHECK", "LINE", productconfig);
            eeee = readIni("MODULE NUMBERING CHECK", "EEEE", productconfig);
            revision = readIni("MODULE NUMBERING CHECK", "REVISION", productconfig);
            mass = readIni("MODULE NUMBERING CHECK", "MASS", productconfig);
            flexure = readIni("MODULE NUMBERING CHECK", "FLEXURE", productconfig);
            cover_base = readIni("MODULE NUMBERING CHECK", "COVER/BASE", productconfig);
            dframe = readIni("MODULE NUMBERING CHECK", "D-FRAME", productconfig);
            fpc = readIni("MODULE NUMBERING CHECK", "FPC", productconfig);
            shift = readIni("MODULE NUMBERING CHECK", "SHIFT", productconfig);
        }

        // ロード時の処理
        private void Form4_Load(object sender, EventArgs e)
        {
            // 製品・データ、マッチング機能用設定内容の取得
            forcedNGSwitch = readIni("MODULE-DATA MATCHING", "FORCED NG SWITCH", appconfig);
            OK2ShipCheckSwitch = readIni("MODULE-DATA MATCHING", "OK2SHIP CHECK SWITCH", appconfig);
            totalSwitch = (OK2ShipCheckSwitch == "OFF" && forcedNGSwitch == "OFF") ? "OFF" : "ON";

            // ログ用フォルダの作成
            if (!Directory.Exists(outPath)) Directory.CreateDirectory(outPath);

            //フォームの場所を指定
            this.Left = 450;
            this.Top = 100;
            dtModule = new DataTable();
            defineDatatable(ref dtModule);
            updateDataGridViews(dtModule, ref dgvModule);

            // 現在の一時登録件数を変数へ保持する
            int okCount = getOkCount(dtModule);
            txtRow.Text = okCount.ToString();
        }

        // サブプロシージャ： テスター用ＳＥＬＥＣＴケース句の作成
        private string makeSqlCaseClause(string criteria)
        {
            string sql = "case ";
            foreach (string c in criteria.Split(','))
            { sql += "when c.process_cd like " + VBS.Left(c, c.Length - 1) + "%' then " + c + " "; };
            sql += "else c.process_cd end as tester_id ";
            System.Diagnostics.Debug.Print(sql);
            return sql;
        }

        // サブプロシージャ： テスター用ＷＨＥＲＥ句の作成
        private string makeSqlWhereClause(string criteria)
        {
            string sql = "where ";
            foreach (string c in criteria.Split(','))
            { sql += "c.process_cd like " + VBS.Left(c, c.Length - 1) + "%' or "; };
            sql = VBS.Left(sql, sql.Length - 3);
            System.Diagnostics.Debug.Print(sql);
            return sql;
        }

        // サブプロシージャ: ＤＴの定義
        private void defineDatatable(ref DataTable dt)
        {
            dt.Columns.Add("module_id", typeof(string));
            dt.Columns.Add("lot", typeof(string));
            dt.Columns.Add("bin", typeof(string));
            dt.Columns.Add("tester_id", typeof(string));
            dt.Columns.Add("test_result", typeof(string));
            dt.Columns.Add("test_date", typeof(DateTime));
            dt.Columns.Add("r_mode", typeof(string));
        }

        // サブプロシージャ：親フォームで呼び出し、親フォームの情報を、テキストボックスへ格納して引き継ぐ
        public void updateControls(string tray, string serno, int row, bool mode, string binselect)
        {
            trayId = tray;
            txtBefore.Text = serno;
            txtRow.Text = row.ToString();
            formReturnMode = mode;
            txtBinSelect.Text = binselect;
        }

        // サブプロシージャ：データグリットビューの更新
        private void updateDataGridViews(DataTable dt, ref DataGridView dgv)
        {
            // データグリットビューへＤＴＡＡＴＡＢＬＥを格納
            dgv.DataSource = dt;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            // テスト結果がＦＡＩＬまたはレコードなしのシリアルをマーキングする
            if (dt.Rows.Count <= 0) return;
 
            if (dgv["test_result", 0].Value.ToString() == "FAIL" || dgv["test_result", 0].Value.ToString() == "NG" || dgv["test_result", 0].Value.ToString() == string.Empty)
            {
                dgv["tester_id", 0].Style.BackColor = Color.Red;
                dgv["test_result", 0].Style.BackColor = Color.Red;
                dgv["test_date", 0].Style.BackColor = Color.Red;
                soundAlarm();
                testerNg = true;
            }
            else
            {
                dgv["tester_id", 0].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                dgv["test_result", 0].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                dgv["test_date", 0].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                testerNg = false;
            }
        }
        // サブプロシージャ：シリアル番号重複なしのＰＡＳＳ個数を取得する
        private int getOkCount(DataTable dt)
        {
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "module_id", "test_result" });
            DataRow[] dr;
            if (totalSwitch == "ON") dr = distinct.Select("test_result in ('OK','n/a')"); //distinct.Select(); ＦＡＩＬも流す、初期段階の設定
            else dr = distinct.Select();  // マッチングスイッチオフの場合は、テスト結果に関係なく行数をカウントする

            return dr.Length;
        }

        // テスト結果を格納するクラス
        public class TestResult
        {
            public string tester_id { get; set; }
            public string test_result { get; set; }
            public string test_date { get; set; }
        }
        // テスト結果のプロセスコードのみを格納するクラス
        public class ProcessList
        {
            public string tester_id { get; set; }
        }

        // 変更後モジュールがスキャンされたときの処理
        private void txtAfter_KeyDown(object sender, KeyEventArgs e)
        {
            // エンターキーの場合、テキストボックスの桁数が１７桁または２４桁の場合のみ、処理を行う
            if (e.KeyCode != Keys.Enter) return;
            if (txtAfter.Text.Length != 17 && txtAfter.Text.Length != 24) return;

            // データテーブルのクリア
            dtModule.Rows.Clear();

            // ＢＡＳＥシリアルからＨＯＯＰシリアルを取得し（ステップ１）、両方のシリアルに該当する、テスト結果・プロセス名・テスト日時を取得する（ステップ２、ステップ３）
            TfSQL tf = new TfSQL();
            DataTable dt = new DataTable();
            string log = string.Empty;
            string module = txtAfter.Text;
            string mdlShort = VBS.Left(module, 17);
            string mdlNtrs = string.Empty;
            string mdlOK2ShipResult = string.Empty;
            // 2016.08.18 FUJIKI FORCED NG CHECK を追加
            string mdlForcedNGResult = string.Empty;
            // 2016.08.18 FUJIKI FORCED NG CHECK を追加
            string scanTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string displayAll = string.Empty;   // ログ用
            string textResult = "PASS";
            string mdlSerialBin = string.Empty;
            string textSelectBin = string.Empty;

            

            // メモリ上のテーブルにレコードを追加
            DataRow dr = dtModule.NewRow();
            dr["module_id"] = module;
            dr["lot"] = VBS.Left(module, 8);
            dr["bin"] = mdlSerialBin;
            dr["tester_id"] = displayAll;
            dr["test_result"] = textResult;
            dr["test_date"] = DateTime.ParseExact(scanTime, "yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture); ;
            dr["r_mode"] = formReturnMode ? "T" : "F";
            dtModule.Rows.Add(dr);

            // アプリケーションフォルダに、日付とテスト結果のログを付ける
            log = Environment.NewLine + scanTime + "," + module + "," + displayAll + ":" + textResult;

            // 同日日付のファイルが存在する場合は追記し、存在しない場合はファイルを作成追記する（AppendAllText がやってくれる）
            try
            {
                string outFile = outPath + DateTime.Today.ToString("yyyyMMdd") + ".txt";
                System.IO.File.AppendAllText(outFile, log, System.Text.Encoding.GetEncoding("UTF-8"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // 現在の一時登録件数を変数へ保持する
            int okCount = getOkCount(dtModule);
            txtRow.Text = okCount.ToString();

            // データグリットビューの更新
            updateDataGridViews(dtModule, ref dgvModule);

        }

        // サブプロシージャ：シリアルの構成要素のパターンが適正か、ユーザーデスクトップの設定ファイルを使用して確認する
        private string matchSerialNumberingPattern(string serial)
        {
            string result = string.Empty;

            // 設定ファイル ＬＩＮＥ ＝ ＸＸＸ の場合は、検証しない
            // if (line == "XXX") return string.Empty;
            if      (!plant.Equals("XXX")      && VBS.Mid(serial,  1, 3) != plant)      return "Plant '"      + VBS.Mid(serial,  1, 3) + "'";
            else if (!year.Equals("XXX")       && VBS.Mid(serial,  4, 1) != year)       return "Year '"       + VBS.Mid(serial,  4, 1) + "'";
            else if (!week.Equals("XXX")       && VBS.Mid(serial,  5, 2) != week)       return "Week '"       + VBS.Mid(serial,  5, 2) + "'";
            else if (!day.Equals("XXX")        && VBS.Mid(serial,  7, 1) != day)        return "Day '"        + VBS.Mid(serial,  7, 1) + "'";
            else if (!line.Equals("XXX")       && VBS.Mid(serial,  8, 1) != line)       return "Line '"       + VBS.Mid(serial,  8, 1) + "'";
            else if (!eeee.Equals("XXX")       && VBS.Mid(serial, 12, 4) != eeee)       return "4E '"         + VBS.Mid(serial, 12, 4) + "'";
            else if (!revision.Equals("XXX")   && VBS.Mid(serial, 16, 1) != revision)   return "Revision '"   + VBS.Mid(serial, 16, 1) + "'";
            else if (!mass.Equals("XXX")       && VBS.Mid(serial, 19, 1) != mass)       return "Mass '"       + VBS.Mid(serial, 19, 1) + "'";
            else if (!flexure.Equals("XXX")    && VBS.Mid(serial, 20, 1) != flexure)    return "Flexure '"    + VBS.Mid(serial, 20, 1) + "'";
            else if (!cover_base.Equals("XXX") && VBS.Mid(serial, 21, 1) != cover_base) return "Cover/base '" + VBS.Mid(serial, 21, 1) + "'";
            else if (!dframe.Equals("XXX")     && VBS.Mid(serial, 22, 1) != dframe)     return "D-Frame '"    + VBS.Mid(serial, 22, 1) + "'";
            else if (!fpc.Equals("XXX")        && VBS.Mid(serial, 23, 1) != fpc)        return "FPC '"        + VBS.Mid(serial, 23, 1) + "'";
            else if (!shift.Equals("XXX")      && VBS.Mid(serial, 24, 1) != shift)      return "Shift '"      + VBS.Mid(serial, 24, 1) + "'";
            else return string.Empty;
        }

        // 登録済みのシリアルおよびその付帯情報を、ＵＰＤＡＴＥ文で置き換える
        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (dtModule.Rows.Count <= 0) return; //if (testerNg || dtModule.Rows.Count <= 0) 
            if (!txtRow.Text.Equals("1"))
            {
                MessageBox.Show("Please check module-id status.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            string mdlBefore = txtBefore.Text;
            string mdlAfter = dtModule.Rows[0]["module_id"].ToString();

            // 重複確認処理を行う
            TfSQL tf = new TfSQL();
            string dbDuplicate = tf.sqlModuleDuplicateCheck(dtModule);
            if (dbDuplicate != string.Empty)
            {
                for (int i = 0; i < dgvModule.Rows.Count; ++i)
                {
                    if (dgvModule["module_id", i].Value.ToString() == dbDuplicate)
                        dgvModule["module_id", i].Style.BackColor = Color.Red;
                }
                // soundAlarm();
                btnReplace.Enabled = false;
                return;
            }

            // 更新処理
            string sql = "update t_module set " +
                "module_id ='" + (string)dtModule.Rows[0]["module_id"].ToString() + "', " +
                "lot ='" + dtModule.Rows[0]["lot"].ToString() + "', " +
                "tester_id ='" + dtModule.Rows[0]["tester_id"].ToString() + "', " +
                "test_result ='" + dtModule.Rows[0]["test_result"].ToString() + "', " +
                "test_date ='" + dtModule.Rows[0]["test_date"].ToString() + "', " +
                "r_mode ='" + dtModule.Rows[0]["r_mode"].ToString() + "' " +
                "where module_id = '" + mdlBefore + "' and tray_id ='" + trayId + "'";
            System.Diagnostics.Debug.Print(sql);
            bool res = tf.sqlExecuteNonQuery(sql, false);

            if (res)
            {
                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                btnReplace.Enabled = false;
                txtAfter.Enabled = false;
                txtRow.Text = string.Empty;
                this.Focus();
                MessageBox.Show("The replacement was successful.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
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

        //閉じるボタンやショートカットでの終了を許さない
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x112;
            const long SC_CLOSE = 0xF060L;
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & 0xFFF0L) == SC_CLOSE) { return; }
            base.WndProc(ref m);
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

        // 他のフォームとの整合性を取るため、キャンセルボタンを設ける
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}