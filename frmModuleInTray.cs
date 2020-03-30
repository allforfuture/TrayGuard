using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Security.Permissions;
using System.Runtime.InteropServices;
using System.Linq;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;   //正規表現
using System.Threading.Tasks;
using System.Threading;

namespace TrayGuard
{
    public partial class frmModuleInTray : Form
    {
        //親フォームfrmTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        // プリント用テキストファイルの保存用フォルダを、基本設定ファイルで設定する
        string appconfig = System.Environment.CurrentDirectory + @"\info.ini";
        string productconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\tray_guard_desktop.ini";
        string outPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\NTRS Log\";

        List<Task> tasks = new List<Task>();    //Just do something when  tasks.waitall is finish,
        bool Flag = true;   //write to db when ‘Flag’ is true, else can't
        List<string> errModule = new List<string>();
        List<string> errBin = new List<string>();

        //その他、非ローカル変数
        DataTable dtModule;
        DataTable dtLot;
        DataTable dtModuleBeforeRefill;
        int countBeforeRefill;
        bool formAddMode;
        bool formRefillMode;
        bool formReturnMode;
        bool formPartialCancelMode;
        bool formReprintMode;
        bool trayIdCanceled;
        bool trayIdPacked;
        string totalSwitch;
        string userRole;
        string bin = string.Empty;
        int okCount;



        //int capacity = 24;

        //20190228修改将固定的28改成配置文件
        //int capacity = 28;
        int capacity = int.Parse(TfSQL.readIni_static("OK COUNT", "CAPACITY", Environment.CurrentDirectory + @"\form.ini"));


        string maxLot;
        DateTime registerDate;
        bool sound;
        int position;
        //string OK2ShipCheckSwitch;
        //string forcedNGSwitch;

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
        public frmModuleInTray()
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
        private void frmModule_Load(object sender, EventArgs e)
        {
            this.Text = this.Text + " " + Assembly.GetExecutingAssembly().GetName().Version;
            //forcedNGSwitch = readIni("MODULE-DATA MATCHING", "FORCED NG SWITCH", appconfig);
            //OK2ShipCheckSwitch = readIni("MODULE-DATA MATCHING", "OK2SHIP CHECK SWITCH", appconfig);
            //totalSwitch = (OK2ShipCheckSwitch == "OFF" && forcedNGSwitch == "OFF") ? "OFF" : "ON";

            // ログ用フォルダの作成
            if (!Directory.Exists(outPath)) Directory.CreateDirectory(outPath); 

            // 当フォームの表示場所を指定
            this.Left = 250;
            this.Top = 20;
            if (position == 2) { this.Left = 400; this.Top = 30; }
            else if (position == 3) { this.Left = 450; this.Top = 40; }
            else if (position == 4) { this.Left = 500; this.Top = 50; }

            // 各種処理用のテーブルを生成し、データを読み込む
            dtModule = new DataTable();
            defineModuleTable(ref dtModule);
            if (!formAddMode) readModuleInfo(ref dtModule);

            // グリットビューの更新
            updateDataGridViews(dtModule, ref dgvModule);

            // 追加モード、閲覧モードの切り替え
            changeMode();

            // バッチコンボボックスの設定
            setShiftComboBox();
        }

        // サブプロシージャ： テスター用ＳＥＬＥＣＴケース句の作成
        private string makeSqlCaseClause(string criteria)
        {
            string sql = " case ";
            foreach (string c in criteria.Split(','))
            { sql += "when c.process_cd like " + VBS.Left(c, c.Length - 1) + "%' then " + c + " ";  };
            sql += "else c.process_cd end as tester_id ";
            System.Diagnostics.Debug.Print(sql);
            return sql;
        }

        // サブプロシージャ： テスター用ＷＨＥＲＥ句の作成
        private string makeSqlWhereClause(string criteria)
        {
            string sql = " where ";
            foreach (string c in criteria.Split(','))
            { sql += "c.process_cd like " + VBS.Left(c, c.Length - 1) + "%' or "; };
            sql = VBS.Left(sql, sql.Length - 3);
            System.Diagnostics.Debug.Print(sql);
            return sql;
        }

        // サブプロシージャ： バッチコンボボックスの設定
        private void setShiftComboBox()
        {
            string sql = "select content from t_criteria where criteria = 'BIN_SHIFT' order by content";
            TfSQL tf = new TfSQL();
            tf.getComboBoxData(sql, ref cmbBinShift);
        }

        // サブプロシージャ： 追加モード、閲覧モードの切り替え
        private void changeMode()
        {
            // 追加モードの場合（充填モードの場合は、ＤＥＬＥＴＥＳＥＬＥＣＴＩＯＮは無効）
            if (formAddMode)
            {
                cmbBinShift.Enabled = formRefillMode ? false : true; 
                txtModuleId.Enabled = true;

                btnRegisterTray.Enabled = false;
                btnCancelTray.Enabled = false;
                btnReplaceModule.Enabled = false;
                btnReprintLabel.Enabled = false;
                btnRefillTray.Enabled = false;
                btnClose.Enabled = true;
                btnDeleteSelection.Enabled = formRefillMode ? false : true;
                //btnChangeCapacity.Enabled = (userRole == "super" && txtLoginDept.Text == "PC") ? true : false;
                btnChangeCapacity.Enabled = (userRole == "super") ? true : false;
            }
            // 閲覧モードの場合
            else
            {
                cmbBinShift.Enabled = false;
                txtModuleId.Enabled = false;

                btnRegisterTray.Enabled = false;
                btnCancelTray.Enabled = false;
                btnReplaceModule.Enabled = false;
                // 2016.08.10 FUJII  再プリントモードを追加（frmTrayにて、モジュールＩＤをキーとして、トレーＩＤを特定した場合のみ、再印刷可能）
                btnReprintLabel.Enabled = trayIdCanceled ? false : ((formReprintMode && userRole == "super") ? true : false); 

                btnRefillTray.Enabled = (trayIdCanceled || trayIdPacked || txtLoginDept.Text == "MFG") ? false : true; 
                btnClose.Enabled = true;
                btnDeleteSelection.Enabled = false;
                //btnChangeCapacity.Enabled = (userRole == "super" && txtLoginDept.Text == "PC") ? true : false;
                btnChangeCapacity.Enabled = (userRole == "super") ? true : false;
                if (userRole == "super")
                {
                    btnCancelTray.Enabled = (trayIdCanceled || trayIdPacked) ? false : true;
                    btnReplaceModule.Enabled = (trayIdCanceled || trayIdPacked) ? false : true; 
                }
            }

            // 一部キャンセルモードの場合、DELETE SELECTION 以外は無効
            if (formPartialCancelMode)
            {
                cmbBinShift.Enabled = false;
                txtModuleId.Enabled = false;

                btnRegisterTray.Enabled = false;
                btnCancelTray.Enabled = false;
                btnReplaceModule.Enabled = false;
                btnReprintLabel.Enabled = false;
                btnRefillTray.Enabled = false;
                btnClose.Enabled = true;
                btnDeleteSelection.Enabled = true;
                btnChangeCapacity.Enabled = false;
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
            bool addMode, bool returnMode, bool canceled, bool packed, int pos, bool partialCancelMode, bool reprintMode)
        {
            txtTrayId.Text = tid;
            dtpRegisterDate.Value = tdate;
            txtLoginName.Text = uname;
            txtLoginDept.Text = udept;
            userRole = urole;
            if (!addMode) cmbBinShift.Text = shift;
            formAddMode = addMode;
            formReturnMode = returnMode;
            trayIdCanceled = canceled;
            trayIdPacked = packed;
            position = pos;
            formPartialCancelMode = partialCancelMode;
            formReprintMode = reprintMode;
        }

        // サブプロシージャ：データテーブルの定義
        private void defineModuleTable(ref DataTable dt)
        {
            dt.Columns.Add("module_id", typeof(string));
            dt.Columns.Add("lot", typeof(string));
            dt.Columns.Add("bin", typeof(string));
            dt.Columns.Add("tester_id", typeof(string));
            dt.Columns.Add("test_result", typeof(string));
            dt.Columns.Add("test_date", typeof(DateTime));
            dt.Columns.Add("r_mode", typeof(string));
        }

        // サブプロシージャ：ＤＢからデータテーブルへの読み込み
        private void readModuleInfo(ref DataTable dt)
        {
            dt.Rows.Clear();
            string sql = "select module_id, lot, bin, tester_id, test_result, test_date, r_mode " +
                "from t_module where tray_id='" + txtTrayId.Text + "'";
            TfSQL tf = new TfSQL();
            System.Diagnostics.Debug.Print(sql);
            tf.sqlDataAdapterFillDatatableFromTrayGuardDb(sql, ref dt);
        }

        // サブプロシージャ：データグリットビューの更新
        private void updateDataGridViews(DataTable dt1, ref DataGridView dgv1)
        {
            // データグリットビューへＤＴＡＡＴＡＢＬＥを格納
            dgv1.DataSource = dt1;
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // テスト結果がＦＡＩＬまたはレコードなしのシリアルをマーキングする（マッチングスイッチ、オンの場合）
            if (totalSwitch == "ON") colorViewForFailAndBlank(ref dgv1);

            // 重複レコード、および１セル２重入力をマーキングする
            colorViewForDuplicateSerial(ref dgv1);

            //行ヘッダーに行番号を表示する
            for (int i = 0; i < dgv1.Rows.Count; i++) dgv1.Rows[i].HeaderCell.Value = (i + 1).ToString();

            //行ヘッダーの幅を自動調節する
            dgv1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            // 一番下の行を表示する（インライン）
            if (dgv1.Rows.Count >= 1) dgv1.FirstDisplayedScrollingRowIndex = dgv1.Rows.Count - 1;
            
            // ロット集計グリッドビューを更新し、数量の最も多いロットを保持する
            maxLot = updateLotSummary(dt1);

            // レコード数とトレイ容量による、ボタンとコントロールの制御、（追加モードの場合）
            okCount = getOkCount(dt1);  //現在の一時登録件数を変数へ保持
            if (formAddMode)
            {
                if (okCount == capacity)
                {
                    cmbBinShift.Enabled = formRefillMode ? false : true;
                    txtModuleId.Enabled = false;
                    btnRegisterTray.Enabled = true;
                    btnDeleteSelection.Enabled = formRefillMode ? false : true;
                    //btnChangeCapacity.Enabled = (userRole == "super" && txtLoginDept.Text == "PC") ? true : false;
                    btnChangeCapacity.Enabled = (userRole == "super") ? true : false;
                }
                else
                {
                    cmbBinShift.Enabled = formRefillMode ? false : true;
                    txtModuleId.Enabled = true;
                    btnRegisterTray.Enabled = false;
                    btnDeleteSelection.Enabled = formRefillMode ? false : true;
                    //btnChangeCapacity.Enabled = (userRole == "super" && txtLoginDept.Text == "PC") ? true : false;
                    btnChangeCapacity.Enabled = (userRole == "super") ? true : false;
                    txtModuleId.SelectAll(); // 連続スキャン用
                }
            }

            //現在一時登録件数を表示
            txtOkCount.Text = okCount.ToString() + "/" + capacity.ToString();
            dgvModule.ClearSelection();
            dgvModule.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        /// <summary>
        /// update dataGridView,and check sn validation
        /// </summary>
        /// <param name="dr"></param>
        private void updateDataGridViews(DataRow dr)
        {
            /*
            DataGridViewRow row = new dgv;
            row = (DataGridViewRow)dr;
            row.Cells[0].Value = dr["module_id"];
            row.Cells[1].Value = dr["lot"];
            row.Cells[2].Value = dr["bin"];
            row.Cells[3].Value = dr["tester_id"];
            row.Cells[4].Value = dr["test_result"];
            row.Cells[5].Value = dr["test_date"];
            row.Cells[6].Value = dr["r_mode"];
            */
            //((DataTable)dgvModule.DataSource).Rows.Add(dr.ItemArray);
            dgvModule.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvModule.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // if totalSwich=on then check  is "result" have ng item
            if (totalSwitch == "ON") colorViewForFailAndBlank(ref dgvModule);

            // check sn is repeat
            colorViewForDuplicateSerial(ref dgvModule);

            //dgvModule headtitle= each  headcell of row
            for (int i = 0; i < dgvModule.Rows.Count; i++) dgvModule.Rows[i].HeaderCell.Value = (i + 1).ToString();

            //dgvModule auto row resize
            dgvModule.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            // set dgvModule row index
            if (dgvModule.Rows.Count >= 1) dgvModule.FirstDisplayedScrollingRowIndex = dgvModule.Rows.Count - 1;

            // update LotSumary, return maxLot
            maxLot = updateLotSummary(dtModule);

            // get count of ok sn in dtmodule
            okCount = getOkCount(dtModule);  //現在の一時登録件数を変数へ保持
            if (formAddMode)
            {
                if (okCount == capacity)
                {
                    cmbBinShift.Enabled = formRefillMode ? false : true;
                    txtModuleId.Enabled = false;
                    btnRegisterTray.Enabled = true;
                    btnDeleteSelection.Enabled = formRefillMode ? false : true;
                    //btnChangeCapacity.Enabled = (userRole == "super" && txtLoginDept.Text == "PC") ? true : false;
                    btnChangeCapacity.Enabled = (userRole == "super") ? true : false;
                }
                else
                {
                    cmbBinShift.Enabled = formRefillMode ? false : true;
                    txtModuleId.Enabled = true;
                    btnRegisterTray.Enabled = false;
                    btnDeleteSelection.Enabled = formRefillMode ? false : true;
                    //btnChangeCapacity.Enabled = (userRole == "super" && txtLoginDept.Text == "PC") ? true : false;
                    btnChangeCapacity.Enabled = (userRole == "super") ? true : false;
                    txtModuleId.SelectAll(); // 連続スキャン用
                }
            }

            //現在一時登録件数を表示
            txtOkCount.Text = okCount.ToString() + "/" + capacity.ToString();
            dgvModule.ClearSelection();
            dgvModule.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        // Get count of ok sn in dtModule
        private int getOkCount(DataTable dt)
        {
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "module_id", "test_result" });
            DataRow[] dr;
            if (totalSwitch == "ON") dr = distinct.Select("test_result in ('PASS','n/a')"); //distinct.Select(); ＦＡＩＬも流す、初期段階の設定
            else dr = distinct.Select();  // マッチングスイッチオフの場合は、テスト結果に関係なく行数をカウントする

            return dr.Length;
        }

        // Update LotSummary ,return maxLot
        public string updateLotSummary(DataTable dt)
        {
            if (dt.Rows.Count <= 0)
            {
                dtLot = new DataTable();
                dgvLotSummary.DataSource = dtLot;
                return string.Empty;
            }
                
            var query = from e in dt.AsEnumerable()
                        group e by e.Field<string>("lot") into Summary
                        orderby Summary.Count() descending
                        select new {lot = Summary.Key, qty = Summary.Count()};

            // グリットビューデータソースの削除と、新データテーブルの生成
            dgvLotSummary.DataSource = null;
            dgvLotSummary.Refresh();
            dtLot = new DataTable();
            // 列の追加
            foreach (var q in query) dtLot.Columns.Add(q.lot, typeof(int));
            dtLot.Columns.Add("total", typeof(int));
            // 行の追加
            dtLot.Rows.Add(); 
            foreach (var q in query) dtLot.Rows[0][q.lot] = q.qty;
            dtLot.Rows[0]["total"] = query.Sum(a => a.qty);
 
            dgvLotSummary.DataSource = dtLot;
            return query.First().lot;
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

/*------------------------------------------------------------------------
 * ☆入り口　：シリアルがスキャンされた時の処理
 * ----------------------------------------------------------------------*/
        private void txtModuleId_KeyDown(object sender, KeyEventArgs e)
        {
            // エンターキーの場合、テキストボックスの桁数が１７桁または２４桁の場合のみ、処理を行う
            if (e.KeyCode != Keys.Enter) return;
            if (txtModuleId.Text.Length != 17 && txtModuleId.Text.Length != 24) return;
            //正規表現で(前から2桁)大文字なのかを確認、それ以外は禁止。例：FH171562074HM9X2VJQYA2N7はOK,fh17はNG
            Regex reg=new Regex("^[A-Z]{2}");
            if (!reg.IsMatch(txtModuleId.Text))
            {
                //不匹配大写和数字
                MessageBox.Show("Small letter is forbidden!","Notice");
                return;
            }
            
            if (frmTray.userRole != "super")//权限
            {
                string codeDate = txtModuleId.Text.Substring(3, 4);
                #region 检验是否全数字
                //n位定长的数字：^\d{n}$
                if (!new Regex(@"^\d{4}$").IsMatch(codeDate))
                {
                    MessageBox.Show("Bad date code format for serial number!", "Notice");
                    return;
                }
                #endregion
                DateTime pd = GetProductDate(codeDate);
                if (pd < DateTime.Today.AddDays(-180))
                {
                    MessageBox.Show("The serial number has a production date of more than 180 days!", "Notice");
                    return;
                }
            }

            // ＢＡＳＥシリアルからＨＯＯＰシリアルを取得し（ステップ１）、両方のシリアルに該当する、テスト結果・プロセス名・テスト日時を取得する（ステップ２、ステップ３）
            TfSQL tf = new TfSQL();
            DataTable dt = new DataTable();
            string log = string.Empty;
            string module = txtModuleId.Text;
            string mdlShort = VBS.Left(module, 17);
            string mdlNtrs = string.Empty;
            string mdlOK2ShipResult = string.Empty;
            // 2016.08.18 FUJIKI FORCED NG CHECK を追加
            string mdlForcedNGResult = string.Empty;
            // 2016.08.18 FUJIKI FORCED NG CHECK を追加
            string scanTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string displayAll = string.Empty;   // ログ用
            DataRow dr = dtModule.NewRow();
            string textResult = "PASS";
            string mdlSerialBin = VBS.Mid(module, 12,4);    //add check of BinA and BinB 
            string textSelectBin = string.Empty;

            
            // 2017.03.08 FUJII  シリアル構成要素のチェックをオフ
            //トレーＩＤの新規採番
            //先ずは、シリアルの構成要素のパターンが適正か、ユーザーデスクトップの設定ファイルを使用して確認する
            //設定ファイルにSN CHECK FUNCTION SWITCHの設定を確認（2017.08.09 呉）
            string SN_ChkSwitch=tf.readIni("SN CHECK", "SN CHECK FUNCTION SWITCH", appconfig);
            if (SN_ChkSwitch == "ON")
            {
                string matchResult = matchSerialNumberingPattern(module);
                if (matchResult != string.Empty)
                {
                    MessageBox.Show(matchResult + " does not match with desktop file's setting.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            };


            // メモリ上のテーブルにレコードを追加
            dr["module_id"] = module;
            dr["lot"] = VBS.Left(module, 8);
            dr["bin"] = mdlSerialBin; 
            dr["tester_id"] = displayAll;
            dr["test_result"] = textResult;
            dr["test_date"] = DateTime.ParseExact(scanTime, "yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture); ;
            dr["r_mode"] = formReturnMode ? "T" : "F";
            DataRow dr1 = dr;
            dtModule.Rows.Add(dr);

            validationSN(dr);

           

            // アプリケーションフォルダに、日付とテスト結果のログを付ける
            log = Environment.NewLine + scanTime + "," + module + "," + displayAll + ":" + textResult;
            // log = Environment.NewLine + earlyTime + "," + module + "," + displayAll;

            // 同日日付のファイルが存在する場合は追記し、存在しない場合はファイルを作成追記する
            try
            {
                string outFile = outPath + DateTime.Today.ToString("yyyyMMdd") + ".txt";
                System.IO.File.AppendAllText(outFile, log, System.Text.Encoding.GetEncoding("UTF-8"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // データグリットビューの更新
            updateDataGridViews(dtModule, ref dgvModule);
            //updateDataGridViews(dr1);
            

            //トレー登録の準備確認
            PreparePrintWhen24();
            
        }
        
        DateTime GetProductDate(string dateCode)
        {
            DateTime today = DateTime.Today;
            string Y = dateCode.Substring(0, 1);
            string W = dateCode.Substring(1, 2);
            string D = dateCode.Substring(3);
            //年
            while (today.Year.ToString().Substring(3) != Y)
                today = today.AddYears(-1);
            //周
            GregorianCalendar gc = new GregorianCalendar();
            while (true)
            {
                string todatWeek = string.Format("{0:d2}", gc.GetWeekOfYear(today, CalendarWeekRule.FirstDay, DayOfWeek.Sunday));
                if (Convert.ToUInt16(todatWeek) > Convert.ToUInt16(W))
                    today = today.AddDays(-7);
                else if (Convert.ToUInt16(todatWeek) < Convert.ToUInt16(W))
                    today = today.AddDays(7);
                else if (Convert.ToUInt16(todatWeek) == Convert.ToUInt16(W))
                    break;
            }
            //日
            while (true)
            {
                //(int)today.DayOfWeek + 1
                if ((int)today.DayOfWeek + 1 > Convert.ToUInt16(D))
                    today = today.AddDays(-1);
                else if ((int)today.DayOfWeek + 1 < Convert.ToUInt16(D))
                    today = today.AddDays(1);
                else if ((int)today.DayOfWeek + 1 == Convert.ToUInt16(D))
                    break;
            }
            return today;
        }

        // サブプロシージャ：シリアルの構成要素のパターンが適正か、ユーザーデスクトップの設定ファイルを使用して確認する
        private string matchSerialNumberingPattern(string serial)
        {
            string result = string.Empty;

            // 設定ファイル ＬＩＮＥ ＝ ＸＸＸ の場合は、検証しない
            // if (line == "XXX") return string.Empty;
            if      (!plant.Equals("XXX")      && VBS.Mid(serial,  1, 3) != plant)        return "Plant '"       + VBS.Mid(serial,  1, 3) + "'";
            else if (!year.Equals("XXX")       && VBS.Mid(serial,  4, 1) != year)         return "Year '"        + VBS.Mid(serial,  4, 1) + "'";
            else if (!week.Equals("XXX")       && VBS.Mid(serial,  5, 2) != week)         return "Week '"        + VBS.Mid(serial,  5, 2) + "'";
            else if (!day.Equals("XXX")        && VBS.Mid(serial,  7, 1) != day)          return "Day '"         + VBS.Mid(serial,  7, 1) + "'";
            else if (!line.Equals("XXX")       && VBS.Mid(serial,  8, 1) != line)         return "Line '"        + VBS.Mid(serial,  8, 1) + "'";
            else if (!eeee.Equals("XXX")       && VBS.Mid(serial, 12, 4) != eeee)         return "4E '"          + VBS.Mid(serial, 12, 4) + "'";
            else if (!revision.Equals("XXX")   && VBS.Mid(serial, 16, 1) != revision)     return "Revision '"    + VBS.Mid(serial, 16, 1) + "'";
            else if (!mass.Equals("XXX")       && VBS.Mid(serial, 19, 1) != mass)         return "Mass '"        + VBS.Mid(serial, 19, 1) + "'";
            else if (!flexure.Equals("XXX")    && VBS.Mid(serial, 20, 1) != flexure)      return "Flexure '"     + VBS.Mid(serial, 20, 1) + "'";
            else if (!cover_base.Equals("XXX") && VBS.Mid(serial, 21, 1) != cover_base)   return "Cover/base '"  + VBS.Mid(serial, 21, 1) + "'";
            else if (!dframe.Equals("XXX")     && VBS.Mid(serial, 22, 1) != dframe)       return "D-Frame '"     + VBS.Mid(serial, 22, 1) + "'";
            else if (!fpc.Equals("XXX")        && VBS.Mid(serial, 23, 1) != fpc)          return "FPC '"         + VBS.Mid(serial, 23, 1) + "'";
            else if (!shift.Equals("XXX")      && VBS.Mid(serial, 24, 1) != shift)        return "Shift '"       + VBS.Mid(serial, 24, 1) + "'";
            else return string.Empty;
        }

        // ビューモードで再印刷を行う
        private void btnPrint_Click(object sender, EventArgs e)
        {
            TfSato tfs = new TfSato();
            tfs.printStart("tray", txtTrayId.Text, dtLot, string.Empty, dtpRegisterDate.Value, string.Empty, string.Empty, 1, string.Empty);
        }

        // 登録ボタン押下時、各種確認、ボックスＩＤの発行、シリアルの登録、バーコードラベルのプリントを行う
        // txtOkCount_TextChangedイベントに変更　2017/08/16　呉
        private void btnRegisterTray_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This button is no more used.", "Notice", MessageBoxButtons.OK,MessageBoxIcon.Information);
            /*
            if (cmbBinShift.Text == string.Empty)
            {
                MessageBox.Show("Please select Shift.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }
            if (getOkCount(dtModule) != dtModule.Rows.Count)
            {
                MessageBox.Show("Module is not 24.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            // 登録処理中は、閉じるボタンを除き、すべてのコントロールを無効にする
            cmbBinShift.Enabled = false;
            txtModuleId.Enabled = false;
            btnRegisterTray.Enabled = false;
            btnDeleteSelection.Enabled = false;
            btnChangeCapacity.Enabled = false;

            if (formRefillMode && countBeforeRefill >= dtModule.Rows.Count)
            {
                MessageBox.Show("Refill is not correct.  Please close this window.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            //一時テーブルのシリアル全てについて、ＤＢテーブルに既に登録がないか、確認する
            //（ＲＥＦＩＬＬモードの場合は、追加分のモジュールのみ確認する）
            DataTable dtTarget = new DataTable();
            if (formRefillMode) dtTarget = subtractTable(dtModule, dtModuleBeforeRefill);
            else dtTarget = dtModule;

            TfSQL tf = new TfSQL();
            string dbDuplicate = tf.sqlModuleDuplicateCheck(dtTarget);
            if (dbDuplicate != string.Empty)
            {
                for (int i = 0; i < dgvModule.Rows.Count; ++i)
                { 
                    if (dgvModule["module_id", i].Value.ToString() == dbDuplicate)
                        dgvModule["module_id", i].Style.BackColor = Color.Red;
                }
                soundAlarm();

                // ＲＥＦＩＬＬモードの場合は、閉じる以外の選択肢を与えない
                if (!formRefillMode)
                {
                    btnDeleteSelection.Enabled = true;
                }
                return;
            }

            // 充填モードか否かで、呼び出すプロシージャを選択して実行する
            if (!formRefillMode) registerModuleNormalMode();
            else registerModuleRefillMode();
            */
        }

        // サブプロシージャ：ＲＥＦＩＬＬモードの新規追加分のみを取得する
        private DataTable subtractTable(DataTable dtAll, DataTable dtSub)
        {
            DataTable dtAfter = (from a in dtAll.AsEnumerable()
                                 join s in dtSub.AsEnumerable()
                                 on a["module_id"].ToString() equals s["module_id"].ToString()
                                 into g
                                 where g.Count() == 0
                                 select a).CopyToDataTable();
            return dtAfter;
        }

        // サブプロシージャ：モジュールの登録、通常モード
        private void registerModuleNormalMode()
        {
            // 2016.08.10 FUJII  トレーＩＤの新規採番プロシージャ「GetNewTrayId」を、トランザクション処理バージョンへ変更
            //トレーＩＤの新規採番
            //string trayNew = getNewTrayId(txtLoginDept.Text, txtLoginName.Text, VBS.Mid(maxLot, 8, 1), cmbBinShift.Text);
            TfSQL tf = new TfSQL();
            string trayNew = tf.sqlGetNewTrayId(txtLoginDept.Text, txtLoginName.Text, VBS.Mid(maxLot, 8, 1), cmbBinShift.Text, maxLot, dtLot, ref registerDate);
            if (trayNew == string.Empty)
            {
                MessageBox.Show("An error happened in the tray id issuing process.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //モジュールテーブルへ、データテーブルのレコードを一括登録する
            bool res = tf.sqlMultipleInsertModule(dtModule, trayNew);

            if (res)
            {
                //ラベルのプリントアウト
                TfSato tfs = new TfSato();
                tfs.printStart("tray", trayNew, dtLot, string.Empty, dtpRegisterDate.Value, string.Empty, string.Empty, 1, string.Empty);

                //登録済みの状態を表示
                txtTrayId.Text = trayNew;
                dtpRegisterDate.Value = registerDate;

                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                this.Focus();
                MessageBox.Show("Tray ID: " + trayNew + Environment.NewLine +
                    "and its modules were registered.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //ユーザーによるメッセージボックス確認後の処理
                txtTrayId.Text = String.Empty;
                txtModuleId.Text = String.Empty;
                dtModule.Clear();
                //capacity = 24;
                capacity = int.Parse(TfSQL.readIni_static("OK COUNT", "CAPACITY", Environment.CurrentDirectory + @"\form.ini"));
                updateDataGridViews(dtModule, ref dgvModule);
            }
        }

        // サブプロシージャ：モジュールの登録、充填モード
        private void registerModuleRefillMode()
        {
            string trayCurrent = txtTrayId.Text;

            //モジュールテーブルの補充前レコードを一旦削除し、トレーテーブルの更新履歴を追加する
            bool res1 = updateTrayIdAndDeleteModule(trayCurrent, txtLoginDept.Text, txtLoginName.Text);
            if (!res1)
            {
                MessageBox.Show("An error happened in the tray id updating process.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //モジュールテーブルへ、データテーブルのレコードを一括登録する
            TfSQL tf = new TfSQL();
            bool res2 = tf.sqlMultipleInsertModule(dtModule, trayCurrent);

            if (res2)
            {
                //ラベルのプリントアウト
                TfSato tfs = new TfSato();
                tfs.printStart("tray", txtTrayId.Text, dtLot, string.Empty, dtpRegisterDate.Value, string.Empty, string.Empty, 1, string.Empty);

                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                this.Focus();
                MessageBox.Show("Tray ID: " + trayCurrent + Environment.NewLine +
                    "and its modules were updated.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //ユーザーによるメッセージボックス確認後の処理
                formAddMode = false;
                formRefillMode = false;
                updateDataGridViews(dtModule, ref dgvModule);

                //閉じるボタンみ有効として、閉じることを促す
                cmbBinShift.Enabled = false;
                txtModuleId.Enabled = false;
                btnRegisterTray.Enabled = false;
                btnDeleteSelection.Enabled = false;
                btnChangeCapacity.Enabled = false;
            }
        }

        // サブプロシージャ：トレー情報の更新（モジュール情報削除処理付き）
        private bool updateTrayIdAndDeleteModule(string trayId, string udept, string uname)
        {
            // モジュール数量、複数ロット区分、更新日、更新部署、更新ユーザーについて更新する
            int qty = (int)dtLot.Rows[0]["total"];
            string multiLot = dtLot.Columns.Count >= 3 ? "T" : "F";
            string sql = "update t_tray set " +
                "qty ='" + qty + "', " +
                "update_date ='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "', " +
                "up_dept ='" + udept + "', " +
                "up_user ='" + uname + "', " +
                "multi_lot ='" + multiLot + "' " +
                "where tray_id ='" + trayId + "'";

            // モジュールレコードは、下記サブプロシージャ内で一旦キャンセルし、別プロシージャで一括登録）
            System.Diagnostics.Debug.Print(sql);
            TfSQL tf = new TfSQL();
            bool res = tf.sqlUpdateModuleInPack(trayId, sql);
            return res;
        }

        // サブプロシージャ：トレー情報の更新
        private bool updateTrayId(string trayId, string udept, string uname)
        {
            // モジュール数量、複数ロット区分、更新日、更新部署、更新ユーザーについて更新する
            int qty = (int)dtLot.Rows[0]["total"];
            string multiLot = dtLot.Columns.Count >= 3 ? "T" : "F";
            string sql = "update t_tray set " +
                "qty ='" + qty + "', " +
                "update_date ='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "', " +
                "up_dept ='" + udept + "', " +
                "up_user ='" + uname + "', " +
                "multi_lot ='" + multiLot + "' " +
                "where tray_id ='" + trayId + "'";

            System.Diagnostics.Debug.Print(sql);
            TfSQL tf = new TfSQL();
            bool res = tf.sqlExecuteNonQuery(sql, false);
            return res;
        }

        // 一時テーブルの選択された複数レコードを、一括消去させる
        private void btnDeleteSelection_Click(object sender, EventArgs e)
        {
            if (dtModule.Rows.Count <= 0) return;

            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
            if (dgvModule.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2)
            {
                MessageBox.Show("Please select range with only one column.", "Notice",　MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            // 部分キャンセルモードでない時は、通常処理（モジュールデータテーブルの選択行削除）
            if (!formPartialCancelMode)
            {
                DialogResult result = MessageBox.Show("Do you really want to delete the selected rows?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.No) return;

                foreach (DataGridViewCell cell in dgvModule.SelectedCells)
                {
                    dtModule.Rows[cell.RowIndex].Delete();
                }
                dtModule.AcceptChanges();
                updateDataGridViews(dtModule, ref dgvModule);
                txtModuleId.Focus();
                txtModuleId.SelectAll();
            }

            // 部分キャンセルモードの時は、ＤＢモジュールテーブルの行削除
            else if (formPartialCancelMode)
            {
                if (dgvModule.SelectedCells.Count == dtModule.Rows.Count)
                {
                    MessageBox.Show("If you wish to delete all modules, please cancel the tray.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    return;
                }

                DialogResult result = MessageBox.Show("Do you really want to delete the selected rows?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.No) return;

                btnDeleteSelection.Enabled = false;
                string sql1 = "delete from t_module where module_id in ('";
                string sql2 = string.Empty;
                foreach (DataGridViewCell cell in dgvModule.SelectedCells)
                {
                    sql2 += dtModule.Rows[cell.RowIndex]["module_id"].ToString() + "','";
                    dtModule.Rows[cell.RowIndex].Delete();
                }
                string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
                TfSQL tf = new TfSQL();
                bool res = tf.sqlExecuteNonQuery(sql3, false);

                if (res)
                {
                    dtModule.AcceptChanges();
                    updateDataGridViews(dtModule, ref dgvModule);

                    string updateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    string sql4 = "update t_tray set update_date = '" + updateTime + "', up_dept = '" + txtLoginDept.Text + "', " + 
                        "up_user = '" + txtLoginName.Text + "', lot = '" + dtLot.Columns[0].ColumnName + "', " +
                        "qty = '" + dtLot.Rows[0]["total"].ToString() + "', multi_lot = '" + (dtLot.Columns.Count >= 3 ? "T" : "F") + "' " +
                        "where tray_id = '" + txtTrayId.Text + "'";
                    System.Diagnostics.Debug.Print(sql4);
                    tf.sqlExecuteNonQuery(sql4, false);

                    //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                    this.RefreshEvent(this, new EventArgs());
                    this.Focus();
                    MessageBox.Show("Partial cancel was successful.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                }
                else
                {
                    MessageBox.Show("Partial cancel failed.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                }
            }
        }

        // １トレーあたりのモジュール数を変更する（管理権限ユーザーのみ）、ボタンのクリックから呼び出し
        private void btnChangecapacity_Click(object sender, EventArgs e)
        {
            //if (userRole == "super" && txtLoginDept.Text == "PC")
            if (userRole == "super")
                changeCapacity(sender, e);
        }

        // １トレーあたりのモジュール数を変更する（管理権限ユーザーのみ）、テキストボックスのダブルクリックから呼び出し
        private void txtOkCount_DoubleClick(object sender, EventArgs e)
        {
            //if (userRole == "super" && txtLoginDept.Text == "PC")
            if (userRole == "super")
                changeCapacity(sender, e);
        }

        // サブプロシージャ：１トレーあたりのモジュール数を変更
        private void changeCapacity(object sender, EventArgs e)
        {
            //2016/08/05 仕様変更(24個以外は一切認めない)
            //2016/09/30 仕様変更(倉庫スーパーユーザーのみ、変更可能)
            //return;

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
                updateDataGridViews(dtModule, ref dgvModule);
            };

            fC.updateControls(capacity);
            fC.Show();
        }

        // スーパーユーザーに限り、登録済みのシリアルの置き換えができる（パック後は不可）
        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (dtModule.Rows.Count <= 0) return;

            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
            if (dgvModule.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2 ||
                    dgvModule.Rows.GetRowCount(DataGridViewElementStates.Selected) >= 2 ||
                    dgvModule.CurrentCell.ColumnIndex != 0)
            {
                MessageBox.Show("Please select only one serial number.", "Notice",　MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            // 置き換え用フォームが既に開いていないか、確認する
            if (TfGeneral.checkOpenFormExists("frmModuleReplace"))
            {
                MessageBox.Show("Please close or complete another form.", "Notice",　MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            string curSerial = dgvModule.CurrentCell.Value.ToString();
            int curRowIndex = dgvModule.CurrentRow.Index;
            frmModuleReplace fR = new frmModuleReplace();

            //置き換えフォームイベントをキャッチして、当フォームデータグリッドを更新し、トレーテーブルも更新する
            fR.RefreshEvent += delegate(object sndr, EventArgs excp)
            {
                //モジュールテーブル・データグリッドの更新
                readModuleInfo(ref dtModule);
                updateDataGridViews(dtModule, ref dgvModule);
                //トレーテーブルの更新、さらに、親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                updateTrayId(txtTrayId.Text, txtLoginDept.Text, txtLoginName.Text);
                this.RefreshEvent(this, new EventArgs());
                this.Focus();
            };

            fR.updateControls(txtTrayId.Text, curSerial, curRowIndex +1, formReturnMode, cmbBinShift.Text);
            fR.Show();
        }

        // スーパーユーザーに限り、登録済みトレーをキャンセルできる（パック後は不可）
        private void btnCancelTray_Click(object sender, EventArgs e)
        {
            string trayId = txtTrayId.Text;
            TfSQL tf = new TfSQL();
            if (!tf.sqlChkEffective(trayId))
                return;
            // 本当に削除してよいか、２重で確認する。
            DialogResult result1 = MessageBox.Show("Do you really want to cancel this tray?",　"Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result1 == DialogResult.No) return;

            DialogResult result2 = MessageBox.Show("Are you really sure? Please select NO if you are not sure.", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result2 == DialogResult.No) return;

            
            //キャンセルの前に、packされたかどうかを確認（2017/09/18 呉）
            if (tf.sqlChkPackofTray(trayId)!="OK") return;

            // キャンセルの実行
            bool res = tf.sqlCancelModuleInTray(trayId, txtLoginDept.Text, txtLoginName.Text);
            if (res)
            {
                //本フォームのデータグリットビュー更新
                dtModule.Clear();
                updateDataGridViews(dtModule, ref dgvModule);

                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                this.Focus();
                MessageBox.Show("Tray ID " + trayId + " and its modules were canceled.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // メッセージボックスの確認後、閉じる
                Close();
            }
            else
            {
                MessageBox.Show("Cancel process was not successful.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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
            // frmModuleReplace を閉じていない場合は、先に閉じるよう通知する
            if (TfGeneral.checkOpenFormExists("frmModuleReplace"))
            {
                MessageBox.Show("You need to close Replace form before canceling..", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            Close();
        }

        /// <summary>
        /// check is dtModule["test_result"] have ng item
        /// </summary>
        /// <param name="dgv"></param>
        private void colorViewForFailAndBlank(ref DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; ++i)
            {
                if (dgv["test_result", i].Value.ToString() == "FAIL" || dgv["test_result", i].Value.ToString() == "NG" || dgv["test_result", i].Value.ToString() == string.Empty)
                {
                    dgv["tester_id", i].Style.BackColor = Color.Red;
                    dgv["test_result", i].Style.BackColor = Color.Red;
                    dgv["test_date", i].Style.BackColor = Color.Red;
                    soundAlarm();
                }
                else
                {
                    dgv["tester_id", i].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                    dgv["test_result", i].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                    dgv["test_date", i].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                }
            }
        }

        /// <summary>
        /// check sn is exists dtModule
        /// </summary>
        /// <param name="dgv"></param>
        private void colorViewForDuplicateSerial(ref DataGridView dgv)
        {
            if (dgv.Rows.Count <= 0) return;

            DataTable dt = ((DataTable)dgv.DataSource).Copy();
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                string module = dgv["module_id", i].Value.ToString();
                DataRow[] dr = dt.Select("module_id = '" + module + "'");
                if (dr.Length >= 2)
                {
                    dgv["module_id", i].Style.BackColor = Color.Red;
                    soundAlarm();
                }
                else
                {
                    //dgv["module_id", i].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
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
        private void dgvOverall_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)dgvModule.DataSource;
            ExcelClass xl = new ExcelClass();
            xl.ExportToExcel(dt);
        }

        // ＬＯＴ集計グリッドビュー左のボタンを押下時、対象のモジュールをエクセルファイルへ出力する
        private void btnExportModule_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)dgvModule.DataSource;
            ExcelClass xl = new ExcelClass();
            // 2016.08.29 FUJII エクセルへの出力から、デスクトップＣＳＶへの出力へ変更
            xl.ExportToCsv(dt, System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\tray.csv");
        }

        // 規定数量に達していないトレイについて、モジュールを追加する（パック後は不可）
        private void btnRefillTray_Click(object sender, EventArgs e)
        {
            formAddMode = true;
            formRefillMode = true;
            changeMode();
            dtModuleBeforeRefill = dtModule.Copy();
            countBeforeRefill = dtModule.Rows.Count;
        }

        // ＢＩＮ・シフトを、マスターで指定する
        private void cmbShift_KeyDown(object sender, KeyEventArgs e)
        {
            // スーパーユーザーが使用中の場合のみ、ユーザーマスター変更フォームを開く
            if (formPartialCancelMode) return;
            if (e.KeyCode != Keys.Enter || userRole != "super") return;
            if (TfGeneral.checkOpenFormExists("frmMasterCriteria")) return;
            frmMasterCriteria fI = new frmMasterCriteria("BIN_SHIFT");
            //子イベントをキャッチして、データグリッドを更新する
            fI.RefreshEvent += delegate (object sndr, EventArgs excp)
            {
                //バッチコンボボックスの設定
                setShiftComboBox();
            };
            fI.Show();
        }

        // 各容器の容量を、マスターで指定する
        private void txtOkCount_KeyDown(object sender, KeyEventArgs e)
        {
            // スーパーユーザーが使用中の場合のみ、ユーザーマスター変更フォームを開く
            if (e.KeyCode != Keys.Enter || userRole != "super") return;
            if (TfGeneral.checkOpenFormExists("frmMasterCriteria")) return;
            frmMasterCriteria fI = new frmMasterCriteria("BIN_SHIFT");
            //子イベントをキャッチして、データグリッドを更新する
            fI.RefreshEvent += delegate (object sndr, EventArgs excp)
            {
                //バッチコンボボックスの設定
                setShiftComboBox();
            };
            fI.Show();
        }

        // データグリッドのモジュールＩＤを検索する
        private void frmModuleInTray_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // 既に同フォームが開かれている場合は、処理を行わない
            if (TfGeneral.checkOpenFormExists("frmModuleFind"))
            {
                MessageBox.Show("Please close or complete another form.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            frmModuleFind fC = new frmModuleFind();
            //子イベントをキャッチして、該当するセルにフォーカスする
            fC.RefreshEvent += delegate (object sndr, EventArgs excp)
            {
                string target = fC.returnTargetModule();
                for (int i = 0; i < dtModule.Rows.Count; i++)
                {
                    if (dtModule.Rows[i]["module_id"].ToString() == target) dgvModule.CurrentCell = dgvModule.Rows[i].Cells[0];
                }
            };
      
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

        //btnRegisterTrayボタンを押さなくても、トレーを登録して、ラベル印刷
        private void PreparePrintWhen24()
        {
            // シリアル番号重複なしのＰＡＳＳ個数を変数へ保持する
            okCount = getOkCount(dtModule);
            if (okCount.ToString() != capacity.ToString()) return;

            //他のチェック
            if (cmbBinShift.Text == string.Empty | cmbBinShift.Text.Length > 3) //シフト設定の文字列数をチェック追加
            {
                MessageBox.Show("Please select correct Shift.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }
            if (okCount != dtModule.Rows.Count)
            {
                MessageBox.Show("Module is not 24.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            // 登録処理中は、閉じるボタンを除き、すべてのコントロールを無効にする
            cmbBinShift.Enabled = false;
            txtModuleId.Enabled = false;
            btnRegisterTray.Enabled = false;
            btnDeleteSelection.Enabled = false;
            btnChangeCapacity.Enabled = false;

            if (formRefillMode && countBeforeRefill >= dtModule.Rows.Count)
            {
                MessageBox.Show("Refill is not correct.  Please close this window.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            //一時テーブルのシリアル全てについて、ＤＢテーブルに既に登録がないか、確認する
            //（ＲＥＦＩＬＬモードの場合は、追加分のモジュールのみ確認する）
            DataTable dtTarget = new DataTable();
            if (formRefillMode) dtTarget = subtractTable(dtModule, dtModuleBeforeRefill);
            else dtTarget = dtModule;

            //Waite all asyns method is excute finish
            Task.WaitAll(tasks.ToArray());
            if (!Flag)
            {
                MessageBox.Show("Please check this items of backgroun red", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                if (!formRefillMode)
                {
                    btnDeleteSelection.Enabled = true;
                }
                soundAlarm();
                return;
            }
            /*  before here had validation sn
            TfSQL tf = new TfSQL();
            string dbDuplicate = tf.sqlModuleDuplicateCheck(dtTarget);
            if (dbDuplicate != string.Empty)
            {
                for (int i = 0; i < dgvModule.Rows.Count; ++i)
                {
                    if (dgvModule["module_id", i].Value.ToString() == dbDuplicate)
                        dgvModule["module_id", i].Style.BackColor = Color.Red;

                }
                soundAlarm();

                // ＲＥＦＩＬＬモードの場合は、閉じる以外の選択肢を与えない
                if (!formRefillMode)
                {
                    btnDeleteSelection.Enabled = true;
                }
                return;
            }
            */
            // 充填モードか否かで、呼び出すプロシージャを選択して実行する
            if (!formRefillMode) registerModuleNormalMode();
            else registerModuleRefillMode();
        }

        //check sn by use asyns method
        private async  void validationSN(DataRow dr)
        {
            
            Task t= Task.Run(() =>
            {
                if (dtModule.Rows.Count > 0)
                {
                    //Thread.Sleep(10000);
                    //check last sn bin and first sn bin is different
                    if (dr["bin"].ToString() != dtModule.Rows[0]["bin"].ToString())
                    {
                        Flag = false;
                        errBin.Add(dr["bin"].ToString());
                    }
                    //check last sn  is exists had shop 
                    //DataGridViewRow row = dgvModule.Rows[dtModule.Rows.Count - 1];
                    //DataRow dr = (row.DataBoundItem as DataRowView).Row;
                    TfSQL tf = new TfSQL();
                    string dbDuplicate = tf.sqlModuleDuplicateCheck(dr);
                    if(dbDuplicate!=string.Empty)
                    {
                        Flag = false;
                        errModule.Add(dr["module_id"].ToString());
                    }

                    for (int i = 0; i < dgvModule.Rows.Count; i++)
                    {
                        foreach (string s1 in errModule)
                        {
                            if (dgvModule["module_id", i].Value.ToString() == s1)
                            {
                                dgvModule["module_id", i].Style.BackColor = Color.Red;
                            }
                        }
                        foreach (string s1 in errBin)
                        {
                            if (dgvModule["bin", i].Value.ToString() == s1)
                            {
                                dgvModule["bin", i].Style.BackColor = Color.Red;
                            }
                        }
                    }
                }
            });
            tasks.Add(t);
            await t;
        }
    }
}