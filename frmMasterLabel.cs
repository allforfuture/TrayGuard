using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Security.Permissions;
using System.Collections;
using Npgsql;
using System.Runtime.InteropServices;

namespace TrayGuard
{
    public partial class frmMasterLabel : Form
    {
        //親フォームfrmModuleInTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        //その他非ローカル変数
        NpgsqlConnection connection;
        NpgsqlCommand command;
        NpgsqlDataAdapter adapter;
        NpgsqlCommandBuilder cmdbuilder;
        DataSet ds;
        DataTable dt;
        //string conStringTrayGuardDb = @"Server=172.27.40.17;Port=5432;User Id=pqm;Password=dbuser;Database=barcodeprint_kk04; CommandTimeout=100; Timeout=100;";
        string conStringTrayGuardDb = string.Empty;
        string appconfig = System.Environment.CurrentDirectory + @"\info.ini"; // 設定ファイルのアドレス

        // コンストラクタ
        public frmMasterLabel()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void frmMasterItems_Load(object sender, EventArgs e)
        {
            //フォームの場所を指定
            this.Left = 450;
            this.Top = 100;

            //ＩＰアドレスの読み込み
            //conStringTrayGuardDb = @"Server=" + readIni("IP ADDRESS", "TRAYGUARD DB", appconfig) + @";Port=5432;User Id=pqm;Password=dbuser;Database=barcodeprint; CommandTimeout=100; Timeout=100;";
            conStringTrayGuardDb = @"Server=" + readIni("IP ADDRESS", "TRAYGUARD DB", appconfig) + @";Port=5432;User Id=pqm;Password=dbuser;Database="+ readIni("DATABASE NAME", "BARCODEPRINT DBNAME", appconfig)+"; CommandTimeout=100; Timeout=100;";

            string sql = "select model, header, content from t_label_content order by model, header";
            System.Diagnostics.Debug.Print(sql);
            defineAndReadTable(sql);
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

        // サブプロシージャ：テーブルを定義し、ＤＢよりデータを読み込む
        private void defineAndReadTable(string sql)
        {
            // ＤＢよりデータを読み込み、ＤＴＡＡＴＡＢＬＥへ格納
            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            command = new NpgsqlCommand(sql, connection);
            adapter = new NpgsqlDataAdapter(command);
            cmdbuilder = new NpgsqlCommandBuilder(adapter);
            adapter.InsertCommand = cmdbuilder.GetInsertCommand();
            adapter.UpdateCommand = cmdbuilder.GetUpdateCommand();
            adapter.DeleteCommand = cmdbuilder.GetDeleteCommand();
            ds = new DataSet();
            adapter.Fill(ds,"buff");
            dt = ds.Tables["buff"];
            
            // データグリットビューへＤＴＡＡＴＡＢＬＥを格納
            dgvTester.DataSource = dt;
            dgvTester.ReadOnly = true;
            btnSave.Enabled = false;
            dgvTester.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvTester.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        // 新規レコードの追加
        private void btnAdd_Click(object sender, EventArgs e)
        {
            dgvTester.ReadOnly = false;
            dgvTester.AllowUserToAddRows = true;
            btnSave.Enabled = true;
            btnAdd.Enabled = false;
            btnDelete.Enabled = false;
        }

        // 既存レコードの削除
        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult dlg = MessageBox.Show("Do you want to delete this row ?", "Delete", MessageBoxButtons.YesNo);
            if (dlg == DialogResult.No) return;

            try
            {
                dgvTester.Rows.RemoveAt(dgvTester.SelectedRows[0].Index);
                adapter.Update(dt);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 保存
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                adapter.Update(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally 
            {
                dgvTester.ReadOnly = true;
                dgvTester.AllowUserToAddRows = false;
                btnSave.Enabled = false;
                btnAdd.Enabled = true;
                btnDelete.Enabled = true;
                //親フォーム更新のため、デレゲートイベントを発生させる
                //this.RefreshEvent(this, new EventArgs());
            }
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

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}