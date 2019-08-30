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
        //�e�t�H�[��frmModuleInTray�փC�x���g������A���i�f���Q�[�g�j
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        //���̑��񃍁[�J���ϐ�
        NpgsqlConnection connection;
        NpgsqlCommand command;
        NpgsqlDataAdapter adapter;
        NpgsqlCommandBuilder cmdbuilder;
        DataSet ds;
        DataTable dt;
        //string conStringTrayGuardDb = @"Server=172.27.40.17;Port=5432;User Id=pqm;Password=dbuser;Database=barcodeprint_kk04; CommandTimeout=100; Timeout=100;";
        string conStringTrayGuardDb = string.Empty;
        string appconfig = System.Environment.CurrentDirectory + @"\info.ini"; // �ݒ�t�@�C���̃A�h���X

        // �R���X�g���N�^
        public frmMasterLabel()
        {
            InitializeComponent();
        }

        // ���[�h���̏���
        private void frmMasterItems_Load(object sender, EventArgs e)
        {
            //�t�H�[���̏ꏊ���w��
            this.Left = 450;
            this.Top = 100;

            //�h�o�A�h���X�̓ǂݍ���
            //conStringTrayGuardDb = @"Server=" + readIni("IP ADDRESS", "TRAYGUARD DB", appconfig) + @";Port=5432;User Id=pqm;Password=dbuser;Database=barcodeprint; CommandTimeout=100; Timeout=100;";
            conStringTrayGuardDb = @"Server=" + readIni("IP ADDRESS", "TRAYGUARD DB", appconfig) + @";Port=5432;User Id=pqm;Password=dbuser;Database="+ readIni("DATABASE NAME", "BARCODEPRINT DBNAME", appconfig)+"; CommandTimeout=100; Timeout=100;";

            string sql = "select model, header, content from t_label_content order by model, header";
            System.Diagnostics.Debug.Print(sql);
            defineAndReadTable(sql);
        }

        // �ݒ�e�L�X�g�t�@�C���̓ǂݍ���
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
        // Windows API ���C���|�[�g
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filepath);

        // �T�u�v���V�[�W���F�e�[�u�����`���A�c�a���f�[�^��ǂݍ���
        private void defineAndReadTable(string sql)
        {
            // �c�a���f�[�^��ǂݍ��݁A�c�s�`�`�s�`�a�k�d�֊i�[
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
            
            // �f�[�^�O���b�g�r���[�ւc�s�`�`�s�`�a�k�d���i�[
            dgvTester.DataSource = dt;
            dgvTester.ReadOnly = true;
            btnSave.Enabled = false;
            dgvTester.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvTester.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        // �V�K���R�[�h�̒ǉ�
        private void btnAdd_Click(object sender, EventArgs e)
        {
            dgvTester.ReadOnly = false;
            dgvTester.AllowUserToAddRows = true;
            btnSave.Enabled = true;
            btnAdd.Enabled = false;
            btnDelete.Enabled = false;
        }

        // �������R�[�h�̍폜
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

        // �ۑ�
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
                //�e�t�H�[���X�V�̂��߁A�f���Q�[�g�C�x���g�𔭐�������
                //this.RefreshEvent(this, new EventArgs());
            }
        }

        // ����{�^����V���[�g�J�b�g�ł̏I���������Ȃ�
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