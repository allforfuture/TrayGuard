using System;
using System.Windows.Forms;
using System.Security.Permissions;

namespace TrayGuard
{
    public partial class frmCapacity : Form
    {
        //親フォームfrmTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        int capacity;

        // コンストラクタ
        public frmCapacity()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void Form4_Load(object sender, EventArgs e)
        {
            //フォームの場所を指定
            this.Left = 450;
            this.Top = 100;
        }

        // 親フォームで呼び出し、親フォームの情報を、テキストボックスへ格納して引き継ぐ
        public void updateControls(int cap)
        {
            capacity = cap;
            txtNewCapacity.Text = cap.ToString();
        }

        // 親フォームで呼び出し、キャパシティーを返す
        public int returnCapacity()
        {
            return capacity;
        }

        // frmModuleInTray ラベルあたりのシリアル数（capacity）を変更
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (int.TryParse(txtNewCapacity.Text, out capacity) && capacity > 0)
            {
                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
                Close();
            }
        }

        // frmModuleInTray ラベルあたりのシリアル数（capacity）を変更
        private void txtNewCapacity_KeyDown(object sender, KeyEventArgs e)
        {
            // エンターキーの場合のみ、処理を行う
            if (e.KeyCode != Keys.Enter) return;
            // ＯＫボタンの押下と同じ処理
            btnOK_Click(sender, e);
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