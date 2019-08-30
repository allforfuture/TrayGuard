using System;
using System.Windows.Forms;
using System.Security.Permissions;
using System.Linq;

namespace TrayGuard
{
    public partial class frmModuleFind : Form
    {
        //親フォームfrmTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        // コンストラクタ
        public frmModuleFind()
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

        // 親フォームで呼び出し、キャパシティーを返す
        public string returnTargetModule()
        {
            return txtModuleId.Text;
        }

        // frmModuleInTray ラベルあたりのシリアル数（capacity）を変更
        private void btnOK_Click(object sender, EventArgs e)
        {
            //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
            if (txtModuleId.Text == string.Empty) return;
            this.RefreshEvent(this, new EventArgs());
            Close();
        }

        // frmModuleInTray ラベルあたりのシリアル数（capacity）を変更
        private void txtNewCapacity_KeyDown(object sender, KeyEventArgs e)
        {
            // エンターキーの場合のみ、処理を行う
            if (e.KeyCode != Keys.Enter) return;
            // ＯＫボタンの押下と同じ処理
            btnOK_Click(sender, e);
        }
    }
}