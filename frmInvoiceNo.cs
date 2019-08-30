using System;
using System.Windows.Forms;
using System.Security.Permissions;

namespace TrayGuard
{
    public partial class frmInvoiceNo : Form
    {
        //親フォームfrmTrayへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        string[] cartonlist;
        string message;


        // コンストラクタ
        public frmInvoiceNo()
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
        public void updateControls(string[] carton, string msg, string invoice)
        {
            cartonlist = carton;
            message = msg;
            txtInvoiceNo.Text = invoice;
        }

        // frmPallet インボイス番号を登録
        private void btnOK_Click(object sender, EventArgs e)
        {
            string invoice = txtInvoiceNo.Text;

            if (invoice == string.Empty)
            {
                DialogResult result = MessageBox.Show("Selected cells' invoice no is to be updated." + Environment.NewLine +
                    "Is it OK?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.No) return;
            }

            TfSQL tf = new TfSQL();
            bool res = tf.sqlMultipleUpdateInvoiceOnPallet(cartonlist, invoice);

            if (res)
            {
                //本フォームのデータグリットビュー更新
                MessageBox.Show("The following carton IDs' invoice number were updated: " + Environment.NewLine + message, "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                //親フォームfrmTrayのデータグリットビューを更新するため、デレゲートイベントを発生させる
                this.RefreshEvent(this, new EventArgs());
            }
            else
            {
                MessageBox.Show("Invoice number registration was not successful.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            Close();
        }

        // frmModuleInTray ラベルあたりのシリアル数（capacity）を変更
        private void txtNewCapacity_KeyDown(object sender, KeyEventArgs e)
        {
            // エンターキーの場合、テキストボックスの桁数が１９桁の場合のみ、処理を行う
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