using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Npgsql;
using System.Text;
using System.Runtime.InteropServices;
using System.Globalization;

namespace TrayGuard
{
    public class TfSQL
    {
        NpgsqlConnection connection;
        string conStringTrayGuardDb;
        string appconfig = System.Environment.CurrentDirectory + @"\info.ini"; // 設定ファイルのアドレス
        string model;

        // コンストラクタ
        public TfSQL()
        {
            conStringTrayGuardDb = @"Server=" + readIni("IP ADDRESS", "TRAYGUARD DB", appconfig) + 
                @";Port=5432;User Id=pqm;Password=dbuser;Database="+ readIni("DATABASE NAME", "BARCODEPRINT DBNAME", appconfig)+"; CommandTimeout=100; Timeout=100;";
            model = readIni("MODEL", "MODEL NAME", appconfig);
        }

        // 設定テキストファイルの読み込み
        public string readIni(string s, string k, string cfs)
        {
            StringBuilder retVal = new StringBuilder(255);
            string section = s;
            string key = k;
            string def = String.Empty;
            int size = 255;
            int strref = GetPrivateProfileString(section, key, def, retVal, size, cfs);
            return retVal.ToString();
        }
        public static string readIni_static(string s, string k, string cfs)
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

        // コンボボックスのリストを挿入する（トレイガードＤＢ）
        public void getComboBoxData(string sql, ref ComboBox cmb)
        {
            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter();
            NpgsqlCommand command;
            DataSet ds = new DataSet();
            try
            {
                connection = new NpgsqlConnection(conStringTrayGuardDb);
                connection.Open();
                command = new NpgsqlCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                cmb.Items.Clear();
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    cmb.Items.Add(row[0].ToString());
                }
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
            }
        }

        // ＤＡＴＥＴＩＭＥスカラを返す（トレイガードＤＢ）
        public DateTime sqlExecuteScalarDateTime(string sql)
        {
            DateTime response;
            try
            {
                connection = new NpgsqlConnection(conStringTrayGuardDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                response = (DateTime)command.ExecuteScalar();
                connection.Close();
                return response;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return DateTime.Now;
            }
        }

        // 文字列スカラを返す（トレイガードＤＢ）
        public string sqlExecuteScalarString(string sql)
        {
            string response;
            try
            {
                connection = new NpgsqlConnection(conStringTrayGuardDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                response = Convert.ToString(command.ExecuteScalar());
                connection.Close();
                return response;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return String.Empty;
            }
        }


        // 非クエリＳＱＬの実行（トレイガードＤＢ）
        public bool sqlExecuteNonQuery(string sql, bool result_message_show)
        {
            try
            {
                connection = new NpgsqlConnection(conStringTrayGuardDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                int response = command.ExecuteNonQuery();
                if (response >= 1)
                {
                    if (result_message_show) { MessageBox.Show("Successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    connection.Close();
                    return true;
                }
                else
                {
                    //MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // 対象レコードをＤＡＴＡＴＡＢＬＥへ格納する（トレイガードＤＢ）
        public void sqlDataAdapterFillDatatableFromTrayGuardDb(string sql, ref DataTable dt)
        {
            NpgsqlConnection connection = new NpgsqlConnection(conStringTrayGuardDb);
            NpgsqlCommand command = new NpgsqlCommand();

            using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter())
            {
                command.CommandText = sql;
                command.Connection = connection;
                adapter.SelectCommand = command;
                adapter.Fill(dt);
            }
        }


        // モジュールの重複チェック
        public string sqlModuleDuplicateCheck(DataTable dt)
        {
            //int res1;
            string currentModule = string.Empty;
            string mdlrmode = string.Empty;
            // 2016.08.22 FUJIKI 17桁チェックへの仕様変更
            // connection = new NpgsqlConnection(conStringTrayGuardDb);
            // connection.Open();
            // NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                // 2016.08.22 FUJIKI 17桁チェックへの仕様変更
                // string sql = "insert into t_module (module_id, lot, tester_id, test_result, test_date, r_mode, tray_id) " +
                //     "VALUES (:module_id, 'zzz', 'zzz', 'zzz', '2000/1/1', :r_mode, 'zzz')";
                // NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                // 
                // command.Parameters.Add(new NpgsqlParameter("module_id", NpgsqlTypes.NpgsqlDbType.Text));
                // command.Parameters.Add(new NpgsqlParameter("r_mode", NpgsqlTypes.NpgsqlDbType.Text));
                // 
                // for (int i = 0; i < dt.Rows.Count; i++)
                // {
                // currentModule = dt.Rows[i]["module_id"].ToString();
                // command.Parameters[0].Value = dt.Rows[i]["module_id"].ToString();
                // command.Parameters[1].Value = dt.Rows[i]["r_mode"].ToString();
                // res1 = command.ExecuteNonQuery();
                // }

                // 2016.08.22 FUJIKI 17桁チェックへの仕様変更
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    currentModule = dt.Rows[i]["module_id"].ToString();
                    // string sql = "select r_mode from t_module where module_id like '" + VBS.Left(currentModule, 17) + "%'";
                    string sql = "select r_mode from t_module where module_id like '" + VBS.Left(currentModule, 17) + "%' and r_mode = '" + dt.Rows[i]["r_mode"].ToString() + "'";
                    mdlrmode = sqlExecuteScalarString(sql);
                    if (!string.IsNullOrEmpty(mdlrmode))
                    {
                        MessageBox.Show(currentModule + "[" + dt.Rows[i]["r_mode"].ToString()  + "] already exists in database." + Environment.NewLine, "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        connection.Close();
                        return dt.Rows[i]["module_id"].ToString();
                    }
                }
                connection.Close();
                return string.Empty;
            }
            catch (Exception ex)
            {
                // 2016.08.22 FUJIKI 17桁チェックへの仕様変更
                // transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return currentModule;
            }
        }

        //check this sn is beed shopping
        public string sqlModuleDuplicateCheck(DataRow dr)
        {
            //int res1;
            string currentModule = string.Empty;
            string mdlrmode = string.Empty;
            try
            {
                currentModule = dr["module_id"].ToString();
                // string sql = "select r_mode from t_module where module_id like '" + VBS.Left(currentModule, 17) + "%'";
                string sql = "select r_mode from t_module where module_id like '" + VBS.Left(currentModule, 17) + "%' and r_mode = '" + dr["r_mode"].ToString() + "'";
                mdlrmode = sqlExecuteScalarString(sql);
                if (!string.IsNullOrEmpty(mdlrmode))
                {
                    //MessageBox.Show(currentModule + "[" + dr["r_mode"].ToString() + "] already exists in database." + Environment.NewLine, "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    return dr["module_id"].ToString();
                }
                connection.Close();
                return string.Empty;
            }
            catch (Exception ex)
            {
                // 2016.08.22 FUJIKI 17桁チェックへの仕様変更
                // transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return currentModule;
            }
        }

        // トレーＩＤ・モジュールの一括登録
        public bool sqlMultipleInsertModule(DataTable dt, string trayId)
        {
            int res1;
            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                string sql = "insert into t_module (module_id, lot, bin, tester_id, test_result, test_date, r_mode, tray_id) " +
                    "VALUES (:module_id, :lot, :bin, :tester_id, :test_result, :test_date, :r_mode, :tray_id)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("module_id", NpgsqlTypes.NpgsqlDbType.Text));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Text));
                command.Parameters.Add(new NpgsqlParameter("bin", NpgsqlTypes.NpgsqlDbType.Text));
                command.Parameters.Add(new NpgsqlParameter("tester_id", NpgsqlTypes.NpgsqlDbType.Text));
                command.Parameters.Add(new NpgsqlParameter("test_result", NpgsqlTypes.NpgsqlDbType.Text));
                command.Parameters.Add(new NpgsqlParameter("test_date", NpgsqlTypes.NpgsqlDbType.Timestamp));
                command.Parameters.Add(new NpgsqlParameter("r_mode", NpgsqlTypes.NpgsqlDbType.Text));
                command.Parameters.Add(new NpgsqlParameter("tray_id", NpgsqlTypes.NpgsqlDbType.Text));
                NpgsqlDataAdapter da = new NpgsqlDataAdapter();
                
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["module_id"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["lot"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["bin"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["tester_id"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["test_result"].ToString();
                    command.Parameters[5].Value = (DateTime)dt.Rows[i]["test_date"];
                    command.Parameters[6].Value = dt.Rows[i]["r_mode"].ToString();
                    command.Parameters[7].Value = trayId;

                    res1 = command.ExecuteNonQuery();
                }

                transaction.Commit();
                connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // トレーＩＤ・モジュールの一括キャンセル
        public bool sqlMultipleCancelModuleInTray(string[] cellVal, string dept, string user)
        {
            string cancelTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            int res1;
            bool res2 = false;
            int res3;
            bool res4 = false;

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted); 

            try
            {
                string sql1 = "update t_tray set cancel_date = '" + cancelTime + "', " +
                    "cl_dept = '" + dept + "', " + "cl_user = '" + user + "' where tray_id = :tray_id";
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                System.Diagnostics.Debug.Print(sql1);
                command1.Parameters.Add(new NpgsqlParameter("tray_id", NpgsqlTypes.NpgsqlDbType.Text));

                for (int i = 0; i < cellVal.Length; i++)
                {
                    command1.Parameters[0].Value = cellVal[i];
                    res1 = command1.ExecuteNonQuery();
                    if (res1 <= 0) res2 = true;
                }

                if (res2)
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    return false;
                }

                string sql2 = "delete from t_module where tray_id = :tray_id";

                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                command2.Parameters.Add(new NpgsqlParameter("tray_id", NpgsqlTypes.NpgsqlDbType.Text));

                for (int i = 0; i < cellVal.Length; i++)
                {
                    command2.Parameters[0].Value = cellVal[i];
                    res3 = command2.ExecuteNonQuery();
                    if (res3 <= -1) res4 = true;
                }

                if (res4)
                {
                    throw new InvalidOperationException(); // 2016.08.22 FUJII 可読性を高めるため、キャッチへ後続処理を移動
                }
                
                //正常に処理が完了した場合
                transaction.Commit();
                connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // トレーＩＤ・モジュールのキャンセル
        public bool sqlCancelModuleInTray(string trayId, string dept, string user)
        {
            int res1;
            int res2;
            string sql1 = "update t_tray set cancel_date = '" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "', " +
                "cl_dept = '" + dept + "', " + "cl_user = '" + user + "' where tray_id ='" + trayId + "'";
            string sql2 = "delete from t_module where tray_id ='" + trayId + "'";

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command = new NpgsqlCommand(sql1, connection);
                res1 = command.ExecuteNonQuery();

                command = new NpgsqlCommand(sql2, connection);
                res2 = command.ExecuteNonQuery();

                if (res1 >= 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    throw new InvalidOperationException(); // 2016.08.22 FUJII 可読性を高めるため、キャッチへ後続処理を移動
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // トレーＩＤ・モジュールの更新（トレー情報は更新し、モジュールレコードは一旦キャンセルした後、一括登録）
        public bool sqlUpdateModuleInPack(string trayId, string sql1)
        {
            int res1;
            int res2;
            string sql2 = "delete from t_module where tray_id ='" + trayId + "'";

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command = new NpgsqlCommand(sql1, connection);
                res1 = command.ExecuteNonQuery();

                command = new NpgsqlCommand(sql2, connection);
                res2 = command.ExecuteNonQuery();

                if (res1 >= 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    throw new InvalidOperationException(); // 2016.08.22 FUJII 可読性を高めるため、キャッチへ後続処理を移動
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // トレーテーブルのパックＩＤフィールド、一括更新
        public bool sqlMultipleUpdateTrayInPack(DataTable dt, string pack)
        {
            string sql1 = "update t_tray set pack_id = '" + pack + "' where tray_id in (";
            string sql2 = string.Empty;
            var query = dt.AsEnumerable()
                        .Select(row => new { tray_id = row.Field<string>("tray_id") });
            foreach (var q in query) sql2 += "'" + q.tray_id + "', ";

            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
            System.Diagnostics.Debug.Print(sql3);

            try
            {
                connection = new NpgsqlConnection(conStringTrayGuardDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql3, connection);
                command.ExecuteNonQuery();
                connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // パックＩＤ・トレーのキャンセル
        public bool sqlCancelTrayInPack(DataTable dt, string pack, string user)
        {
            int res1;
            int res2;
            string sql0 = "update t_pack set cancel_date = '" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "', " +
                "cl_user = '" + user + "' where pack_id ='" + pack + "'";

            string sql1 = "update t_tray set pack_id = null where tray_id in (";
            string sql2 = string.Empty;
            var query = dt.AsEnumerable()
                        .Select(row => new { tray_id = row.Field<string>("tray_id") });
            foreach (var q in query) sql2 += "'" + q.tray_id + "', ";
            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
            System.Diagnostics.Debug.Print(sql3);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command = new NpgsqlCommand(sql0, connection);
                res1 = command.ExecuteNonQuery();

                command = new NpgsqlCommand(sql3, connection);
                res2 = command.ExecuteNonQuery();

                if (res1 >= 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    throw new InvalidOperationException(); // 2016.08.22 FUJII 可読性を高めるため、キャッチへ後続処理を移動
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // パックテーブルのカートンＩＤフィールド、一括更新
        public bool sqlMultipleUpdatePackInCarton(DataTable dt, string carton)
        {
            string sql1 = "update t_pack set carton_id = '" + carton + "' where pack_id in (";
            string sql2 = string.Empty;
            var query = dt.AsEnumerable()
                        .Select(row => new { pack_id = row.Field<string>("pack_id") });
            foreach (var q in query) sql2 += "'" + q.pack_id + "', ";

            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
            System.Diagnostics.Debug.Print(sql3);

            try
            {
                connection = new NpgsqlConnection(conStringTrayGuardDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql3, connection);
                command.ExecuteNonQuery();
                connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // カートンＩＤ・パックのキャンセル
        public bool sqlCancelPackInCarton(DataTable dt, string carton, string user)
        {
            int res1;
            int res2;
            string sql0 = "update t_carton set cancel_date = '" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "', " +
                "cl_user = '" + user + "' where carton_id ='" + carton + "'";

            string sql1 = "update t_pack set carton_id = null where pack_id in (";
            string sql2 = string.Empty;
            var query = dt.AsEnumerable()
                        .Select(row => new { pack_id = row.Field<string>("pack_id") });
            foreach (var q in query) sql2 += "'" + q.pack_id + "', ";
            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
            System.Diagnostics.Debug.Print(sql3);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command = new NpgsqlCommand(sql0, connection);
                res1 = command.ExecuteNonQuery();

                command = new NpgsqlCommand(sql3, connection);
                res2 = command.ExecuteNonQuery();

                if (res1 >= 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    throw new InvalidOperationException(); // 2016.08.22 FUJII 可読性を高めるため、キャッチへ後続処理を移動
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // カートンテーブルのインボイスＮＯフィールド、一括更新
        public bool sqlMultipleUpdateInvoiceOnPallet(string[] carton, string invoice)
        {
            string sql1 = "update t_pallet set invoice_no = '" + invoice + "' where pallet_id in (";
            string sql2 = string.Empty;
            for (int i = 0; i < carton.Length; i ++) sql2 += "'" + carton[i] + "', ";
            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
            System.Diagnostics.Debug.Print(sql3);

            try
            {
                connection = new NpgsqlConnection(conStringTrayGuardDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql3, connection);
                command.ExecuteNonQuery();
                connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // カートンテーブルのパレットＩＤフィールド、一括更新
        public bool sqlMultipleUpdateCartonOnPallet(DataTable dt, string pallet)
        {
            string sql1 = "update t_carton set pallet_id = '" + pallet + "' where carton_id in (";
            string sql2 = string.Empty;
            var query = dt.AsEnumerable().Select(row => new { carton_id = row.Field<string>("carton_id") });
            foreach (var q in query) sql2 += "'" + q.carton_id + "', ";
            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
            System.Diagnostics.Debug.Print(sql3);

            try
            {
                connection = new NpgsqlConnection(conStringTrayGuardDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql3, connection);
                command.ExecuteNonQuery();
                connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // パレットＩＤ・カートンのキャンセル
        public bool sqlCancelCartonOnPallet(DataTable dt, string pallet, string user)
        {
            int res1;
            int res2;
            string sql0 = "update t_pallet set cancel_date = '" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "', " +
                "cl_user = '" + user + "' where pallet_id ='" + pallet + "'";

            string sql1 = "update t_carton set pallet_id = null where carton_id in (";
            string sql2 = string.Empty;
            var query = dt.AsEnumerable().Select(row => new { carton_id = row.Field<string>("carton_id") });
            foreach (var q in query) sql2 += "'" + q.carton_id + "', ";
            string sql3 = sql1 + VBS.Left(sql2, sql2.Length - 2) + ")";
            System.Diagnostics.Debug.Print(sql3);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command = new NpgsqlCommand(sql0, connection);
                res1 = command.ExecuteNonQuery();

                command = new NpgsqlCommand(sql3, connection);
                res2 = command.ExecuteNonQuery();

                if (res1 >= 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    throw new InvalidOperationException(); // 2016.08.22 FUJII 可読性を高めるため、キャッチへ後続処理を移動
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // 2016.08.29 FUJII  
        // 登録済パレットに対する、カートンの置き換え
        public bool sqlReplaceCartonOnPallet(string pallet, string before, string after, string user)
        {
            int res1;
            int res2;
            string sql1 = "update t_carton set pallet_id = null where carton_id = '" + before + "'";
            string sql2 = "update t_carton set pallet_id = '" + pallet + "' where carton_id = '" + after + "'";
            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command = new NpgsqlCommand(sql1, connection);
                res1 = command.ExecuteNonQuery();

                command = new NpgsqlCommand(sql2, connection);
                res2 = command.ExecuteNonQuery();

                if (res1 ==1 && res2 ==1)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    throw new InvalidOperationException(); // 2016.08.22 FUJII 可読性を高めるため、キャッチへ後続処理を移動
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // 2016.08.10 FUJII  トレーＩＤの新規採番プロシージャ「GetNewTrayId」を、トランザクション処理バージョンへ変更
        // 2016.08.26 FUJII  トレーＩＤの１５～１７桁シリアルの先頭１桁を「１～Ｚ」とし、シリアル上限を増加
        // frmModuleInTrayにおける、トレーＩＤの採番
        public string sqlGetNewTrayId(string udept, string uname, string line, string shift, string maxLot, DataTable dtLot, ref DateTime registerDate)
        {
            // 含まれるモジュールが最も多いロットを元に、同ロットの最も新しいＤＢ登録済みトレーＩＤを取得する（旧トレーＩＤ）
            string lot = dtLot.Columns[0].ColumnName;
            string multiLot = dtLot.Columns.Count >= 3 ? "T" : "F";
            int qty = (int)dtLot.Rows[0]["total"];
            string lotLine = VBS.Mid(maxLot, 8, 1);
            int year = int.Parse("201" + VBS.Mid(maxLot, 4, 1));
            int week = int.Parse(VBS.Mid(maxLot, 5, 2));
            int day = int.Parse(VBS.Mid(maxLot, 7, 1));
            DateTime lotDate = getFirstDateOfWeek(year, week, CultureInfo.CurrentCulture).AddDays(day - 1); //
            string sql0 = "LOCK TABLE t_tray IN ACCESS EXCLUSIVE MODE";   // 重複トレーＩＤ防止のため、テーブル t_tray の読み取りロック
            string sql1 = "select MAX(tray_id) from t_tray where lot_date ='" + lotDate.ToString("yyMMdd") + "' and " +
                "rg_dept ='" + udept + "' and " + "line ='" + lotLine + "'";
            System.Diagnostics.Debug.Print(sql1);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                // 重複トレーＩＤ防止のため、テーブル t_tray の読み取りロック
                NpgsqlCommand command = new NpgsqlCommand(sql0, connection);
                command.ExecuteNonQuery();

                // 旧トレーＩＤを取得
                command = new NpgsqlCommand(sql1, connection);
                string trayOld = command.ExecuteScalar().ToString();

                // 新トレーＩＤを採番する
                string trayNew;
                if (trayOld != string.Empty)
                {
                    int serialOld = convertToIntegerSerial(VBS.Mid(trayOld, 15, 3));
                    string serialNew = convertToCharSerial(serialOld + 1);
                    trayNew = VBS.Left(udept, 1) + "_"+ model+"_" + lotDate.ToString("yyMMdd") + "_" + lotLine + "_" + serialNew + "_" + shift;
                }
                else
                {
                    trayNew = VBS.Left(udept, 1) + "_" + model + "_" + lotDate.ToString("yyMMdd") + "_" + lotLine + "_001_" + shift;
                }
                System.Diagnostics.Debug.Print(trayNew);

                // 新トレーＩＤの登録
                registerDate = DateTime.Now;
                string sql2 = "insert into t_tray (tray_id, lot, lot_date, line, shift, qty, register_date, rg_dept, rg_user, multi_lot) " +
                    "VALUES('" + trayNew + "','" + lot + "','" + lotDate.ToString("yyMMdd") + "','" + lotLine + "','" + shift + "','" +
                         qty + "','" + registerDate.ToString("yyyy/MM/dd HH:mm:ss") + "','" + udept + "','" + uname + "','" + multiLot + "')";
                System.Diagnostics.Debug.Print(sql2);

                command = new NpgsqlCommand(sql2, connection);
                int res = command.ExecuteNonQuery();

                // 正常な処理が行われた場合はコミット、それ以外は例外を投げ、例外処理内でロールバックする
                if (res == 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return trayNew;
                }
                else
                {
                    throw new InvalidOperationException();
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return string.Empty;
            }
        }

        // 整数２桁とアルファベットをマッピングするための配列
        string[] oneDigitMap = new string[] { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        string[] twoDigitMap = new string[] { "00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35" };

        // トレーＩＤシリアル３桁、左一桁を整数へ変換
        private int convertToIntegerSerial(string charSerial)
        {
            string firstBefore = VBS.Left(charSerial, 1);
            string firstAfter = twoDigitMap[Array.IndexOf(oneDigitMap, firstBefore)];
            return int.Parse(firstAfter + VBS.Mid(charSerial, 2, 2));
        }

        // トレーＩＤシリアル３桁、左一桁をアルファベットへ変換
        private string convertToCharSerial(int numberSerial)
        {
            string firstBefore = VBS.Left(numberSerial.ToString("0000"), 2);
            string firstAfter = oneDigitMap[Array.IndexOf(twoDigitMap, firstBefore)];
            return firstAfter + VBS.Mid(numberSerial.ToString("0000"), 3, 2);
        }

        // 2016.08.22 FUJII  パックＩＤの新規採番プロシージャ「getNewPackId」を、トランザクション処理バージョンへ変更
        // frmTrayInPackにおける、パックＩＤの採番
        public string sqlGetNewPackId(string lot, string batch, string user, DataTable dtLot, ref DateTime registerDate)
        {
            int lotCnt = dtLot.Columns.Count - 1;
            int qty = (int)dtLot.Rows[0]["total"];
            registerDate = DateTime.Now;

            string sql0 = "LOCK TABLE t_pack IN ACCESS EXCLUSIVE MODE";
            string sql1 = "select MAX(pack_id) from t_pack where register_date >='" + registerDate.Date.ToString("yyyy/MM/dd") + "'";
            System.Diagnostics.Debug.Print(sql1);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                // 重複パックＩＤ防止のため、テーブル t_pack の読み取りロック
                NpgsqlCommand command = new NpgsqlCommand(sql0, connection);
                command.ExecuteNonQuery();

                // 旧パックＩＤを取得
                command = new NpgsqlCommand(sql1, connection);
                string packOld = command.ExecuteScalar().ToString();

                // 旧パックＩＤが存在しない場合、または、
                // 旧パックＩＤが存在し、その日付が本日より前の日付の場合は、本日日付０００１番のＩＤを発行
                // それ以外の場合は、旧パックＩＤ＋１で、ＩＤを発行を発行
                //DateTime packOldDate = DateTime.ParseExact("20" + VBS.Mid(packOld, 4, 6), "yyyyMMdd", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
                //bool packOldDateSmallerThanToday = packOldDate < DateTime.Today;

                int numberOld = 0;
                string packNew;
                //if (packOld != string.Empty && !packOldDateSmallerThanToday)
                if (packOld != string.Empty)
                {
                    numberOld = int.Parse(VBS.Mid(packOld, 12, 4));
                    packNew = VBS.Left(packOld, 11) + (numberOld + 1).ToString("0000");
                }
                else
                {
                    packNew = model+"_" + registerDate.Date.ToString("yyMMdd") + "_P" + "0001";
                }
                System.Diagnostics.Debug.Print(packNew);

                // 新パックＩＤの登録
                string sql2 = "insert into t_pack (pack_id, lot, l_cnt, m_qty, batch, register_date, rg_user) " +
                    "VALUES('" + packNew + "','" + lot + "','" + lotCnt + "','" + qty + "','" + batch + "','" +
                         registerDate.ToString("yyyy/MM/dd HH:mm:ss") + "','" + user + "')";
                System.Diagnostics.Debug.Print(sql2);

                command = new NpgsqlCommand(sql2, connection);
                int res = command.ExecuteNonQuery();

                // 正常な処理が行われた場合はコミット、それ以外は例外を投げ、例外処理内でロールバックする
                if (res == 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return packNew;
                }
                else
                {
                    throw new InvalidOperationException();
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return string.Empty;
            }
        }

        // 2016.08.22 FUJII  カートンＩＤの新規採番プロシージャ「getNewCartonId」を、トランザクション処理バージョンへ変更
        // frmPackInCartonにおける、カートンＩＤの採番
        public string sqlGetNewCartonId(string lot, string batch, string user, DataTable dtLot, ref DateTime registerDate)
        {
            int lotCnt = dtLot.Columns.Count - 1;
            int qty = (int)dtLot.Rows[0]["total"];
            registerDate = DateTime.Now;

            string sql0 = "LOCK TABLE t_carton IN ACCESS EXCLUSIVE MODE";
            string sql1 = "select MAX(carton_id) from t_carton where register_date >='" + registerDate.Date.ToString("yyyy/MM/dd") + "'";
            System.Diagnostics.Debug.Print(sql1);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                // 重複カートンＩＤ防止のため、テーブル t_pack の読み取りロック
                NpgsqlCommand command = new NpgsqlCommand(sql0, connection);
                command.ExecuteNonQuery();

                // 旧カートンＩＤを取得
                command = new NpgsqlCommand(sql1, connection);
                string cartonOld = command.ExecuteScalar().ToString();

                // 旧カートンＩＤが存在しない場合、または、
                // 旧カートンＩＤが存在し、その日付が本日より前の日付の場合は、本日日付０００１番のＩＤを発行
                // それ以外の場合は、旧カートンＩＤ＋１で、ＩＤを発行を発行
                //DateTime cartonOldDate = DateTime.ParseExact("20" + VBS.Mid(cartonOld, 4, 6), "yyyyMMdd", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
                //bool cartonOldDateSmallerThanToday = cartonOldDate < DateTime.Today;

                int numberOld = 0;
                string cartonNew;
                //if (cartonOld != string.Empty && !cartonOldDateSmallerThanToday)
                if (cartonOld != string.Empty)
                {
                    numberOld = int.Parse(VBS.Mid(cartonOld, 12, 4));
                    cartonNew = VBS.Left(cartonOld, 11) + (numberOld + 1).ToString("0000");
                }
                else
                {
                    cartonNew = model+"_" + registerDate.Date.ToString("yyMMdd") + "_C" + "0001";
                }
                System.Diagnostics.Debug.Print(cartonNew);

                // 新カートンＩＤの登録
                string sql2 = "insert into t_carton (carton_id, lot, l_cnt, m_qty, batch, register_date, rg_user) " +
                    "VALUES('" + cartonNew + "','" + lot + "','" + lotCnt + "','" + qty + "','" + batch + "','" +
                    registerDate.ToString("yyyy/MM/dd HH:mm:ss") + "','" + user + "')";
                System.Diagnostics.Debug.Print(sql2);

                command = new NpgsqlCommand(sql2, connection);
                int res = command.ExecuteNonQuery();

                // 正常な処理が行われた場合はコミット、それ以外は例外を投げ、例外処理内でロールバックする
                if (res == 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return cartonNew;
                }
                else
                {
                    throw new InvalidOperationException();
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return string.Empty;
            }
        }

        // 2016.08.22 FUJII  パレットＩＤの新規採番プロシージャ「getNewPalletId」を、トランザクション処理バージョンへ変更
        // frmCartonOnPalletにおける、パレットＩＤの採番
        public string sqlGetNewPalletId(string lot, string batch, string user, DataTable dtLot, ref DateTime registerDate)
        {
            int lotCnt = dtLot.Columns.Count - 1;
            int qty = (int)dtLot.Rows[0]["total"];
            registerDate = DateTime.Now;

            string sql0 = "LOCK TABLE t_pallet IN ACCESS EXCLUSIVE MODE";
            string sql1 = "select MAX(pallet_id) from t_pallet where register_date >='" + registerDate.Date.ToString("yyyy/MM/dd") + "'";
            System.Diagnostics.Debug.Print(sql1);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                // 重複パレットＩＤ防止のため、テーブル t_pack の読み取りロック
                NpgsqlCommand command = new NpgsqlCommand(sql0, connection);
                command.ExecuteNonQuery();

                // 旧パレットＩＤを取得
                command = new NpgsqlCommand(sql1, connection);
                string palletOld = command.ExecuteScalar().ToString();

                // 旧パレットＩＤが存在しない場合、または、
                // 旧パレットＩＤが存在し、その日付が本日より前の日付の場合は、本日日付０００１番のＩＤを発行
                // それ以外の場合は、旧パレットＩＤ＋１で、ＩＤを発行を発行
                //DateTime palletOldDate = DateTime.ParseExact("20" + VBS.Mid(palletOld,4,6), "yyyyMMdd", DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
                //bool palletOldDateSmallerThanToday = palletOldDate < DateTime.Today;

                int numberOld = 0;
                string palletNew;
                //if (palletOld != string.Empty && !palletOldDateSmallerThanToday)
                if (palletOld != string.Empty)
                {
                    numberOld = int.Parse(VBS.Mid(palletOld, 12, 4));
                    palletNew = VBS.Left(palletOld, 11) + (numberOld + 1).ToString("0000");
                }
                else
                {
                    palletNew = model+"_" + registerDate.Date.ToString("yyMMdd") + "_L" + "0001";
                }
                System.Diagnostics.Debug.Print(palletNew);

                // 新パレットＩＤの登録
                string sql2 = "insert into t_pallet (pallet_id, lot, l_cnt, m_qty, batch, register_date, rg_user) " +
                    "VALUES('" + palletNew + "','" + lot + "','" + lotCnt + "','" + qty + "','" + batch + "','" +
                         registerDate.ToString("yyyy/MM/dd HH:mm:ss") + "','" + user + "')";
                System.Diagnostics.Debug.Print(sql2);

                command = new NpgsqlCommand(sql2, connection);
                int res = command.ExecuteNonQuery();

                // 正常な処理が行われた場合はコミット、それ以外は例外を投げ、例外処理内でロールバックする
                if (res == 1)
                {
                    transaction.Commit();
                    connection.Close();
                    return palletNew;
                }
                else
                {
                    throw new InvalidOperationException();
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return string.Empty;
            }
        }

        // サブプロシージャ：週番号を日付へ変換する
        private DateTime getFirstDateOfWeek(int year, int weekOfYear, System.Globalization.CultureInfo ci)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = (int)ci.DateTimeFormat.FirstDayOfWeek - (int)jan1.DayOfWeek;
            DateTime firstWeekDay = jan1.AddDays(daysOffset);
            int firstWeek = ci.Calendar.GetWeekOfYear(jan1, ci.DateTimeFormat.CalendarWeekRule, ci.DateTimeFormat.FirstDayOfWeek);
            if (firstWeek <= 1 || firstWeek > 50) weekOfYear -= 1;
            return firstWeekDay.AddDays(weekOfYear * 7);
        }

        // カートンのディープキャンセル（カートン・パック・トレーのキャンセル日を登録し、モジュールレコードを削除
        public bool sqlMultipleDeepCancelCartonPackTray(string[] cellVal, string user)
        {
            string cancelTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string cartonList = string.Empty;

            for (int i = 0; i < cellVal.Length; i++)
                cartonList += ("'" + cellVal[i] + "', ");

            cartonList = VBS.Left(cartonList, cartonList.Length - 2);

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                // カートンのキャンセル
                string sql1 = "update t_carton set cancel_date = '" + cancelTime + "', " +
                    "cl_user = '" + user + "' where carton_id in (" + cartonList + ")";
                System.Diagnostics.Debug.Print(sql1);
                NpgsqlCommand command = new NpgsqlCommand(sql1, connection);
                int res1 = command.ExecuteNonQuery();

                // パックのキャンセル
                string sql2 = "update t_pack set cancel_date = '" + cancelTime + "', cl_user = '" + user + "' " +
                    "where pack_id in (select pack_id from t_pack where carton_id in (" + cartonList + "))";
                System.Diagnostics.Debug.Print(sql2);
                command = new NpgsqlCommand(sql2, connection);
                int res2 = command.ExecuteNonQuery();

                // トレーのキャンセル
                string sql3 = "update t_tray set cancel_date = '" + cancelTime + "', cl_user = '" + user + "' " +
                    "where tray_id in (select tray_id from t_tray where pack_id in (" +
                        "select pack_id from t_pack where carton_id in (" + cartonList + ")))";
                System.Diagnostics.Debug.Print(sql3);
                command = new NpgsqlCommand(sql3, connection);
                int res3 = command.ExecuteNonQuery();

                // モジュールの消去
                string sql4 = "delete from t_module where tray_id in (" +
                        "select tray_id from t_tray where pack_id in (" +
                            "select pack_id from t_pack where carton_id in (" + cartonList + ")))";
                System.Diagnostics.Debug.Print(sql4);
                command = new NpgsqlCommand(sql4, connection);
                int res4 = command.ExecuteNonQuery();

                //正常に処理が完了した場合
                transaction.Commit();
                connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }


        // トレーのキャンセルの前のチェックを追加（呉）
        public string sqlChkPackofTray(string tray_id)
        {
            string sql0 = "select pack_id,status from pnt_pack where tray_id='" + tray_id + "' order by p_date desc limit 1;";

            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();

            try
            {
                NpgsqlCommand command = new NpgsqlCommand(sql0, connection);
                IDataReader dr = command.ExecuteReader();
                //データ取得
                string res0=string.Empty;
                string pack_id = string.Empty;
                int i = 0;
                while (dr.Read())
                {
                    pack_id=dr[0].ToString();
                    res0=dr[1].ToString();
                    i++;
                }
                dr.Close();
                dr = null;
                connection.Close();

                int valchk;
                if (i == 0)
                {
                    valchk = -1;     //IDataReaderはempty
                }
                else
                {
                    valchk = int.Parse(res0);  //実際の値
                }

                //statusが1の場合（梱包済みの状態）、トレーの取り消し作業を中止
                if (valchk == 1)
                {
                    MessageBox.Show("Tray[" + tray_id + "] has been related to Pack[" + pack_id + "].\nStatus is [" + res0 + "].", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return "NG";
                }
                else
                {
                    return "OK";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return "NG";
            }
        }
                
        public bool sqlChkEffective(string tray_id)
        {
            string sql = string.Format(@"select count(*) from pnt_pack
                                where tray_id = '{0}'
                                and status = '1'"
                                , tray_id);
            connection = new NpgsqlConnection(conStringTrayGuardDb);
            connection.Open();

            try
            {
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                IDataReader dr = command.ExecuteReader();
                string count="";
                while (dr.Read())
                {
                    count = dr[0].ToString();
                }
                dr.Close();
                dr = null;
                connection.Close();

                if (count == "0")
                    return true;
                else
                {
                    MessageBox.Show("这盆马达已经捆包，请勿取消", "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }
    }
}
