using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TrayGuard
{
    class TfTime
    {
        public DateTime serverTime;

        // コンストラクタ
        public TfTime()
        {
            TfSQL tf = new TfSQL();
            string sql = "select current_timestamp";
            serverTime = tf.sqlExecuteScalarDateTime(sql);
        }
    }
}
