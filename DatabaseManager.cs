using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace YumejitateApp
{
    /// <summary>
    /// 夢仕立て.mdb への接続・クエリを管理するクラス
    /// </summary>
    public class DatabaseManager : IDisposable
    {
        private const string DbPassword = "shigeshi";
        private const string Provider = "Microsoft.Jet.OLEDB.4.0";

        private OleDbConnection _connection;
        private bool _disposed;

        // ---------------------------------------------------------------
        // 接続管理
        // ---------------------------------------------------------------

        /// <summary>指定パスのDBに接続する</summary>
        public void Connect(string dbFilePath)
        {
            if (_connection != null && _connection.State == ConnectionState.Open)
                return;

            string connStr = "Provider=" + Provider +
                             ";Data Source=" + dbFilePath +
                             ";Jet OLEDB:Database Password=" + DbPassword + ";";

            _connection = new OleDbConnection(connStr);
            _connection.Open();
        }

        /// <summary>DB接続を閉じる</summary>
        public void Disconnect()
        {
            if (_connection != null && _connection.State != ConnectionState.Closed)
                _connection.Close();
        }

        public bool IsConnected
        {
            get { return _connection != null && _connection.State == ConnectionState.Open; }
        }

        // ---------------------------------------------------------------
        // スキーマ取得
        // ---------------------------------------------------------------

        /// <summary>DB内の全テーブル名を返す</summary>
        public List<string> GetTableNames()
        {
            EnsureConnected();

            var tables = new List<string>();
            DataTable schema = _connection.GetOleDbSchemaTable(
                OleDbSchemaGuid.Tables,
                new object[] { null, null, null, "TABLE" });

            if (schema != null)
            {
                foreach (DataRow row in schema.Rows)
                    tables.Add(row["TABLE_NAME"].ToString());
            }

            return tables;
        }

        // ---------------------------------------------------------------
        // クエリ実行
        // ---------------------------------------------------------------

        /// <summary>SELECT クエリを実行して DataTable で返す</summary>
        public DataTable ExecuteQuery(string sql)
        {
            EnsureConnected();

            using (var cmd = new OleDbCommand(sql, _connection))
            {
                var adapter = new OleDbDataAdapter(cmd);
                var dt = new DataTable();
                adapter.Fill(dt);
                return dt;
            }
        }

        /// <summary>INSERT/UPDATE/DELETE を実行して影響行数を返す</summary>
        public int ExecuteNonQuery(string sql)
        {
            EnsureConnected();

            using (var cmd = new OleDbCommand(sql, _connection))
            {
                return cmd.ExecuteNonQuery();
            }
        }

        /// <summary>品番で商品を検索する</summary>
        public DataTable SearchByHinban(string hinban)
        {
            EnsureConnected();

            string sql = "SELECT * FROM 基本テーブル WHERE a LIKE '%" + hinban + "%'";
            return ExecuteQuery(sql);
        }

        /// <summary>地金相場テーブルを取得する</summary>
        public DataTable GetSouba()
        {
            return ExecuteQuery("SELECT * FROM 地金相場テーブル");
        }

        // ---------------------------------------------------------------
        // 内部ユーティリティ
        // ---------------------------------------------------------------

        private void EnsureConnected()
        {
            if (!IsConnected)
                throw new InvalidOperationException(
                    "DBに接続されていません。Connect() を先に呼び出してください。");
        }

        // ---------------------------------------------------------------
        // IDisposable
        // ---------------------------------------------------------------

        /// <summary>パラメータ付きINSERT/UPDATE/DELETEを実行</summary>
        /// <summary>INSERT/UPDATE/DELETE を実行して影響行数を返す</summary>
        public int ExecuteNonQuery(string sql, params OleDbParameter[] parameters)
        {
            EnsureConnected();

            using (var cmd = new OleDbCommand(sql, _connection))
            {
                if (parameters != null)
                {
                    foreach (var p in parameters)
                        cmd.Parameters.Add(p);
                }
                return cmd.ExecuteNonQuery();
            }
        }

        /// <summary>パラメータ数を明示したExecuteNonQuery（3個）</summary>
        public int ExecuteNonQuery(string sql, OleDbParameter p1, OleDbParameter p2, OleDbParameter p3)
        {
            return ExecuteNonQuery(sql, new OleDbParameter[] { p1, p2, p3 });
        }

        /// <summary>パラメータ数を明示したExecuteNonQuery（4個）</summary>
        public int ExecuteNonQuery(string sql, OleDbParameter p1, OleDbParameter p2,
                                   OleDbParameter p3, OleDbParameter p4)
        {
            return ExecuteNonQuery(sql, new OleDbParameter[] { p1, p2, p3, p4 });
        }

        /// <summary>パラメータ数を明示したExecuteNonQuery（8個）</summary>
        public int ExecuteNonQuery(string sql, OleDbParameter p1, OleDbParameter p2,
                                   OleDbParameter p3, OleDbParameter p4, OleDbParameter p5,
                                   OleDbParameter p6, OleDbParameter p7, OleDbParameter p8)
        {
            return ExecuteNonQuery(sql, new OleDbParameter[] { p1, p2, p3, p4, p5, p6, p7, p8 });
        }

        public void Dispose()
        {
            if (_disposed) return;
            Disconnect();
            if (_connection != null)
            {
                _connection.Dispose();
                _connection = null;
            }
            _disposed = true;
        }
    }
}