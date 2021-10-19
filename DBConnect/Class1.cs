using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DBConnect
{
    public class OleDBCommand
    {
        string _sqlConnect; //接続文字を格納します。

        //コンストラクター
        public OleDBCommand(String sqlConnect)
            {
                _sqlConnect = sqlConnect;
            }

        /// <summary>
        /// 接続文字を格納します。
        /// </summary>
        //（WriteOnly Property)
        public string SqlConnect
        {
            set
            {
                _sqlConnect = value;
            }
        }

        /// <summary>
        /// SQLコマンドを実行し、実行結果を取得します。
        /// </summary>
        /// <param name="sqlCommand">実行するSQL文（SELECT文）</param>
        /// <returns></returns>
        public DataTable OleDBDataTable(string sqlCommand)
        {

            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection(_sqlConnect);
            System.Data.OleDb.OleDbCommand Com = new System.Data.OleDb.OleDbCommand(sqlCommand, cn);
            DataTable respTable = new DataTable();

            cn.Open();

            try
            {
                respTable.Load(Com.ExecuteReader());
                return respTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);                
                return null;
            }
            finally
            {
                cn.Close();
            }

        }

        /// <summary>
        /// SQLコマンドを実行します。
        /// </summary>
        /// <param name="sqlCommands">実行するSQL文（UPDATE,INSERT,DELETE文）</param>
        /// <returns>正常終了:=true,異常終了:=false</returns>
        public Boolean OleDbExcuteNonQuery(params string[] sqlCommands)
        {
            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection(_sqlConnect);
            System.Data.OleDb.OleDbTransaction OleTran;
            System.Data.OleDb.OleDbCommand Com = new System.Data.OleDb.OleDbCommand();
            Com.Connection = cn;

            cn.Open();
            OleTran = cn.BeginTransaction();
            Com.Transaction = OleTran;

            try
            {
                for (int i1 = 0; i1 < sqlCommands.Length; i1++)
                {
                    Com.CommandText = sqlCommands[i1];
                    Com.ExecuteNonQuery();
                }

                OleTran.Commit();
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);   
                OleTran.Rollback();
                return false;
            }
            finally
            {
                OleTran.Dispose();
                cn.Close();
            }

        }

        /// <summary>
        /// SQLコマンドを実行し、実行結果を１つだけ取得します。
        /// </summary>
        /// <param name="sqlCommand">実行結果が１行１列となるSQL文</param>
        /// <returns>実行結果（1つだけ返す。SQL文の集計関数等に有効）</returns>
        public object OleDBExcuteScalar(string sqlCommand)
        {
            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection(_sqlConnect);
            System.Data.OleDb.OleDbCommand Com = new System.Data.OleDb.OleDbCommand(sqlCommand,cn);
            Object RespObject;

            cn.Open();

            try
            {
                RespObject = Com.ExecuteScalar();
                return RespObject;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);   
                return null;
            }
            finally
            {
                cn.Close();
            }
        }
    }
}
