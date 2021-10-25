# DBConnect (C#)

# OleDBCommand クラス
 OleDBCommand は OLE DB APIを用いてデータベースの操作を行うクラスです。

# コンストラクター
## OleDBCommand()
 インスタンスを初期化します。
## OleDBCommand(string)
 接続文字を指定してインスタンスを初期化します。

# プロパティ
## SqlConnect
 接続文字を指定します。書き込み専用プロパティです。

# メソッド
## OleDBDataTable
### 概要
 SELECTを扱うSQLを実行し、結果をDataTableへ格納します。
### 引数
 sqlCommand：string型です。SELECTを扱うSQLを指定します。
### 戻り値
 DataTable：実行したSQLの結果をDataTable型で返します。

## OleDbExcuteNonQuery
### 概要
 UPDATE,INSERT,DELETEを扱うSQLを実行し、成功可否を返します。
### 引数
 sqlCommands：文字列型の可変長引数です。UPDATE,INSERT,DELETEを扱うSQLを指定します。１要素につき１つのSQLを指定します。
### 戻り値
 Boolean：SQLがエラーなく成功した場合はtrue、エラーが発生した場合はfalseを返します。

## OleDBExcuteScalar
### 概要
 集計関数などの実行結果が１行１列となるSQLを実行し、結果を返します。
### 引数
 sqlCommand：string型です。集計関数などの実行結果が１行１列となるSQLを指定します。
### 戻り値
 object：集計関数などの実行結果が１行１列となるSQLの実行結果を返します。
 
# 使用例
```
//インスタンス化する際に接続文字を入力する。
DBConnect.OleDBCommand Ref = new DBConnect.OleDBCommand(
    "｛ここへ接続文字列を入力する｝"
);

//OleDBDataTableメソッドの使用例
System.Data.DataTable ResultTable;
ResultTable = Ref.OleDBDataTable("｛ここへSELECTを文を使うSQLを入力する｝");

//実行結果の各行の１列目を参照
for(int i1 = 0; i1 < ResultTable.Rows.Count; i1++)
{
    ResultTable.Rows[i1][0].ToString();
}

//OleDBExcuteScalarメソッドの使用例
object ob;
ob = Ref.OleDBExcuteScalar("ここへ集計関数などの１行１列の結果になるSQLを入力する。");

//OleDbExcuteNonQueryメソッドの使用例
Ref.OleDbExcuteNonQuery("｛ここへINSERTやUPDATTE、DELETEを文を使うSQLを入力する｝");
Ref.OleDbExcuteNonQuery("｛ここへINSERTやUPDATTE、DELETEを文を使うSQLを入力する｝", "｛ここへINSERTやUPDATTE、DELETEを文を使うSQLを入力する｝");

string[] sql = new string[3];
sql[0] = "｛ここへINSERTやUPDATTE、DELETEを文を使うSQLを入力する｝";
sql[1] = "｛ここへINSERTやUPDATTE、DELETEを文を使うSQLを入力する｝";
sql[2] = "｛ここへINSERTやUPDATTE、DELETEを文を使うSQLを入力する｝";
Ref.OleDbExcuteNonQuery(sql);
```
