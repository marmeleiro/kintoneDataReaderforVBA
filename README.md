# kintoneDataReaderforVBA

サイボウズ・kintoneのデータをExcel側から抽出するためのクラスです。（REST API利用）  
JSON形式で取得したデータ処理の面倒な部分を少し扱いやすくしています。

### Features 

* REST　APIの一度に取得可能な上限の500件を超えた場合も、ループ処理を書かずに取得できます。
* 複数選択やチェックボックスなどの複数の値があるフィールドを、カンマ区切りで1項目として取得できます。
* リッチエディタのタグ除去ができます。
* アプリのフィールド情報取得も可能です。
* ゲストスペースは非対応です。

### Requirement 

* Dictionary.cls（v1.4.1）
  * https://github.com/VBA-tools/VBA-Dictionary/releases/tag/v1.4.1

* JsonConverter.bas（v2.2.2）
  * https://github.com/VBA-tools/VBA-JSON/releases/tag/v2.2.2


### Getting started 

* 上記の依存ファイルをダウンロードします。
* 当リポジトリのkitoneDataReaderforVBA.clsと上記2つのファイルをExcelVBAにインポートします。

### Usage 

* kintone接続

auth引数はログインユーザーID:パスワードの形式で指定します。

    Dim kintoneUtil As kintoneDataReaderforVBA 
    Set kintoneUtil = New kintoneDataReaderforVBA
    Call kintoneUtil.Setup(subdomain:="XX", app:="XX", auth:="XX:XX")
 
 
* レコード取得


データを500件を超えて取得する場合は、isAll引数にTrueを設定します。  
sFields引数とquery引数をオプションで指定することもできます。  
sFields引数には、取得したいフィールドコードをカンマ区切りで指定します。  
query引数には、抽出条件を記述します。記述方法は、cybozu developers network（ https://developer.cybozu.io/hc/ja/articles/202331474#step2 ）を参照。

    Call kintoneUtil.GetRecords(isAll:=True)
    

*note:*  
全件取得した場合、取得件数・項目数によってはPCのメモリ不足が起きる可能性があります。



* 取得したデータのセル貼り付け
 
RecordCountで取得件数が判定できます。
FieldValueには、データインデックスとフィールドコードを指定します。

    For i = 1 To kintoneUtil.RecordCount 
        Sheets(1).Cells(i, 1).Value =  kintoneUtil.FieldValue(i, "文字列_0") 
        Sheets(1).Cells(i, 2).Value =  kintoneUtil.FieldValue(i, "ドロップダウン_0") 
        Sheets(1).Cells(i, 3).Value =  kintoneUtil.FieldValue(i, "作成日時") 
    Next 


* その他 


    '取得したレコードのフィールドコード一覧取得  
    Dim fields() As String     
    fields = kintoneUtil.RecordFields    
    
	'フィールドタイプ取得
	Dim fieldtype As String   
    fieldtype = kintoneUtil.FieldType("XXX")   
    
    'フィールドラベル取得   
	Dim fieldlabel As String   
    fieldlabel = kintoneUtil.FieldLabel("XXX")   


### example.xlsm
簡単なサンプルです。
依存ファイル2つをダウンロードし、example.xlsmにインポートしてください。 
