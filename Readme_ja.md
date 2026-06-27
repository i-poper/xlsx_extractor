# xlsx_extractor

Tool to extract data from xlsx(xlsm) by specifying headers.

## Description

**xlsx_extractor** is extracts data by specifying a header string.

For example, if you have the following Excel files and specify `test1` and `test2` as headers, you will get `ab,cd` and extraction results.

![figure1.png](graph/figure1.png)

```
$ xlsx_extractor --f test.xlsx -d , test1 test2
test1,test2
ab,cd
```

Find the rows in the sheet where all the specified headers exist, and output the data after the headers in the order of the specified headers.

```
$ xlsx_extractor --f test.xlsx -s test_sheet -d , test2 test1 test3
test2,test1,test3
cd,ab,3
b,a,c
```

## 使用方法
```
Tool to extract data from xlsx(xlsm) by specifying headers.

Usage: xlsx_extractor [OPTIONS] --file <XLSX> [HEADERS]...

Arguments:
  [HEADERS]...  Header names

Options:
  -f, --file <XLSX>            Excel file(.xlsx or .xlsm)
  -d, --delimiter <DELIMITER>  Output delimiter [default: "\t"]
  -s, --sheet <SHEET>          Sheet name
  -H, --header                 Suppress header output
  -q, --quote <QUOTE>          Quote [default: "]
  -t, --style <STYLE>          Quote Style [default: necessary] [possible values: always, necessary, non-numeric, never]
  -o, --output <FILE>          Place the output into <FILE>
  -h, --help                   Print help
  -V, --version                Print version
```
このツールはHEADERSで指定されたヘッダ名が設定されたセルをシートの左上から順に探していきます。全てのヘッダが見つかった行をヘッダ業と認識します。

![figure2.png](graph/figure2.png)

```
$ xlsx_extractor --f test.xlsx -s test_sheet -d , test2 test1 test3
```
図のようなシートに対してあなたが上記のコマンドを実行すると、ツールはまずA1のセルから右方向にヘッダ名を探していきます。行内にヘッダが見つからない場合は次の行に進みます。図ではI3セルにヘッダに指定された文字がありますが指定されたヘッダが全て揃っているわけではないのでヘッダ行としては認識しません。
7行目には全てのヘッダの文字列が揃っているのでこの行をヘッダ行として認識します。
ヘッダ行の次の行からデータ行です。図では矢印が付いているB列、C列、D列が抽出されます。列の出力順序はHEADERSで指定した順序になります。

結果として
```
test2,test1,test3
cd,ab,3
b,a,c
```
が、出力されます。

### エスケープシーケンス

ヘッダーにはエスケープ・シーケンスを使うことができる。セルE7をヘッダーとして指定したい場合は、以下のようにします。
```
$ xlsx_extractor --f test.xlsx -s test_sheet -d , test2 test1 test3 Test\\r\\n4
test2,test1,test3,"Test
4"
cd,ab,3,4
b,a,c,c
```
改行が含まれているヘッダはダブルクオーテーションで括られています。デフォルトでは必要がある場合にクオーティングする設定になっています。

あなたはエスケープシーケンスは、デリミタやクオートでも使用できます。
```
$ xlsx_extractor --f test.xlsx -s test_sheet -d \\t -q \' test2 test1 Test\\r\\n4
test2   test1   test3   'Test
4'
cd      ab      3       4
b       a       c       c
```
### クオーティングスタイル
クオーティングのスタイルを指定できます。
- always
  ```
  $ xlsx_extractor --f test.xlsx -s test_sheet -d , -t always test2 test1 Test\\r\\n4
  "test2","test1","test3","Test
  4"
  "cd","ab","3","4"
  "b","a","c","c"
  ```
- necessary
  ```
  $ xlsx_extractor --f test.xlsx -s test_sheet -d , -t necessary test2 test1 Test\\r\\n4
  test2,test1,test3,"Test
  4"
  cd,ab,3,4
  b,a,c,c
  ```
- non-numeric
  ```
  $ xlsx_extractor --f test.xlsx -s test_sheet -d , -t non-numeric test2 test1 Test\\r\\n4
  "test2","test1","test3","Test
  4"
  "cd","ab",3,4
  "b","a","c","c"
  ```
- never
  ```
  $ xlsx_extractor --f test.xlsx -s test_sheet -d , -t never test2 test1 Test\\r\\n4
  test2,test1,test3,Test
  4
  cd,ab,3,4
  b,a,c,c
  ```
