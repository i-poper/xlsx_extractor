# escape sequence
```console
$ xlsx_extractor -f tests/test.xlsx -s test_sheet -d , test5 test1 Test\\r\\n4
test5,test1,"Test
4"
xxxx,zzzz,aaaa
d,f,g

```

# quote
```console
$ xlsx_extractor -f tests/test.xlsx -s test_sheet -d \\t -q \' test2 test1 test3 Test\\r\\n4
test2	test1	test3	'Test
4'
cd	ab	3.0	4.0
b	a	c	c

```

# Extract duplicate headers
```console
$ xlsx_extractor -f tests/test.xlsx -s test_sheet -d , test2 test3 test3 Test\\r\\n4
test2,test3,test3,"Test
4"
cd,3.0,x,4.0
b,c,y,c

```

# default sheet
```console
$ xlsx_extractor -f tests/test.xlsx -d _ test1 test2 test3
test1_test2_test3
ab_cd_3

```

# No data
```console
$ xlsx_extractor -f tests/test.xlsx -s Sheet2 -d / test1 test2 test3
test1/test2/test3

```

# Rows with no data are not output
```console
$ xlsx_extractor -f tests/test.xlsx -d , -s Sheet3 test1 test2 "Test\r\n4"
test1,test2,"Test
4"
1,,
,2.3,
,,4
a,,
,d,e
x,y,
end1,end2,

```

# Formatted values and Japanese sheet name
```console
$ xlsx_extractor -f tests/test.xlsm -s シート4 -d , テスト1 時刻 日付 テスト\\r\\n4
テスト1,時刻,日付,"テスト
4"
aaa,16:30,2026/5/14,-
bbb,18:00,2026年5月14日,1時30分
ccc,20:00,5月14日,2時00分

```

# No data
```console
$ xlsx_extractor -f tests/test.xlsx -s Sheet2 -d / test1 test2 test3
test1/test2/test3

```

# No Option
```console
$ xlsx_extractor
? failed
error: the following required arguments were not provided:
  --file <XLSX>

Usage: xlsx_extractor[EXE] --file <XLSX> [HEADERS]...

For more information, try '--help'.

```

# help
```console
$ xlsx_extractor --help
Tool to extract data from xlsx(xlsm) by specifying headers.

Usage: xlsx_extractor[EXE] [OPTIONS] --file <XLSX> [HEADERS]...

Arguments:
  [HEADERS]...  Header names

Options:
  -f, --file <XLSX>            Excel file(.xlsx or .xlsm)
  -d, --delimiter <DELIMITER>  Output delimiter [default: "/t"]
  -s, --sheet <SHEET>          Sheet name
  -H, --header                 Suppress header output
  -q, --quote <QUOTE>          Quote [default: "]
  -t, --style <STYLE>          Quote Style [default: necessary] [possible values: always, necessary, non-numeric, never]
  -o, --output <FILE>          Place the output into <FILE>
  -h, --help                   Print help
  -V, --version                Print version

```

# Missing headers
```console
$ xlsx_extractor -f tests/test.xlsx test1 test2 testX
? failed
error: `[HEADERS]...` not found.

Usage: xlsx_extractor [OPTIONS] --file <XLSX> [HEADERS]...

For more information, try '--help'.

```

# Invalid delimiter
```console
$ xlsx_extractor -f tests/test.xlsx -d "\d" test1 test2
? failed
error: invalid value '/d' for '--delimiter <DELIMITER>': `/d` is not a valid escape string.

For more information, try '--help'.

```

# Invalid quote
```console
$ xlsx_extractor -f tests/test.xlsx -q "\x" test1 test2
? failed
error: invalid value '/x' for '--quote <QUOTE>': `/x` is not a valid escape string.

For more information, try '--help'.

```

# Erroneous escape sequence in specified header
```console
$ xlsx_extractor -f tests/test.xlsx test1 test2 "Test\r\x4"
? failed
error: invalid value 'Test/r/x4' for '[HEADERS]...': `Test/r/x4` is not a valid escape string.

For more information, try '--help'.

```

# No specified sheet
```console
$ xlsx_extractor -f tests/test.xlsx -s test___ test1 test2
? failed
error: Sheet not found:test___

Usage: xlsx_extractor [OPTIONS] --file <XLSX> [HEADERS]...

For more information, try '--help'.

```

# Output Error
```console
$ xlsx_extractor -f tests/test.xlsx -o tests test1 test2
? failed
error: Can't create tests: [..]

Usage: xlsx_extractor [OPTIONS] --file <XLSX> [HEADERS]...

For more information, try '--help'.

```

# Input file not found
```console
$ xlsx_extractor -f test/not-found.xlsx test1 test2
? failed
error: Can't read `test/not-found.xlsx`: [..]

Usage: xlsx_extractor [OPTIONS] --file <XLSX> [HEADERS]...

For more information, try '--help'.

```

# Invalid xlsx file
```console
$ xlsx_extractor -f tests/test.md test1
? failed
error: Can't read `tests/test.md`: [..]

Usage: xlsx_extractor [OPTIONS] --file <XLSX> [HEADERS]...

For more information, try '--help'.

```
