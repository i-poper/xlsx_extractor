# Extract duplicate headers
```console
$ xlsx_extractor -f tests/test.xlsm -s test_sheet -d , test2 test3 test3 Test\\r\\n4
test2,test3,test3,"Test
4"
cd,3.0,x,4.0
b,c,y,c

```

# default sheet
```console
$ xlsx_extractor -f tests/test.xlsm -d _ test1 test2 test3
test1_test2_test3
ab_cd_3

```

# Rows with no data are not output
```console
$ xlsx_extractor -f tests/test.xlsm -d , -s Sheet3 test1 test2 "Test\r\n4"
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
