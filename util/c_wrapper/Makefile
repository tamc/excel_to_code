example: example.c
	clang -g example.c -o example

example.c: parser.c excelspreadsheet.c manualparser.c
	touch example.c

parser.c: parser.rl
	ragel parser.rl

libexcelspreadsheet.dylib: excelspreadsheet.o
	gcc -shared -o libexcelspreadsheet.dylib excelspreadsheet.o

excelspreadsheet.o: excelspreadsheet.c
	gcc -fPIC -c excelspreadsheet.c

excelspreadsheet.c: example.xlsx
	../../bin/excel_to_c -n -s 'Sheet! Two' --no-makefile  example.xlsx

clean:
	rm excelspreadsheet.o
	rm libexcelspreadsheet.dylib
