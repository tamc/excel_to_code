libexcelspreadsheet.dylib: excelspreadsheet.o
	gcc -shared -o libexcelspreadsheet.dylib excelspreadsheet.o

excelspreadsheet.o:
	gcc -fPIC -c excelspreadsheet.c

clean:
	rm excelspreadsheet.o
	rm libexcelspreadsheet.dylib
