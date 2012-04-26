libeu.dylib: eu.o
	gcc -shared -o libeu.dylib eu.o

eu.o:
	gcc -Wall -fPIC -c eu.c

clean:
	rm eu.o
	rm libeu.dylib
