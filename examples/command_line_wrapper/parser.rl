#include <string.h>
#include <stdio.h>

%%{
  machine instructions;

	action numberAction { 
    double number; 
    sscanf(p, "%lf", &number); 
    latestValue = EXCEL_NUMBER(number);
  }

  action startString { 
    stringStart = p;
  }

  action endString {
    int length = p - stringStart;
    char *s = malloc(length + 1);
    if(s == 0) { printf("Out of memory"); exit(-1); }
    memcpy(s, stringStart, length);
    string[length] = '\0';
    free_later(s);
    latestValue = EXCEL_STRING(s);
  }
  
  action startVector { printf("Start vector\n"); }
  action endVector { printf("End vector\n"); }
  action startMatrix { printf("Start matrix\n"); }
  action endMatrix { printf("End matrix\n"); }

	emptyValue = ('""' | 'nil') @{ latestValue = BLANK; };

  number = (('-' | '+')? digit+ ('.' digit+)? ([eE][\-+]?digit+)?) >numberAction;
  quote = '"';
  notQuote = [^"] | '\\"';
	string = quote . (notQuote+ >startString %endString) . quote;

  true = 'true' @{ latestValue = TRUE; };
  false = 'false' @{ latestValue = FALSE; };
  boolean = true | false;

	valueError = ('VALUE' | '#VALUE!') @{ latestValue = VALUE; };
	nameError = 'NAME' | '#NAME?' @{ latestValue = NAME; };
	div0Error = 'DIV0' | '#DIV/0!' @{ latestValue = DIV0; };
	refError = 'REF' | '#REF!' @{ latestValue = REF; };
	naError = 'NA' | '#N/A' @{ latestValue = NA; };
	numError = 'NUM' | '#NUM!' @{ latestValue = NUM; };

	error = valueError | nameError | div0Error | refError | naError | numError;

	excelValue = (emptyValue | number | string | boolean | error) >{ printf("v"); };

  vector = ('[' . excelValue . (',' . excelValue  )+ . ']') >startVector @endVector;
	
	row = ('[' . excelValue . (',' . excelValue )+ . ']') >{printf("Row start");} @{printf("Row end");};
  matrix = ('[' . row . (',' . row)* . ']') >startMatrix @endMatrix;

	cellReference = [A-Z]+ . [0-9]+;

 main := (  matrix | vector | excelValue ); 

}%%

%% write data;

ExcelValue parse(char *string ) {
  int cs;
  char *p = string;
  char *pe = p + strlen(p) + 1;
  char *stringStart;
  char *eof = 0;

  ExcelValue latestValue = BLANK;

  %% write init;
  %% write exec;
  return latestValue;
}
