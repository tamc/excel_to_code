#include "excelspreadsheet.c"
#include "manualparser.c"

void fail(ExcelValue actual, ExcelValue expected) {
  printf("Expected: ");
  inspect_excel_value(expected);
  printf("Actual: ");
  inspect_excel_value(actual);
  exit(-1);
}


void test(ExcelValue actual, ExcelValue expected) {
  int i, j, k;
  ExcelValue *actual_array, *expected_array;

  if(actual.type != expected.type) { fail(actual, expected); }
  switch(expected.type) {
    case ExcelNumber:
    case ExcelBoolean:
    case ExcelEmpty:
  	case ExcelError:
      if(actual.number != expected.number) { fail(actual, expected); }
      break;
	  case ExcelString:
	  	if(strcasecmp(actual.string,expected.string) != 0 ) { fail(actual, expected); }
      break;
  	case ExcelRange:
      if(actual.rows != expected.rows) { fail(actual, expected); }
      if(actual.columns != expected.columns) { fail(actual, expected); }
      actual_array = actual.array;
      expected_array = expected.array;
      for(i = 0; i < actual.rows; i++) { 
        for(j = 0; j < actual.columns; j++) {
				  k = (i * expected.columns) + j;
          test(actual_array[k], expected_array[k]);
        }
      }
      break;
    default:
      printf("Type %d not recognised", expected.type);
      exit(-1);
  }
}

void testRoundTrip(ExcelValue value) {
  char *string = excelValueToString(value);
  printf("String '%s'\n", string);
  ExcelValue parsedValue = parseValue(string);
  test(parsedValue, value);
}

int main() {
  //test(parseValue(" 1.23"), EXCEL_NUMBER(1.23));
  //test(parseValue("-1.23"), EXCEL_NUMBER(-1.23));
  //test(parseValue("-0.23e-3"), EXCEL_NUMBER(-2.3e-4));
  //test(parseValue("\n\"Hello world\""), EXCEL_STRING("Hello world"));
  ////test(parValuese("\"Hello\\\"world\""), EXCEL_STRING("Hello\"world"));
  //test(parseValue("true"), TRUE);
  //test(parseValue("false"), FALSE);
  //test(parseValue("#VALUE!"), VALUE);
  //test(parseValue("#NAME?"), NAME);
  //test(parseValue("#DIV/0!"), DIV0);
  //test(parseValue("#N/A"), NA);
  //test(parseValue("#NUM!"), NUM);
  ExcelValue range[] = { ONE, TWO, THREE }; 
  test(parseValue("[1,2,3]"), EXCEL_RANGE(range, 1, 3));
  ExcelValue emptyRange[] = {}; 
  test(parseValue("[]"), EXCEL_RANGE(emptyRange, 0, 0));
  ExcelValue manyTypedRange[] = { ONE, TRUE, THREE, EXCEL_STRING("Four") }; 
  test(parseValue("[1,true,3,\"Four\"]"), EXCEL_RANGE(manyTypedRange, 1, 4));
  test(parseValue("[[]]"), EXCEL_RANGE(emptyRange, 1, 0));
  ExcelValue columnRange[] = { ONE, TWO, THREE }; 
  test(parseValue("[[1],[2],[3]]"), EXCEL_RANGE(columnRange, 3, 1));
  ExcelValue matrixRange[] = { ONE, TWO, THREE, FOUR }; 
  test(parseValue("[[1,2],[3,4]]"), EXCEL_RANGE(matrixRange, 2, 2));

  //// Test some error modes
  //test(parseValue("[\"Unterminated array\",2"), VALUE);
  //test(parseValue("\"Unterminated string"), VALUE);
  //test(parseValue("falt"), VALUE);
  //test(parseValue("#NAA"), VALUE);
  //test(parseValue("[1,[1,2]]"), VALUE);
  //test(parseValue("[[[1],[2]]]"), VALUE);
  
  // Test the output of excelValueToString
  testRoundTrip(NA);
  testRoundTrip(VALUE);
  testRoundTrip(NUM);
  testRoundTrip(DIV0);
  testRoundTrip(TRUE);
  testRoundTrip(FALSE);
  testRoundTrip(TRUE);
  testRoundTrip(EXCEL_NUMBER(1));
  testRoundTrip(EXCEL_NUMBER(1.231241));
  testRoundTrip(EXCEL_STRING("Hello world!"));
  testRoundTrip(EXCEL_STRING(""));
  testRoundTrip(EXCEL_RANGE(range, 1, 3));
  testRoundTrip(EXCEL_RANGE(emptyRange, 0, 0));
  testRoundTrip(EXCEL_RANGE(columnRange, 3, 1));
  testRoundTrip(EXCEL_RANGE(matrixRange, 2, 2));


  printf("All tests passed\n");
};
