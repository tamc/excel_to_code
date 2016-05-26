#include "excelspreadsheet.c"
#include "parser.c"

int main(int argc, char *argv[]) {
  if(argc > 1) {
    ExcelValue result = parse(argv[1]);
    set_sheet_two_a1(result);
  }
  inspect_excel_value(out());
}
