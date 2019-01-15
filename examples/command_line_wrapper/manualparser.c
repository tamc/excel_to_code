#include <string.h>
#include <stdio.h>
#include <stdbool.h>

/**
 * ISSUES with this code:
 * 1. Does not handle non-ascii
 * 2. Does not handle escaped strings
 **/

ExcelValue parse(char **string);
void excelValueToStringBuffer(ExcelValue v, char **buffer, int *used, int *allocated);

const int BUFFER_ALLOC_SIZE = 100;

char* excelValueToString(ExcelValue v) {
  char *buffer = malloc(BUFFER_ALLOC_SIZE);
  if(!buffer) { return "Unable to allocate buffer for string"; }
  int used = 0;
  int allocated = BUFFER_ALLOC_SIZE;
  excelValueToStringBuffer(v, &buffer, &used, &allocated);
  buffer = realloc(buffer, used + 1);
  buffer[used] = '\0';
  return buffer;
}

void excelValueToStringBuffer(ExcelValue v, char **buffer, int *used, int *allocated) {
  ExcelValue *array;
  int i, j, k;
  switch (v.type) {
    case ExcelNumber:
      *used += sprintf(*buffer + *used, "%f",v.number);
      break;
    case ExcelBoolean:
      if(v.number == true) {
        *used += sprintf(*buffer + *used, "%s","true");
      } else if(v.number == false) {
        *used += sprintf(*buffer + *used, "%s","false");
      } else {
        *used += sprintf(*buffer + *used, "Boolean with undefined state %f\n",v.number);
      }
      break;
    case ExcelEmpty:
      *used += sprintf(*buffer + *used, "%s","");
      break;
    case ExcelRange:
      array = v.array;
      if(v.rows == 0 && v.columns == 0) { 
        *used += sprintf(*buffer + *used, "[]");
      } else {
        if(v.rows > 1) {
          *used += sprintf(*buffer + *used, "[");
        }
        for(i = 0; i < v.rows; i++) {
          *used += sprintf(*buffer + *used, "[");
          for(j = 0; j < v.columns; j++ ) {
            k = (i * v.columns) + j;
            excelValueToStringBuffer(array[k], buffer, used, allocated);
            if(j < v.columns - 1) {
              *used += sprintf(*buffer + *used, ",");
            }
          }
          *used += sprintf(*buffer + *used, "]");
          if(i < v.rows - 1) {
            *used += sprintf(*buffer + *used, ",");
          }
        }
        if(v.rows > 1) {
          *used += sprintf(*buffer + *used, "]");
        }
      }
      break;
    case ExcelString:
      *used += sprintf(*buffer + *used, "\"%s\"", v.string);
      break;
    case ExcelError:
      switch( (int)v.number) {
        case 0: 
          *used += sprintf(*buffer + *used, "%s", "#VALUE!");
          break;
        case 1:
          *used += sprintf(*buffer + *used, "%s", "#NAME?");
          break;
        case 2:
          *used += sprintf(*buffer + *used, "%s", "#DIV/0!");
          break;
        case 3:
          *used += sprintf(*buffer + *used, "%s", "#REF!");
          break;
        case 4:
          *used += sprintf(*buffer + *used, "%s", "#N/A");
          break;
        case 5:
          *used += sprintf(*buffer + *used, "%s", "#NUM!");
          break;
      }
      break;
    default:
      printf("Type %d not recognised",v.type);
  };
}

ExcelValue parseFailed(char *message, char *string) {
  fprintf(stderr, "\n%s\n", message);
  if(string) {
    fprintf(stderr, "%s\n\n", string);
  }

  return VALUE;
}

ExcelValue parseExcelError(char **string) {
  char *errorStrings[] = {"#VALUE!","#NAME?", "#DIV/0!", "#N/A", "#NUM!"};
  ExcelValue errorValues[] = { VALUE, NAME, DIV0, NA, NUM };
  int i;
  int numberOfErrorTypes = sizeof(errorValues) / sizeof(ExcelValue);
  for(i=0;i<numberOfErrorTypes;i++) {
    if(strcasecmp(*string, errorStrings[i]) != 0) { continue; }
    *string += strlen(errorStrings[i]);
    return errorValues[i];
  }
  return parseFailed("No error found that matches string", *string);
}

ExcelValue parseNumber(char **string) {
  double number;
  int matchLength;
  sscanf(*string, "%lf%n", &number, &matchLength);
  *string += matchLength;
  return EXCEL_NUMBER(number);
}

ExcelValue parseBoolean(char **string) {
  if(strncasecmp(*string, "true", 4) == 0) { *string += 4; return TRUE; };
  if(strncasecmp(*string, "false", 5) == 0) { *string += 5; return FALSE; };
  return parseFailed("Boolean not recognised", *string);
  return VALUE;
}

// Starts and ends with a "
// FIXME: utf8 compatibility
// FIXME: escaping characters
ExcelValue parseString(char **string) {
  assert(*string[0] == '"'); // We've started with a "
  (*string)++;
  char *start = *string;
  char currentCharacter;
  while(true) {
    currentCharacter = **string;
    if(currentCharacter == '\0') {
      return parseFailed("Expected to see a closing \" before reaching the end of the line", start-1); // -1 for opening "
    }

    if(currentCharacter == EOF) {
      return parseFailed("Expected to see a closing \" before reaching the end of the file", start-1); // -1 for opening "
    }

    if(currentCharacter == '"') { break; }
    (*string)++;
  }
  int numberOfCharacters = *string - start;
  (*string)++; // Advance over the closing "
  char *matchedString = malloc(numberOfCharacters+1); // +1 for closing \0
  if(matchedString == 0) { return parseFailed("Out of memory when allocating string", start-1); }
  memcpy(matchedString, start, numberOfCharacters);
  matchedString[numberOfCharacters] = '\0';
  free_later(matchedString);
  return EXCEL_STRING(matchedString);
}

const int ARRAY_ALLOC_SIZE = 100;
const ExcelValue END_OF_ARRAY = {.type = ExcelError, .number = -1};

ExcelValue startArray(char **string) {
  char *start = *string; // For error reporting
  (*string)++; // Advance over the [
  ExcelValue *array = malloc(ARRAY_ALLOC_SIZE * sizeof(ExcelValue));
  if(array == 0) { return parseFailed("Out of memory allocating the array", start); }
  int length = 0; // The first empty slot in the array
  int allocatedLength = ARRAY_ALLOC_SIZE;
  int rows = 0;
  int columns = 0;
  char *previousPosition;

  while(true) {
    previousPosition = *string;
    ExcelValue nextExcelValue = parse(string);

    if(*string == previousPosition) {
      // We have not made any progress, aborting
      return parseFailed("Expected to see a closing ] before reaching the end of the file", start);
    }

    if(nextExcelValue.type == END_OF_ARRAY.type && nextExcelValue.number == END_OF_ARRAY.number) {
      if(allocatedLength > length) {
        // Shrink our memory allocation
        array = realloc(array, length * sizeof(ExcelValue));
      }
      // Retain the array
      free_later(array);
      // Fix the number of rows
      if(rows == 0 && length > 0) { rows = 1; }
      // Return the right sort of excel value
      return EXCEL_RANGE(array,rows, columns);

    } else if(nextExcelValue.type == ExcelRange) {
      // If we are here, then we are a multi dimentional array

      if(nextExcelValue.rows > 1) {
        return parseFailed("Excel arrays can be, at most 2 dimensional.", start);
      }

      if(length > 0 && nextExcelValue.columns != columns) {
        return parseFailed("Array rows must be identically sized.", start);
      }

      // We have a new row
      rows++;

      // If the first row, define the number of columns
      if(length == 0) { columns = nextExcelValue.columns; }

      // FIXME: Should always have 1 row?
      int nextValueLength = nextExcelValue.rows * nextExcelValue.columns;

      // Resize the array if we need to
      while((length + nextValueLength) > allocatedLength) {
        array = realloc(array, (allocatedLength + ARRAY_ALLOC_SIZE)*sizeof(ExcelValue));
        if(array == 0) { return parseFailed("Out of memory extending the array", start); }
        allocatedLength = allocatedLength + ARRAY_ALLOC_SIZE;
      }

      // Copy over the values
      memcpy(&array[length], nextExcelValue.array, nextValueLength * sizeof(ExcelValue));

      // Update our count
      length = length + nextValueLength;

      // Clean up
      free(nextExcelValue.array);

    } else {
      // If we are here, then we are just adding a normal value to the end of a row
      columns++;

      // Resize the array if we need to
      if(length >= allocatedLength) {
        array = realloc(array, (allocatedLength + ARRAY_ALLOC_SIZE)*sizeof(ExcelValue));
        if(array == 0) { return parseFailed("Out of memory extending the array", start); }
        allocatedLength = allocatedLength + ARRAY_ALLOC_SIZE;
      }

      // Then put this element onto the end of the array
      array[length] = nextExcelValue;

      // Increment our count
      length++;
    }
  } 
  return parseFailed("Expected to see a closing ] before reaching the end of the file", start);
}

ExcelValue endArray(char **string) {
  (*string)++; // Advance over the closing ]
  return END_OF_ARRAY;
}

// Will exit the program if cannot parse
ExcelValue parse(char **string) {
  while(true) {
    switch(**string) {
      case '"':
        return parseString(string);
        break;
      case '-':
      case '+':
      case '0':
      case '1':
      case '2':
      case '3':
      case '4':
      case '5':
      case '6':
      case '7':
      case '8':
      case '9':
        return parseNumber(string);
        break;
      case 't':
      case 'f':
        return parseBoolean(string);
        break;
      case '#':
        return parseExcelError(string);
        break;
      case ' ':
      case '\t':
      case '\n':
      case '\r':
      case ',': // Important to progress through arrays
        (*string)++;
        break;
      case '[':
        return startArray(string);
        break;
      case ']':
        return endArray(string);
        break;
      case EOF:
        return parseFailed("Parse error - reached end of file", 0);
        break;
      case '\0':
        return parseFailed("Parse error - reached end of string", 0);
        break;
      default:
        return parseFailed("Parse error unrecognised character", *string);
        break;
    }
  }
}

ExcelValue parseValue(char *string) {
  return parse(&string);
}
